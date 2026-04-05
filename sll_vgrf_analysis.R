library(ggplot2)
library(dplyr)
library(openxlsx)

# --- User Settings ---

input_dir <- "data/"
output_dir <- "output/"
subject_id <- "subject_01"
body_mass_kg <- 70
landing_leg <- "R"  # "R" = right force plate, "L" = left force plate

# --- Constants ---

N_NORM_POINTS <- 101
GRAVITY <- 9.81

FZ_COL_LEFT  <- 5
FZ_COL_RIGHT <- 14

# --- Functions ---

parse_csv_file <- function(filepath, fz_column) {
  raw_lines <- readLines(filepath, warn = FALSE)
  
  events <- data.frame(name = character(), time = numeric(), stringsAsFactors = FALSE)
  
  for (i in 3:length(raw_lines)) {
    line <- trimws(raw_lines[i])
    if (grepl("^Devices", line, ignore.case = TRUE)) break
    if (line == "" || grepl("^,+$", line)) next
    
    parts <- strsplit(line, ",")[[1]]
    if (length(parts) < 4) next
    
    evt_name <- trimws(parts[3])
    time_val <- suppressWarnings(as.numeric(trimws(parts[4])))
    if (is.na(time_val)) next
    
    if (grepl("strike|off", evt_name, ignore.case = TRUE)) {
      events <- rbind(events, data.frame(name = evt_name, time = time_val, stringsAsFactors = FALSE))
    }
  }
  
  events <- events[order(events$time), ]
  
  if (nrow(events) < 2) {
    warning(paste("Less than 2 events found in", basename(filepath)))
    return(NULL)
  }
  
  fs_time <- events$time[1]
  fo_time <- events$time[2]
  
  devices_line <- which(grepl("^Devices", raw_lines, ignore.case = TRUE))[1]
  if (is.na(devices_line)) {
    warning(paste("No Devices section found in", basename(filepath)))
    return(NULL)
  }
  
  analog_freq <- as.numeric(strsplit(raw_lines[devices_line + 1], ",")[[1]][1])
  force_start <- devices_line + 5
  
  force_end <- length(raw_lines)
  for (i in force_start:length(raw_lines)) {
    line <- trimws(raw_lines[i])
    if (line == "" || grepl("^,+$", line)) {
      force_end <- i - 1
      break
    }
    first_val <- suppressWarnings(as.numeric(strsplit(line, ",")[[1]][1]))
    if (is.na(first_val)) {
      force_end <- i - 1
      break
    }
  }
  
  n_force_rows <- force_end - force_start + 1
  
  fz_data <- numeric(n_force_rows)
  for (i in 1:n_force_rows) {
    parts <- strsplit(raw_lines[force_start + i - 1], ",")[[1]]
    if (length(parts) >= fz_column) {
      fz_data[i] <- as.numeric(parts[fz_column])
    } else {
      fz_data[i] <- NA
    }
  }
  
  time_vec <- (0:(n_force_rows - 1)) / analog_freq
  
  large_vals <- fz_data[abs(fz_data) > 50]
  if (length(large_vals) > 0 && mean(large_vals, na.rm = TRUE) < 0) {
    fz_data <- -fz_data
  }
  
  return(list(
    fs_time = fs_time,
    fo_time = fo_time,
    fz = fz_data,
    time = time_vec,
    freq = analog_freq
  ))
}

time_to_sample <- function(t, freq, n_samples) {
  idx <- round(t * freq) + 1
  return(max(1, min(idx, n_samples)))
}

# --- Main Processing ---

cat(strrep("=", 60), "\n")
cat("Single Leg Landing vGRF Analysis\n")
cat(strrep("=", 60), "\n\n")

if (!dir.exists(input_dir)) {
  stop(paste("Input directory not found:", input_dir))
}

if (!dir.exists(output_dir)) {
  dir.create(output_dir, recursive = TRUE)
  cat("Created output directory:", output_dir, "\n")
}

fz_column <- ifelse(toupper(landing_leg) == "R", FZ_COL_RIGHT, FZ_COL_LEFT)
cat("Landing leg:", toupper(landing_leg), "\n")
cat("Using force plate:", ifelse(toupper(landing_leg) == "R", "Right (col 14)", "Left (col 5)"), "\n")
cat("Subject:", subject_id, "\n")
cat("Body mass:", body_mass_kg, "kg\n\n")

csv_files <- list.files(input_dir, pattern = "\\.csv$", full.names = TRUE, ignore.case = TRUE)
cat("Found", length(csv_files), "CSV file(s) in input directory\n")

if (length(csv_files) == 0) {
  stop("No CSV files found in input directory")
}

# --- Process Trials ---

cat("\nProcessing trials...\n")
cat(strrep("-", 40), "\n")

trial_results <- list()
valid_trial_count <- 0

for (csv_file in csv_files) {
  trial_name <- tools::file_path_sans_ext(basename(csv_file))
  cat("  ", trial_name, "... ")
  
  parsed <- tryCatch({
    parse_csv_file(csv_file, fz_column)
  }, error = function(e) {
    cat("ERROR:", conditionMessage(e), "\n")
    return(NULL)
  })
  
  if (is.null(parsed)) {
    cat("SKIPPED\n")
    next
  }
  
  n_samples <- length(parsed$fz)
  fs_idx <- time_to_sample(parsed$fs_time, parsed$freq, n_samples)
  fo_idx <- time_to_sample(parsed$fo_time, parsed$freq, n_samples)
  
  if (fo_idx <= fs_idx) {
    cat("SKIPPED (invalid event order)\n")
    next
  }
  
  landing_fz <- parsed$fz[fs_idx:fo_idx]
  landing_time <- parsed$time[fs_idx:fo_idx] - parsed$time[fs_idx]
  landing_len <- length(landing_fz)
  
  if (landing_len < 50) {
    cat("SKIPPED (landing phase too short:", landing_len, "samples)\n")
    next
  }
  
  peak_idx <- which.max(landing_fz)
  peak_force_N <- landing_fz[peak_idx]
  peak_time <- landing_time[peak_idx]
  
  landing_fz_nkg <- landing_fz / body_mass_kg
  peak_force_nkg <- peak_force_N / body_mass_kg
  
  # Loading rate: 20-80% of rise to peak
  idx_20 <- max(1, round(1 + (peak_idx - 1) * 0.20))
  idx_80 <- min(peak_idx, round(1 + (peak_idx - 1) * 0.80))
  
  force_20 <- landing_fz[idx_20]
  force_80 <- landing_fz[idx_80]
  time_20 <- landing_time[idx_20]
  time_80 <- landing_time[idx_80]
  
  if ((time_80 - time_20) > 0) {
    lr_N_per_s <- (force_80 - force_20) / (time_80 - time_20)
    lr_BW_per_s <- lr_N_per_s / (body_mass_kg * GRAVITY)
  } else {
    lr_BW_per_s <- NA
  }
  
  # Time-normalize to 101 points
  pct_in <- seq(0, 100, length.out = landing_len)
  pct_out <- 0:100
  fz_norm <- approx(pct_in, landing_fz_nkg, xout = pct_out, method = "linear")$y
  
  valid_trial_count <- valid_trial_count + 1
  trial_results[[trial_name]] <- list(
    fz_norm = fz_norm,
    peak_force_nkg = peak_force_nkg,
    peak_time = peak_time,
    peak_pct = (peak_idx - 1) / (landing_len - 1) * 100,
    loading_rate = lr_BW_per_s,
    freq = parsed$freq,
    landing_len = landing_len
  )
  
  cat("OK (peak:", round(peak_force_nkg, 1), "N/kg, LR:", round(lr_BW_per_s, 1), "BW/s)\n")
}

cat(strrep("-", 40), "\n")
cat("Valid trials:", valid_trial_count, "of", length(csv_files), "\n\n")

if (valid_trial_count == 0) {
  stop("No valid trials found")
}

# --- Mean Waveform ---

cat("Calculating mean waveform across", valid_trial_count, "trials...\n")

fz_matrix <- do.call(rbind, lapply(trial_results, function(x) x$fz_norm))

mean_fz <- colMeans(fz_matrix, na.rm = TRUE)
sd_fz <- apply(fz_matrix, 2, sd, na.rm = TRUE)

mean_df <- data.frame(
  percent_landing = 0:100,
  mean_vGRF_Nkg = mean_fz,
  sd_vGRF_Nkg = sd_fz
)

# --- Loading Rate from Mean Waveform ---

cat("Calculating loading rate from mean waveform...\n")

mean_peak_idx <- which.max(mean_fz)
mean_peak_pct <- mean_df$percent_landing[mean_peak_idx]
mean_peak_force <- mean_fz[mean_peak_idx]

pct_20 <- mean_peak_pct * 0.20
pct_80 <- mean_peak_pct * 0.80

force_20_mean <- approx(mean_df$percent_landing, mean_fz, xout = pct_20)$y
force_80_mean <- approx(mean_df$percent_landing, mean_fz, xout = pct_80)$y

avg_freq <- mean(sapply(trial_results, function(x) x$freq))
avg_landing_len <- mean(sapply(trial_results, function(x) x$landing_len))
avg_landing_duration <- avg_landing_len / avg_freq

time_20_mean <- (pct_20 / 100) * avg_landing_duration
time_80_mean <- (pct_80 / 100) * avg_landing_duration

delta_force_N <- (force_80_mean - force_20_mean) * body_mass_kg
delta_time <- time_80_mean - time_20_mean

if (delta_time > 0) {
  mean_lr_N_per_s <- delta_force_N / delta_time
  mean_lr_BW_per_s <- mean_lr_N_per_s / (body_mass_kg * GRAVITY)
} else {
  mean_lr_BW_per_s <- NA
}

cat("  Peak force (mean):", round(mean_peak_force, 2), "N/kg at", round(mean_peak_pct, 1), "%\n")
cat("  Loading rate (mean waveform):", round(mean_lr_BW_per_s, 2), "BW/s\n")
cat("  20% point:", round(pct_20, 1), "% (", round(force_20_mean, 2), "N/kg)\n")
cat("  80% point:", round(pct_80, 1), "% (", round(force_80_mean, 2), "N/kg)\n")

lr_df <- data.frame(
  metric = c("loading_rate_BW_s", "peak_force_N_kg", "peak_percent", 
             "pct_20_force_N_kg", "pct_80_force_N_kg", "pct_20", "pct_80",
             "n_trials", "avg_landing_duration_s", "sampling_freq_Hz"),
  value = c(mean_lr_BW_per_s, mean_peak_force, mean_peak_pct,
            force_20_mean, force_80_mean, pct_20, pct_80,
            valid_trial_count, avg_landing_duration, avg_freq)
)

# --- Plot ---

cat("\nGenerating plot...\n")

trial_curves <- data.frame()
for (trial_name in names(trial_results)) {
  trial_curves <- rbind(trial_curves, data.frame(
    trial = trial_name,
    percent = 0:100,
    force = trial_results[[trial_name]]$fz_norm
  ))
}

y_max <- max(mean_fz + sd_fz, na.rm = TRUE) * 1.1

p <- ggplot() +
  geom_line(data = trial_curves,
            aes(x = percent, y = force, group = trial),
            color = "gray70", alpha = 0.5, linewidth = 0.5) +
  geom_ribbon(data = mean_df,
              aes(x = percent_landing, 
                  ymin = mean_vGRF_Nkg - sd_vGRF_Nkg,
                  ymax = mean_vGRF_Nkg + sd_vGRF_Nkg),
              fill = "#3366CC", alpha = 0.2) +
  geom_line(data = mean_df,
            aes(x = percent_landing, y = mean_vGRF_Nkg),
            color = "#3366CC", linewidth = 1.5) +
  geom_point(aes(x = mean_peak_pct, y = mean_peak_force),
             shape = 24, size = 4, fill = "#CC3333", color = "black", stroke = 1) +
  geom_segment(aes(x = pct_20, y = force_20_mean, xend = pct_80, yend = force_80_mean),
               color = "#009933", linewidth = 1.5, linetype = "solid") +
  geom_point(aes(x = pct_20, y = force_20_mean),
             shape = 21, size = 3, fill = "#009933", color = "black") +
  geom_point(aes(x = pct_80, y = force_80_mean),
             shape = 21, size = 3, fill = "#009933", color = "black") +
  annotate("label", x = 95, y = y_max * 0.98,
           label = sprintf("Peak: %.1f N/kg\nLR: %.1f BW/s", 
                           mean_peak_force, mean_lr_BW_per_s),
           hjust = 1, vjust = 1, size = 3.5, fontface = "bold",
           fill = "white", alpha = 0.85, label.size = 0.5) +
  labs(
    title = sprintf("Single Leg Landing - %s (%s Leg)", subject_id, toupper(landing_leg)),
    subtitle = sprintf("Mean of %d trials | Body mass: %.1f kg", valid_trial_count, body_mass_kg),
    x = "Landing Phase (%)",
    y = "vGRF (N/kg)"
  ) +
  scale_x_continuous(breaks = seq(0, 100, 20), limits = c(0, 105)) +
  scale_y_continuous(expand = expansion(mult = c(0.02, 0.1))) +
  theme_minimal(base_size = 12) +
  theme(
    plot.title = element_text(face = "bold", size = 14),
    plot.subtitle = element_text(color = "gray40"),
    panel.grid.minor = element_blank(),
    axis.title = element_text(face = "bold")
  )

plot_file <- file.path(output_dir, paste0(subject_id, "_", toupper(landing_leg), "_sll_vGRF_plot.png"))
ggsave(plot_file, plot = p, width = 10, height = 6, dpi = 300)
cat("  Plot saved:", plot_file, "\n")

# --- Export to Excel ---

cat("\nExporting to Excel...\n")

output_file <- file.path(output_dir, paste0(subject_id, "_", toupper(landing_leg), "_sll_vGRF_results.xlsx"))

wb <- createWorkbook()

addWorksheet(wb, "mean_timeseries")
writeData(wb, "mean_timeseries", mean_df)

header_style <- createStyle(textDecoration = "bold", halign = "center")
addStyle(wb, "mean_timeseries", header_style, rows = 1, cols = 1:3, gridExpand = TRUE)
setColWidths(wb, "mean_timeseries", cols = 1:3, widths = 18)

num_style <- createStyle(numFmt = "0.0000")
addStyle(wb, "mean_timeseries", num_style, rows = 2:102, cols = 2:3, gridExpand = TRUE)

addWorksheet(wb, "loading_rate")
writeData(wb, "loading_rate", lr_df)

addStyle(wb, "loading_rate", header_style, rows = 1, cols = 1:2, gridExpand = TRUE)
setColWidths(wb, "loading_rate", cols = 1:2, widths = c(25, 15))

saveWorkbook(wb, output_file, overwrite = TRUE)
cat("  Excel saved:", output_file, "\n")

# --- Summary ---

cat("\n", strrep("=", 60), "\n")
cat("ANALYSIS COMPLETE\n")
cat(strrep("=", 60), "\n")
cat("  Subject:", subject_id, "\n")
cat("  Landing leg:", toupper(landing_leg), "\n")
cat("  Body mass:", body_mass_kg, "kg\n")
cat("  Valid trials:", valid_trial_count, "\n")
cat("\n  Results:\n")
cat("    Peak force (mean):", round(mean_peak_force, 2), "N/kg\n")
cat("    Loading rate:     ", round(mean_lr_BW_per_s, 2), "BW/s\n")
cat("\n  Output files:\n")
cat("    Excel:", output_file, "\n")
cat("    Plot: ", plot_file, "\n")
cat(strrep("=", 60), "\n")
