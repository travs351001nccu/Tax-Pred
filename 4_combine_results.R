# ============================================================================
# 4_combine_results.R - 合併所有模型結果並輸出
# ============================================================================

library(data.table)
library(writexl)

# ---- 載入所有結果 ----
cat("載入模型結果...\n")

# 用目前的 h 找資料夾（平行跑才不會亂）
if (!exists("h")) stop("沒有設定 h，請在外面給 h 再跑這支 script。")

results_dir <- file.path("outputs", paste0("h_", h), "results")
if (!dir.exists(results_dir)) stop(paste("找不到", results_dir, "，請先跑對應的 0_setup.R。"))

# 讀取 h 值
setup <- readRDS(file.path(results_dir, "prepared_data.rds"))
h <- setup$h
cat(sprintf("使用 h = %d 的資料\n", h))

result_files <- c(
  file.path(results_dir, "roll_ts.rds"),
  file.path(results_dir, "roll_lasso.rds"),
  file.path(results_dir, "roll_ml.rds")
)

# 檢查哪些檔案存在
existing_files <- result_files[file.exists(result_files)]
if(length(existing_files) == 0) {
  stop("沒有找到任何模型結果檔案！請先執行 1_models_ts.R, 2_models_lasso.R, 或 3_models_ml.R")
}

cat("找到以下結果檔案:\n")
print(existing_files)

# 載入並合併
all_rolls <- list()
all_scores <- list()

if(file.path(results_dir, "roll_ts.rds") %in% existing_files){
  res_ts <- readRDS(file.path(results_dir, "roll_ts.rds"))
  all_rolls$roll_ts <- res_ts$roll_ts
  all_scores$score_ts <- res_ts$score_ts
  cat("✓ 時間序列模型\n")
}

if(file.path(results_dir, "roll_lasso.rds") %in% existing_files){
  res_lasso <- readRDS(file.path(results_dir, "roll_lasso.rds"))
  all_rolls$roll_lasso <- res_lasso$roll_lasso
  all_scores$score_lasso <- res_lasso$score_lasso
  cat("✓ Lasso 家族模型\n")
}

if(file.path(results_dir, "roll_ml.rds") %in% existing_files){
  res_ml <- readRDS(file.path(results_dir, "roll_ml.rds"))
  all_rolls$roll_ml <- res_ml$roll_ml
  all_scores$score_ml <- res_ml$score_ml
  cat("✓ 機器學習模型\n")
}

# ---- 合併所有預測 ----
cat("\n合併所有預測結果...\n")
metrics <- rbindlist(all_rolls, fill=TRUE)
metrics <- metrics[!is.na(yhat)]

# ---- 合併所有評分 ----
cat("合併所有評分...\n")
score <- rbindlist(all_scores, fill=TRUE)
setorder(score, RMSE)

print(score)

# ---- 輸出到 Excel ----
cat("\n輸出結果...\n")

if (!exists("h")) stop("沒有設定 h，請先設定 h 再執行這段。")

excel_dir <- file.path("outputs", paste0("h_", h), "excel_files")
dir.create(excel_dir, recursive = TRUE, showWarnings = FALSE)

# 檔案名稱加上 h 和日期
today <- format(Sys.Date(), "%Y%m%d")
filename_base <- sprintf("h%d_%s", h, today)

# 分別輸出兩個 xlsx 檔
writexl::write_xlsx(score, 
                    path = file.path(excel_dir, 
                                   sprintf("model_metrics_%s.xlsx", filename_base)))
writexl::write_xlsx(metrics, 
                    path = file.path(excel_dir, 
                                   sprintf("oos_predictions_%s.xlsx", filename_base)))

# 合併兩個 Sheet
writexl::write_xlsx(
  list(
    model_metrics_rmse_mae_mad = score,
    oos_predictions_by_model   = metrics
  ),
  path = file.path(excel_dir, sprintf("forecast_outputs_%s.xlsx", filename_base))
)

cat("\n結果已輸出至:\n")
cat(sprintf("  %s\n", excel_dir))
cat(sprintf("  - forecast_outputs_%s.xlsx\n", filename_base))
cat(sprintf("  - model_metrics_%s.xlsx\n", filename_base))
cat(sprintf("  - oos_predictions_%s.xlsx\n", filename_base))
cat("\n完成！\n")
