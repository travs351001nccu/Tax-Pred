# ============================================================================
# run_all.R - 主控腳本：選擇要執行的模組
# ============================================================================

# ---- 設定要執行的模組 ----
run_modules <- list(
  setup      = TRUE,   # 資料準備（必須先跑）
  ts         = TRUE,   # 時間序列模型
  lasso      = TRUE,   # Lasso 家族
  ml         = TRUE,   # 機器學習
  combine    = TRUE,   # 合併結果
  nowcast    = TRUE,   # 未來預測
  importance = TRUE,   # 變數重要性
  visualize  = TRUE    # 視覺化分析
)

# ============================================================================

cat("=====================================================\n")
cat("  稅收預測系統 - 主控腳本\n")
cat("=====================================================\n\n")

# ---- 0. 資料準備 ----
if(run_modules$setup) {
  cat("[1/6] 執行資料準備...\n")
  cat("-----------------------------------------------------\n")
  #source("0_setup.R")
  source("0_setup.R", local = TRUE)
  cat("\n✓ 資料準備完成\n\n")
} else {
  cat("[1/6] 跳過資料準備（請確保 results/prepared_data.rds 已存在）\n\n")
}

# ---- 1. 時間序列模型 ----
if(run_modules$ts) {
  cat("[2/6] 執行時間序列模型...\n")
  cat("-----------------------------------------------------\n")
  source("1_models_ts.R")
  cat("\n✓ 時間序列模型完成\n\n")
} else {
  cat("[2/6] 跳過時間序列模型\n\n")
}

# ---- 2. Lasso 家族 ----
if(run_modules$lasso) {
  cat("[3/6] 執行 Lasso 家族模型...\n")
  cat("-----------------------------------------------------\n")
  source("2_models_lasso.R")
  cat("\n✓ Lasso 家族模型完成\n\n")
} else {
  cat("[3/6] 跳過 Lasso 家族模型\n\n")
}

# ---- 3. 機器學習 ----
if(run_modules$ml) {
  cat("[4/6] 執行機器學習模型...\n")
  cat("-----------------------------------------------------\n")
  source("3_models_ml.R")
  cat("\n✓ 機器學習模型完成\n\n")
} else {
  cat("[4/6] 跳過機器學習模型\n\n")
}

# ---- 4. 合併結果 ----
if(run_modules$combine) {
  cat("[5/6] 合併所有結果...\n")
  cat("-----------------------------------------------------\n")
  source("4_combine_results.R")
  cat("\n✓ 結果合併完成\n\n")
} else {
  cat("[5/6] 跳過結果合併\n\n")
}

# ---- 5. Nowcast ----
if(run_modules$nowcast) {
  cat("[6/7] 執行 Nowcast...\n")
  cat("-----------------------------------------------------\n")
  source("5_nowcast.R")
  cat("\n✓ Nowcast 完成\n\n")
} else {
  cat("[6/7] 跳過 Nowcast\n\n")
}

# ---- 6. 變數重要性 ----
if(run_modules$importance) {
  cat("[7/8] 計算變數重要性...\n")
  cat("-----------------------------------------------------\n")
  source("6_variable_importance.R")
  cat("\n✓ 變數重要性分析完成\n\n")
} else {
  cat("[7/8] 跳過變數重要性分析\n\n")
}

# ---- 7. 視覺化 ----
if(run_modules$visualize) {
  cat("[8/8] 執行視覺化分析...\n")
  cat("-----------------------------------------------------\n")
  source("7_visualize.R")
  cat("\n✓ 視覺化分析完成\n\n")
} else {
  cat("[8/8] 跳過視覺化分析\n\n")
}

# ---- 完成 ----
cat("=====================================================\n")
cat("  所有選定模組執行完畢！\n")
cat("=====================================================\n")

# 顯示執行時間
cat("\n執行模組:\n")
cat(sprintf("  資料準備: %s\n", ifelse(run_modules$setup, "✓", "✗")))
cat(sprintf("  時間序列: %s\n", ifelse(run_modules$ts, "✓", "✗")))
cat(sprintf("  Lasso家族: %s\n", ifelse(run_modules$lasso, "✓", "✗")))
cat(sprintf("  機器學習: %s\n", ifelse(run_modules$ml, "✓", "✗")))
cat(sprintf("  合併結果: %s\n", ifelse(run_modules$combine, "✓", "✗")))
cat(sprintf("  Nowcast: %s\n", ifelse(run_modules$nowcast, "✓", "✗")))
cat(sprintf("  變數重要性: %s\n", ifelse(run_modules$importance, "✓", "✗")))
cat(sprintf("  視覺化分析: %s\n", ifelse(run_modules$visualize, "✓", "✗")))