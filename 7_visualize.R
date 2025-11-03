# ============================================================================
# 7_visualize.R - 視覺化分析
# ============================================================================

library(data.table)
library(ggplot2)
library(scales)
library(readxl)

# ---- 載入資料並取得 h 值 ----
cat("載入資料...\n")
# 用目前的 h 找資料夾（平行跑才不會亂）
if (!exists("h")) stop("沒有設定 h，請在外面給 h 再跑這支 script。")

results_dir <- file.path("outputs", paste0("h_", h), "results")
if (!dir.exists(results_dir)) stop(paste("找不到", results_dir, "，請先跑對應的 0_setup.R。"))
setup <- readRDS(file.path(results_dir, "prepared_data.rds"))
h <- setup$h
output_root <- setup$output_root

# ---- 設定輸出目錄 ----
fig_dir <- file.path(output_root, "figures")
dir.create(fig_dir, recursive = TRUE, showWarnings = FALSE)

today <- format(Sys.Date(), "%Y-%m-%d")

cat(sprintf("\n視覺化分析開始 (h=%d)...\n", h))
cat("輸出目錄:", fig_dir, "\n\n")

# ============================================================================
# Part 1: 樣本外預測視覺化
# ============================================================================

cat("[1/8] 載入樣本外預測資料...\n")

excel_dir <- file.path(output_root, "excel_files")
metrics_files <- list.files(excel_dir, pattern="^forecast_outputs_.*\\.xlsx$", 
                           full.names=TRUE)

if(length(metrics_files) == 0) {
  cat("警告：找不到樣本外預測結果，跳過 Part 1\n\n")
  skip_p6 <- TRUE
} else {
  latest_metrics <- metrics_files[which.max(file.info(metrics_files)$mtime)]
  metrics <- as.data.table(read_excel(latest_metrics, sheet = "oos_predictions_by_model"))
  score <- as.data.table(read_excel(latest_metrics, sheet = "model_metrics_rmse_mae_mad"))
  skip_p6 <- FALSE
  cat("  已載入:", latest_metrics, "\n")
}

if(!skip_p6){
  
  # ---- 圖 1-3: 時間序列走勢圖（分三個家族）----
  cat("[2/8] 繪製樣本外預測走勢圖...\n")
  
  ts_models <- c("RandomWalk", "AR", "ARMA")
  lasso_models <- c("Lasso", "Ridge", "ElasticNet", "AdaLasso", "AdaElasticNet", "PCR")
  ml_models <- c("BaggingTree", "RandomForest")
  
  # 時間序列家族
  if(any(ts_models %in% metrics$model)){
    p1 <- ggplot(metrics[model %in% ts_models], aes(x=date)) +
      geom_line(aes(y=y, color="實際值"), linewidth=0.8) +
      geom_line(aes(y=yhat, color=model), linewidth=0.6, alpha=0.8) +
      scale_color_manual(values = c("實際值"="black", "RandomWalk"="#E41A1C", 
                                    "AR"="#377EB8", "ARMA"="#4DAF4A")) +
      scale_x_date(date_labels="%Y-%m") +
      labs(title=sprintf("樣本外預測：時間序列模型 (h=%d)", h),
           subtitle=sprintf("產出日期：%s", today),
           x="日期", y="YoY", color="模型") +
      theme_minimal(base_size=12) +
      theme(legend.position="bottom")
    
    ggsave(file.path(fig_dir, sprintf("oos_timeseries_h%d_%s.png", h, gsub("-","",today))), 
           p1, width=12, height=6, dpi=300)
    cat("  已儲存: oos_timeseries.png\n")
  }
  
  # Lasso 家族
  if(any(lasso_models %in% metrics$model)){
    p2 <- ggplot(metrics[model %in% lasso_models], aes(x=date)) +
      geom_line(aes(y=y, color="實際值"), linewidth=0.8) +
      geom_line(aes(y=yhat, color=model), linewidth=0.6, alpha=0.7) +
      scale_color_manual(values = c("實際值"="black", 
                                    "Lasso"="#E41A1C", "Ridge"="#377EB8", 
                                    "ElasticNet"="#4DAF4A", "AdaLasso"="#984EA3",
                                    "AdaElasticNet"="#FF7F00", "PCR"="#A65628")) +
      scale_x_date(date_labels="%Y-%m") +
      labs(title=sprintf("樣本外預測：Lasso 家族模型 (h=%d)", h),
           subtitle=sprintf("產出日期：%s", today),
           x="日期", y="YoY", color="模型") +
      theme_minimal(base_size=12) +
      theme(legend.position="bottom")
    
    ggsave(file.path(fig_dir, sprintf("oos_lasso_h%d_%s.png", h, gsub("-","",today))), 
           p2, width=12, height=6, dpi=300)
    cat("  已儲存: oos_lasso.png\n")
  }
  
  # 機器學習家族
  if(any(ml_models %in% metrics$model)){
    p3 <- ggplot(metrics[model %in% ml_models], aes(x=date)) +
      geom_line(aes(y=y, color="實際值"), linewidth=0.8) +
      geom_line(aes(y=yhat, color=model), linewidth=0.6, alpha=0.8) +
      scale_color_manual(values = c("實際值"="black", 
                                    "BaggingTree"="#E41A1C", 
                                    "RandomForest"="#377EB8")) +
      scale_x_date(date_labels="%Y-%m") +
      labs(title=sprintf("樣本外預測：機器學習模型 (h=%d)", h),
           subtitle=sprintf("產出日期：%s", today),
           x="日期", y="YoY", color="模型") +
      theme_minimal(base_size=12) +
      theme(legend.position="bottom")
    
    ggsave(file.path(fig_dir, sprintf("oos_ml_h%d_%s.png", h, gsub("-","",today))), 
           p3, width=12, height=6, dpi=300)
    cat("  已儲存: oos_ml.png\n")
  }
  
  # ---- 圖 4: 誤差分布圖 ----
  cat("[3/8] 繪製誤差分布圖...\n")
  
  metrics[, error := y - yhat]
  
  p4 <- ggplot(metrics, aes(x=reorder(model, error, FUN=median), y=error)) +
    geom_boxplot(fill="skyblue", alpha=0.7) +
    geom_hline(yintercept=0, linetype="dashed", color="red") +
    coord_flip() +
    labs(title=sprintf("各模型預測誤差分布 (h=%d)", h),
         subtitle=sprintf("產出日期：%s", today),
         x="模型", y="預測誤差 (實際值 - 預測值)") +
    theme_minimal(base_size=12)
  
  ggsave(file.path(fig_dir, sprintf("error_distribution_h%d_%s.png", h, gsub("-","",today))), 
         p4, width=10, height=8, dpi=300)
  cat("  已儲存: error_distribution.png\n")
  
  # ---- 圖 5: RMSE 排名圖 ----
  cat("[4/8] 繪製模型排名圖...\n")
  
  score_sorted <- score[order(RMSE)]
  score_sorted[, rank := .I]
  
  p5 <- ggplot(score_sorted, aes(x=reorder(model, -RMSE), y=RMSE)) +
    geom_col(fill="steelblue", alpha=0.8) +
    geom_text(aes(label=sprintf("%.4f", RMSE)), hjust=-0.2, size=3.5) +
    coord_flip() +
    labs(title=sprintf("模型表現排名 (h=%d)", h),
         subtitle=sprintf("產出日期：%s｜依 RMSE 排序（越小越好）", today),
         x="模型", y="RMSE") +
    theme_minimal(base_size=12) +
    theme(panel.grid.major.y = element_blank())
  
  ggsave(file.path(fig_dir, sprintf("model_ranking_h%d_%s.png", h, gsub("-","",today))), 
         p5, width=10, height=8, dpi=300)
  cat("  已儲存: model_ranking.png\n")
  
} else {
  cat("跳過圖 1-5（樣本外預測視覺化）\n")
}

# ============================================================================
# Part 2: 未來預測視覺化
# ============================================================================

cat("\n[5/8] 載入未來預測資料...\n")

nowcast_files <- list.files(excel_dir, pattern="^nowcast_.*\\.xlsx$", full.names=TRUE)

if(length(nowcast_files) == 0) {
  cat("警告：找不到 nowcast 結果，跳過 Part 2\n\n")
  skip_p8 <- TRUE
} else {
  latest_nowcast <- nowcast_files[which.max(file.info(nowcast_files)$mtime)]
  nowcast <- as.data.table(read_excel(latest_nowcast))
  skip_p8 <- FALSE
  cat("  已載入:", latest_nowcast, "\n")
}

if(!skip_p8){
  
  # ---- 圖 6: 未來預測路徑圖（YoY）----
  cat("[6/8] 繪製未來預測路徑圖...\n")
  
  nowcast_summary <- nowcast[, .(
    mean = mean(yhat_yoy, na.rm=TRUE),
    min = min(yhat_yoy, na.rm=TRUE),
    max = max(yhat_yoy, na.rm=TRUE),
    q25 = quantile(yhat_yoy, 0.25, na.rm=TRUE),
    q75 = quantile(yhat_yoy, 0.75, na.rm=TRUE)
  ), by=.(date, horizon)]
  
  p6 <- ggplot(nowcast_summary, aes(x=date)) +
    geom_ribbon(aes(ymin=min, ymax=max), fill="lightblue", alpha=0.3) +
    geom_ribbon(aes(ymin=q25, ymax=q75), fill="steelblue", alpha=0.4) +
    geom_line(aes(y=mean), color="darkblue", linewidth=1) +
    geom_point(aes(y=mean), color="darkblue", size=2) +
    scale_x_date(date_labels="%Y-%m") +
    labs(title=sprintf("未來 12 期預測路徑 - YoY (h=%d)", h),
         subtitle=sprintf("產出日期：%s｜深藍線=平均，深藍區=25-75%%分位，淺藍區=最小-最大", today),
         x="日期", y="YoY 預測") +
    theme_minimal(base_size=12)
  
  ggsave(file.path(fig_dir, sprintf("forecast_paths_yoy_h%d_%s.png", h, gsub("-","",today))), 
         p6, width=12, height=6, dpi=300)
  cat("  已儲存: forecast_paths.png\n")
  
  # ---- 圖 7: 實際稅收預測圖 ----
  cat("[7/8] 繪製實際稅收預測圖...\n")
  
  nowcast_actual <- nowcast[, .(
    mean = mean(actual_forecast, na.rm=TRUE),
    min = min(actual_forecast, na.rm=TRUE),
    max = max(actual_forecast, na.rm=TRUE),
    q25 = quantile(actual_forecast, 0.25, na.rm=TRUE),
    q75 = quantile(actual_forecast, 0.75, na.rm=TRUE)
  ), by=.(date, horizon)]
  
  p7 <- ggplot(nowcast_actual, aes(x=date)) +
    geom_ribbon(aes(ymin=min, ymax=max), fill="lightgreen", alpha=0.3) +
    geom_ribbon(aes(ymin=q25, ymax=q75), fill="forestgreen", alpha=0.4) +
    geom_line(aes(y=mean), color="darkgreen", linewidth=1) +
    geom_point(aes(y=mean), color="darkgreen", size=2) +
    scale_x_date(date_labels="%Y-%m") +
    scale_y_continuous(labels=comma) +
    labs(title=sprintf("未來 12 期實際稅收預測 (h=%d)", h),
         subtitle=sprintf("產出日期：%s｜深綠線=平均，深綠區=25-75%%分位，淺綠區=最小-最大", today),
         x="日期", y="實際稅收預測值") +
    theme_minimal(base_size=12)
  
  ggsave(file.path(fig_dir, sprintf("forecast_actual_h%d_%s.png", h, gsub("-","",today))), 
         p7, width=12, height=6, dpi=300)
  cat("  已儲存: forecast_actual.png\n")
  
} else {
  cat("跳過圖 6-7（未來預測視覺化）\n")
}

# ============================================================================
# Part 3: 變數重要性視覺化
# ============================================================================

cat("\n[8/8] 載入變數重要性資料...\n")

imp_files <- list.files(excel_dir, pattern="^var_importance_BASE.*\\.xlsx$", 
                       full.names=TRUE)

if(length(imp_files) == 0) {
  cat("警告：找不到變數重要性結果，跳過 Part 3\n\n")
  skip_p9 <- TRUE
} else {
  latest_imp <- imp_files[which.max(file.info(imp_files)$mtime)]
  imp_short <- as.data.table(read_excel(latest_imp, sheet="BASE_top_SHORT"))
  skip_p9 <- FALSE
  cat("  已載入:", latest_imp, "\n")
}

if(!skip_p9){
  
  # ---- 圖 8: 重要變數排名圖（只看前5名）----
  cat("[8/8] 繪製變數重要性圖...\n")
  
  model_cols <- setdiff(names(imp_short), "排序")
  imp_long <- melt(imp_short[1:5], id.vars="排序", 
                   measure.vars=model_cols,
                   variable.name="model", value.name="variable")
  imp_long[, 排序 := as.integer(排序)]
  imp_long <- imp_long[!is.na(variable)]
  
  p8 <- ggplot(imp_long, aes(x=排序, y=reorder(variable, -排序), 
                             color=model, group=model)) +
    geom_point(size=3, alpha=0.7) +
    geom_line(alpha=0.5) +
    scale_x_continuous(breaks=1:5, trans="reverse") +
    labs(title=sprintf("各模型前 5 名重要變數 (h=%d)", h),
         subtitle=sprintf("產出日期：%s", today),
         x="排名", y="變數", color="模型") +
    theme_minimal(base_size=12) +
    theme(legend.position="bottom")
  
  ggsave(file.path(fig_dir, sprintf("importance_ranking_h%d_%s.png", h, gsub("-","",today))), 
         p8, width=12, height=8, dpi=300)
  cat("  已儲存: importance_ranking.png\n")
  
  # ---- 圖 9: 重要變數熱力圖（共識度）----
  cat("[8/8+] 繪製變數重要性熱力圖...\n")
  
  all_top_vars <- unlist(imp_short[1:5, ..model_cols])
  var_freq <- as.data.table(table(all_top_vars))
  setnames(var_freq, c("variable", "count"))
  var_freq <- var_freq[order(-count)]
  
  top_vars <- head(var_freq$variable, 10)
  heat_matrix <- matrix(0, nrow=length(top_vars), ncol=length(model_cols))
  rownames(heat_matrix) <- top_vars
  colnames(heat_matrix) <- model_cols
  
  for(i in seq_along(model_cols)){
    mc <- model_cols[i]
    vars <- imp_short[1:5, get(mc)]
    for(v in vars){
      if(v %in% top_vars){
        heat_matrix[v, mc] <- which(imp_short[[mc]] == v)[1]
      }
    }
  }
  
  heat_dt <- as.data.table(melt(heat_matrix))
  setnames(heat_dt, c("variable", "model", "rank"))
  heat_dt[rank == 0, rank := NA]
  
  p9 <- ggplot(heat_dt, aes(x=model, y=variable, fill=rank)) +
    geom_tile(color="white", linewidth=0.5) +
    geom_text(aes(label=ifelse(!is.na(rank), rank, "")), 
              color="white", size=4, fontface="bold") +
    scale_fill_gradient(low="darkred", high="lightyellow", 
                        na.value="grey90", name="排名", breaks=1:5) +
    labs(title=sprintf("變數重要性共識度熱力圖 (h=%d)", h),
         subtitle=sprintf("產出日期：%s｜數字=排名，灰色=未進前5", today),
         x="模型", y="變數") +
    theme_minimal(base_size=12) +
    theme(axis.text.x = element_text(angle=45, hjust=1),
          panel.grid = element_blank())
  
  ggsave(file.path(fig_dir, sprintf("importance_heatmap_h%d_%s.png", h, gsub("-","",today))), 
         p9, width=12, height=8, dpi=300)
  cat("  已儲存: importance_heatmap.png\n")
  
} else {
  cat("跳過圖 8-9（變數重要性視覺化）\n")
}

# ============================================================================
# 完成
# ============================================================================

cat("\n=====================================================\n")
cat(sprintf("  視覺化分析完成！(h=%d)\n", h))
cat("=====================================================\n")
cat(sprintf("\n所有圖表已儲存至:\n  %s\n\n", fig_dir))

all_plots <- list.files(fig_dir, pattern="\\.png$")
if(length(all_plots) > 0){
  cat("產生的圖表:\n")
  for(plot in all_plots){
    cat(sprintf("  ✓ %s\n", plot))
  }
} else {
  cat("警告：沒有產生任何圖表\n")
}

cat("\n")
