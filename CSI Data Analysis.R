library(readxl)
library(dplyr)
library(lubridate)
library(stringr)
library(ggplot2)
library(car)
library(tidyr)
library(corrplot)
library(ggrepel)
library(plotly)
library(tibble)
library(fastDummies)
library(tidyverse)
library(psych)
library(forcats)
library(reticulate)
library(pROC)
library(pheatmap)
library(writexl)
library(scales)

folder_path <- "C:\\Users\\A10132\\Desktop\\楊欣恩\\CS\\CSI Data Analysis\\Database"

# Anchor the column title
key_headers <- c("專案代碼", "車牌號碼", "工單號碼","據點名稱")
detect_header_row <- function(file_path, max_scan = 10) {
  raw_preview <- read_excel(
    path = file_path,
    col_names = FALSE,
    n_max = max_scan,
    col_types = "text"
  )
  
  for (i in seq_len(nrow(raw_preview))) {
    row_values <- raw_preview[i, ] %>%
      unlist(use.names = FALSE) %>%
      as.character() %>%
      str_trim()
    
    hit_n <- sum(key_headers %in% row_values, na.rm = TRUE)
    
    if (hit_n >= 2) {
      return(i)
    }
  }
  
  return(NA_integer_)
}
# 讀單一檔案：先偵測欄名列，再依該列欄名讀進來
read_CSIDATA_by_names <- function(file_path) {
  cat("讀取中：", basename(file_path), "\n")
  
  header_row <- tryCatch(
    detect_header_row(file_path),
    error = function(e) NA_integer_
  )
  
  if (is.na(header_row)) {
    message("❌ 找不到欄名列：", file_path)
    return(NULL)
  }
  df <- tryCatch({
    read_excel(
      path = file_path,
      skip = header_row - 1,
      col_names = TRUE,
      col_types = "text"
    )
  }, error = function(e) {
    message("❌ 讀檔失敗：", file_path)
    return(NULL)
  })
  
  if (is.null(df) || nrow(df) == 0) return(NULL)
  
  # 清理欄名
  names(df) <- names(df) %>%
    str_replace_all("[\r\n\t]", "") %>%
    str_trim()
  
  # 清理內容
  df <- df %>%
    mutate(across(everything(), ~ str_replace_all(.x, "[\r\n\t]", ""))) %>%
    mutate(across(everything(), ~ str_trim(.x)))
  
  # 加來源資訊
  df <- df %>%
    mutate(
      source_file = basename(file_path),
      source_path = file_path
    )
  
  return(df)
}
# 抓全部 Excel
file_list <- list.files(
  path = folder_path,
  pattern = "\\.(xls|xlsx)$",
  full.names = TRUE,
  recursive = TRUE,
  ignore.case = TRUE
)

if (length(file_list) == 0) {
  stop("找不到任何 Excel 檔案，請確認路徑。")
}

# 全部讀取
df_list <- map(file_list, read_CSIDATA_by_names)

# 依欄名自動對齊
df <- bind_rows(df_list)

# 去重
df <- df %>%
  filter(!is.na(工單號碼)) %>%
  mutate(
    填寫問卷時間_parsed = ymd_hms(填寫問卷時間, quiet = TRUE),
    has_response = !is.na(填寫問卷時間_parsed)
  ) %>%
  arrange(工單號碼, desc(has_response), desc(填寫問卷時間_parsed)) %>%
  group_by(工單號碼) %>%
  slice(1) %>%
  ungroup()

cat("\n====================\n")
cat("檔案數：", length(file_list), "\n")
cat("合併後筆數：", nrow(df), "\n")
cat("====================\n")

question_cols <- grep("^[0-9]+-", names(df), value = TRUE)
score_questions <- question_cols[!grepl("原因|改進|建議|類別", question_cols)]
date_cols <- c("簡訊發送", "LINE發送", "點開簡訊時間", "填寫問卷時間", "轉回CSI專案日期")
date_only_cols <- c("入廠日")
num_cols <- c("滿意度總分", "落實度總分","高滿意","低滿意","A+落實", score_questions)
char_cols <- c("服務專員代碼", "據點代碼", "車牌號碼", "工單號碼", "專案代碼")

parse_num_safe <- function(x) {
  x_chr <- as.character(x) %>% str_trim()
  x_chr[x_chr %in% c("", "NA", "NaN", "-", "—")] <- NA
  x_chr <- str_replace_all(x_chr, ",", "")
  x_num <- str_extract(x_chr, "-?\\d+\\.?\\d*")
  suppressWarnings(as.numeric(x_num))
}

parse_dt_safe <- function(x) {
  x_chr <- as.character(x) %>% str_trim()
  x_chr[x_chr %in% c("", "NA", "NaN", "-", "—")] <- NA
  suppressWarnings(
    parse_date_time(
      x_chr,
      orders = c(
        "ymd HMS", "Y/m/d H:M:S", "Y-m-d H:M:S",
        "ymd HM",  "Y/m/d H:M",   "Y-m-d H:M",
        "ymd",     "Y/m/d",       "Y-m-d"
      )
    )
  )
}

parse_date_safe <- function(x) {
  as_date(parse_dt_safe(x))
}

df <- df %>%
  mutate(
    across(any_of(date_cols), parse_dt_safe),
    across(any_of(date_only_cols), parse_date_safe),
    across(any_of(num_cols), parse_num_safe),
    across(any_of(char_cols), as.character)
  )

df <- df %>%
  mutate(
    是否回函 = if_else(!is.na(填寫問卷時間), 1L, 0L),
    月份 = floor_date(入廠日, "month")
  )

check_month <- df %>% count(月份, 是否回函) %>% tidyr::pivot_wider(names_from=是否回函, values_from=n, values_fill=0)
check_agent <- df %>% count(服務專員姓名, 是否回函) %>% tidyr::pivot_wider(names_from=是否回函, values_from=n, values_fill=0)
print(check_month)
print(check_agent)


# 全體
overall <- df %>%
  summarise(發送數 = n(),
            已回函 = sum(是否回函),
            回函率 = 已回函 / 發送數)

by_month <- df %>%
  group_by(月份) %>%
  summarise(發送數 = n(),
            已回函 = sum(是否回函),
            回函率 = 已回函 / 發送數) %>%
  arrange(月份)

valid_months <- check_month %>%
  filter(`0` > 0, `1` > 0) %>%
  pull(月份)

df_m <- df %>% filter(月份 %in% valid_months)

by_SA <- df %>%
  group_by(服務專員姓名) %>%
  summarise(發送數 = n(),
            已回函 = sum(是否回函),
            回函率 = 已回函 / 發送數) %>%
  arrange(desc(發送數))

ggplot(by_month, aes(x = 月份, y = 回函率)) +
  geom_line() + geom_point() +
  scale_y_continuous(labels = scales::percent_format()) +
  labs(title = "各月回函率趨勢", x = "月份", y = "回函率")

top_agents <- by_SA %>% filter(發送數 >= 100)  # 可視需要調門檻

ggplot(top_agents, aes(x = reorder(服務專員姓名, 回函率), y = 回函率)) +
  geom_col() +
  coord_flip() +
  scale_y_continuous(labels = scales::percent_format()) +
  labs(title = "專員回函率（發送數≥100）", x = "服務專員", y = "回函率")

# 只看有回函的樣本做平均滿意度
avg_score_by_month <- df %>%
  filter(是否回函 == 1) %>%
  group_by(月份) %>%
  summarise(平均滿意度 = mean(滿意度總分, na.rm = TRUE))

by_month2 <- by_month %>%
  left_join(avg_score_by_month, by = "月份")

p1 <- ggplot(by_month2, aes(月份, 回函率)) + geom_line(color="steelblue") + geom_point(color="steelblue") +
  scale_y_continuous(labels = scales::percent_format()) +
  labs(title="各月回函率", y="回函率")
p2 <- ggplot(by_month2, aes(月份, 平均滿意度)) + geom_line(color="tomato") + geom_point(color="tomato") +
  labs(title="各月平均滿意度（有回函）", y="平均滿意度")
p1; p2




# 用最簡單的二元邏輯回歸：是否回函 ~ 滿意度總分 + 月份（控制季節）
fit0 <- glm(是否回函 ~ 滿意度總分 + 月份, data = df_m, family = "binomial")
summary(fit0)

top_agents <- df %>%
  group_by(服務專員姓名) %>%
  summarise(
    發送數 = n(),
    已回函 = sum(是否回函, na.rm = TRUE),
    回函率 = 已回函 / 發送數,
    平均滿意度 = mean(滿意度總分[是否回函 == 1], na.rm = TRUE)
  )

top_agents %>%
  filter(發送數 >= 30) %>%
  ggplot(aes(x = 平均滿意度, y = 回函率)) +
  geom_point() +
  geom_smooth(method = "lm", se = FALSE, color = "red") +
  labs(title = "平均滿意度 vs 回函率",
       x = "平均滿意度",
       y = "回函率")


df <- df %>%
  mutate(月份 = lubridate::month(ymd_hms(填寫問卷時間)))
df_model <- df %>%
  group_by(服務專員姓名, 月份) %>%   # ← 加上月份分群
  summarise(
    發送數 = n(),
    已回函 = sum(!is.na(填寫問卷時間)),
    回函率 = 已回函 / 發送數,
    平均滿意度 = mean(滿意度總分[!is.na(填寫問卷時間)], na.rm = TRUE),
    簡訊比例 = mean(!is.na(簡訊發送)),
    LINE比例 = mean(!is.na(LINE發送)),
    平均落實度 = mean(落實度總分[!is.na(填寫問卷時間)], na.rm = TRUE)
  )
# 建立多元線性迴歸模型
fit_multi <- lm(回函率 ~ 平均滿意度 + 簡訊比例 + 發送數, data = df_model)


vif(fit_multi)  # 檢查多重共線性

fit_multi2 <- lm(回函率 ~ 平均滿意度 + 平均落實度 + 簡訊比例 + 發送數, data = df_model)
summary(fit_multi2)

fit_multi3 <- lm(回函率 ~ 平均滿意度 * 平均落實度 + 簡訊比例 + 發送數, data = df_model)
summary(fit_multi3)

fit_multi4 <- lm(回函率 ~ 平均滿意度 + 平均落實度 + 簡訊比例 + 發送數 + factor(月份), data = df_model)
summary(fit_multi4)


#question_cols <- grep("^[0-9]+-", names(df), value = TRUE)
#score_questions <- question_cols[!grepl("原因|改進|建議|類別", question_cols)]
df_scored <- df %>%
  filter(是否回函 == 1) %>%
  dplyr::select(c(滿意度總分, all_of(score_questions)))
df_scored <- df_scored %>%
  mutate(across(everything(), ~ as.numeric(as.character(.x))))

df_long <- df %>%
  filter(是否回函 == 1) %>%
  dplyr::select(
    c(服務專員姓名, 滿意度總分, all_of(score_questions))
  ) %>%
  pivot_longer(
    cols = all_of(score_questions),
    names_to = "問題",
    values_to = "分數"
  )

cor_mat <- cor(df_scored, use = "pairwise.complete.obs", method = "spearman")
cor_target <- cor_mat[, "滿意度總分", drop = FALSE]

corrplot(cor_target,
         method = "color",
         is.corr = FALSE,
         col = colorRampPalette(c("blue","white","red"))(200),
         cl.pos = "b",
         tl.col = "black",
         tl.cex = 0.8,
         main = "各題與滿意度總分之相關性")

cor_df <- data.frame(
  問題 = rownames(cor_target),
  相關性 = cor_target[,1]
)

# 排序（從相關性高 → 低）
cor_df <- cor_df %>% arrange(desc(相關性))

ggplot(cor_df, aes(x = reorder(問題, 相關性), y = 相關性, fill = 相關性)) +
  geom_col() +
  coord_flip() +  # 換成橫條圖，超清楚
  scale_fill_gradient2(low = "blue", mid = "white", high = "red", midpoint = 0) +
  labs(title = "各題與滿意度總分之相關性",
       x = "題目", y = "相關性") +
  theme_minimal(base_size = 12) +
  theme(
    axis.text.y = element_text(size = 9),   # 調整文字大小
    plot.title = element_text(hjust = 0.5)
  )

question_stats <- df_scored %>%
  summarise(across(all_of(score_questions), list(
    平均分 = ~ mean(.x, na.rm = TRUE),
    高滿意比例 = ~ mean(.x >= 9, na.rm = TRUE)
  ))) %>%
  pivot_longer(everything(),
               names_to = c("題目","指標"),
               names_sep = "_",
               values_to = "數值") %>%
  pivot_wider(names_from = 指標, values_from = "數值")

df_long <- df_long %>%
  mutate(
    分數 = as.numeric(分數), 
    滿意度總分 = as.numeric(滿意度總分)
  )
cor_by_agent <- df_long %>%
  group_by(服務專員姓名, 問題) %>%
  # 至少要2筆可計算觀測值
  filter(!is.na(分數), !is.na(滿意度總分)) %>%
  filter(n() >= 2) %>%
  summarise(
    重要度 = cor(分數, 滿意度總分, method = "spearman"),
    .groups = "drop"
  )

perf_by_agent <- df_long %>%
  group_by(服務專員姓名, 問題) %>%
  summarise(表現度 = mean(分數, na.rm = TRUE), .groups = "drop")

plot_df_agent <- cor_by_agent %>%
  left_join(perf_by_agent, by = c("服務專員姓名", "問題"))

#互動式
plot_df_agent$服務專員姓名 <- trimws(plot_df_agent$服務專員姓名)

# 取得全部專員名單
agents <- unique(plot_df_agent$服務專員姓名)

# 初始顯示第一位專員
target_agent <- agents[1]

df_a <- plot_df_agent %>% filter(服務專員姓名 == target_agent)

imp_med <- median(df_a$重要度, na.rm = TRUE)
perf_med <- median(df_a$表現度, na.rm = TRUE)
overall_avg_score <- df %>%
  filter(是否回函 == 1) %>%   # 只看有填過問卷者
  summarise(
    平均滿意度總分 = mean(滿意度總分, na.rm = TRUE)
  ) %>%
  pull(平均滿意度總分) / 100   # 1000 分 → /100 → 回到 0~10

p <- plot_ly(
  data = df_a,
  x = ~表現度,
  y = ~重要度,
  type = "scatter",
  mode = "markers+text",
  text = ~問題,
  textposition = "top center",
  marker = list(
    size = 12,
    color = ~重要度,
    colorscale = list(c(0, "white"), c(1, "red")),
    colorbar = list(title = "重要度相關性")
  ),
  showlegend = FALSE
) %>% 
  layout(
    title = paste0("題目重要度-表現度象限圖（專員：", target_agent, "）"),
    xaxis = list(title = "表現度（該題平均分）"),
    yaxis = list(title = "重要度（與滿意度相關）"),
    shapes = list(
      list(type="line", x0=perf_med, x1=perf_med,
           y0=min(df_a$重要度, na.rm=TRUE), y1=max(df_a$重要度, na.rm=TRUE),
           xref="x", yref="y",
           line=list(dash="dot", width=1.5)),
      
      list(type="line", y0=imp_med, y1=imp_med,
           x0=min(df_a$表現度, na.rm=TRUE), x1=max(df_a$表現度, na.rm=TRUE),
           xref="x", yref="y",
           line=list(dash="dot", width=1.5)),
      
      # ★★★ 這條是新加的：整體平均滿意度（固定，不會因人變動）
      list(type="line",
           x0 = overall_avg_score, x1 = overall_avg_score,
           y0 = min(df_a$重要度, na.rm = TRUE), y1 = max(df_a$重要度, na.rm = TRUE),
           xref="x", yref="y",
           line=list(color="blue", dash="dash", width=2))
    ),
    annotations = list(
      list(x = overall_avg_score,
           y = max(df_a$重要度, na.rm = TRUE),
           text = "整體平均滿意度基準",
           showarrow = FALSE,
           xanchor = "left", yanchor = "bottom",
           font=list(color="blue", size=12))
    ),
    updatemenus = list(
      list(
        y = 1.12, x = 1.2,
        buttons = lapply(agents, function(a){
          
          df_sub <- plot_df_agent %>% filter(服務專員姓名 ==!!a)
          
          if(nrow(df_sub) == 0) return(NULL)
          
          imp_med2 <- median(df_sub$重要度, na.rm = TRUE)
          perf_med2 <- median(df_sub$表現度, na.rm = TRUE)
          
          list(
            method = "update",
            label = a,
            args = list(
              list(
                x = list(df_sub$表現度),
                y = list(df_sub$重要度),
                text = list(df_sub$問題),
                marker = list(
                  color = df_sub$重要度,
                  colorscale = list(c(0, "white"), c(1, "red")),
                  colorbar = list(title = "重要度相關性")
                )
              ),
              list(
                title = paste0("題目重要度-表現度象限圖（專員：", a, "）"),
                
                shapes = list(
                  list(type="line",
                       x0 = perf_med2, x1 = perf_med2,
                       y0 = min(df_sub$重要度, na.rm = TRUE), y1 = max(df_sub$重要度, na.rm = TRUE),
                       xref="x", yref="y",
                       line=list(dash="dot", width=1.5)),
                  
                  list(type="line",
                       y0 = imp_med2, y1 = imp_med2,
                       x0 = min(df_sub$表現度, na.rm = TRUE), x1 = max(df_sub$表現度, na.rm = TRUE),
                       xref="x", yref="y",
                       line=list(dash="dot", width=1.5)),
                  
                  # ★★★ 全體固定平均滿意度基準線（垂直藍線）
                  list(type="line",
                       x0 = overall_avg_score, x1 = overall_avg_score,
                       y0 = min(df_sub$重要度, na.rm = TRUE), y1 = max(df_sub$重要度, na.rm = TRUE),
                       xref="x", yref="y",
                       line=list(color="blue", dash="dash", width=2))
                ),
                
                annotations = list(
                  list(x = overall_avg_score,
                       y = max(df_sub$重要度, na.rm = TRUE),
                       text = "整體平均滿意度基準",
                       showarrow = FALSE,
                       xanchor = "left", yanchor = "bottom",
                       font=list(color="blue", size=12))
                )
              )
            )
          )
        })
      )
    )
  )

p

# =================================================
# Construction of the SA Matches Matrix（修正版）
# =================================================

# 1. 顧客回答表：只保留可量化題目
score_questions_raw <- score_questions

# 2. 專員端有做出重要度的題目
question_agent_raw <- unique(plot_df_agent$問題)

# 3. 只取雙方共通題目
common_raw <- intersect(score_questions_raw, question_agent_raw)

# 4. 題目對照表（raw -> safe）
question_map <- data.frame(
  raw  = common_raw,
  safe = make.names(common_raw, unique = TRUE),
  stringsAsFactors = FALSE
)

# 5. 建立 safe 版資料表
df_safe <- df %>%
  rename_with(~ make.names(.x, unique = TRUE))

# -------------------------------------------------
# Step 1. 建立顧客題目寬表模型資料 df_model
# -------------------------------------------------
df_long_match <- df_safe %>%
  filter(是否回函 == 1) %>%
  pivot_longer(
    cols = all_of(question_map$safe),
    names_to = "問題_safe",
    values_to = "分數"
  ) %>%
  mutate(分數 = parse_num_safe(分數)) %>%
  filter(!is.na(分數))

df_model <- df_long_match %>%
  pivot_wider(
    names_from = 問題_safe,
    values_from = 分數,
    values_fn = mean
  )

# 若你想保留原始背景變數供後續回歸使用，要在 pivot 前先留著
bg_vars <- c("主要工作代碼", "交車型態", "車名")
bg_vars_safe <- make.names(bg_vars[bg_vars %in% names(df)], unique = TRUE)

df_model <- df_safe %>%
  filter(是否回函 == 1) %>%
  select(any_of(c(bg_vars_safe, question_map$safe))) %>%
  pivot_longer(
    cols = all_of(question_map$safe),
    names_to = "問題_safe",
    values_to = "分數"
  ) %>%
  mutate(分數 = parse_num_safe(分數)) %>%
  filter(!is.na(分數)) %>%
  group_by(across(any_of(bg_vars_safe)), 問題_safe) %>%
  summarise(分數 = mean(分數, na.rm = TRUE), .groups = "drop") %>%
  pivot_wider(
    names_from = 問題_safe,
    values_from = 分數
  )

# -------------------------------------------------
# Step 2. 建立服務專員特徵矩陣 agent_style_matrix
# -------------------------------------------------
agent_style_matrix <- plot_df_agent %>%
  filter(問題 %in% question_map$raw) %>%
  left_join(question_map, by = c("問題" = "raw")) %>%
  select(服務專員姓名, safe, 重要度) %>%
  pivot_wider(
    names_from  = safe,
    values_from = 重要度
  ) %>%
  column_to_rownames("服務專員姓名")

agent_style_matrix[is.na(agent_style_matrix)] <- 0

# -------------------------------------------------
# Step 3. 逐題建立迴歸模型 (question ~ 車名/工作代碼)
# -------------------------------------------------
vars <- c("主要工作代碼", "交車型態", "車名")
vars_safe <- make.names(vars, unique = TRUE)
vars_safe <- vars_safe[vars_safe %in% names(df_model)]
vars_safe <- vars_safe[sapply(df_model[vars_safe], dplyr::n_distinct) > 1]

coef_list <- list()

for (q in question_map$safe) {
  if (!q %in% names(df_model)) next
  
  if (length(vars_safe) == 0) {
    f <- as.formula(paste0("`", q, "` ~ 1"))
  } else {
    f <- as.formula(
      paste0("`", q, "` ~ ", paste(sprintf("`%s`", vars_safe), collapse = " + "))
    )
  }
  
  fit <- try(lm(f, data = df_model, na.action = na.exclude), silent = TRUE)
  if (inherits(fit, "try-error")) next
  
  cf <- coef(fit)
  if (length(cf) <= 1 || all(is.na(cf))) next
  
  coef_list[[q]] <- cf
}

# -------------------------------------------------
# Step 4. 組合成係數矩陣 coef_matrix
# -------------------------------------------------
all_coef_names <- unique(unlist(lapply(coef_list, names)))

coef_matrix <- do.call(
  rbind,
  lapply(coef_list, function(cf) {
    v <- rep(0, length(all_coef_names))
    names(v) <- all_coef_names
    v[names(cf)] <- cf
    v
  })
)

rownames(coef_matrix) <- names(coef_list)
coef_matrix <- coef_matrix[, colnames(coef_matrix) != "(Intercept)", drop = FALSE]

# -------------------------------------------------
# Step 5. 對齊題目
# -------------------------------------------------
missing_cols <- setdiff(colnames(agent_style_matrix), rownames(coef_matrix))

if (length(missing_cols) > 0) {
  zero_mat <- matrix(0, nrow = length(missing_cols), ncol = ncol(coef_matrix))
  rownames(zero_mat) <- missing_cols
  colnames(zero_mat) <- colnames(coef_matrix)
  coef_matrix_full <- rbind(coef_matrix, zero_mat)
} else {
  coef_matrix_full <- coef_matrix
}

common_questions_safe <- intersect(rownames(coef_matrix_full), colnames(agent_style_matrix))

coef_sub_full <- coef_matrix_full[common_questions_safe, , drop = FALSE]
agent_sub <- agent_style_matrix[, common_questions_safe, drop = FALSE]

# -------------------------------------------------
# Step 6. 正規化與 Match Score 計算
# -------------------------------------------------
normalize <- function(x) {
  if (all(is.na(x))) return(rep(0, length(x)))
  denom <- sqrt(sum(x^2, na.rm = TRUE))
  if (denom == 0) return(rep(0, length(x)))
  x / denom
}

agent_norm <- t(apply(agent_sub, 1, normalize))
coef_norm  <- apply(coef_sub_full, 2, normalize)

match_scores <- as.matrix(agent_norm) %*% as.matrix(coef_norm)

cat("✅ match_scores 維度：", dim(match_scores), "\n")

coef_cols <- colnames(coef_matrix_full)

var_prefix <- c("主要工作代碼", "交車型態", "車名")
var_prefix <- var_prefix[var_prefix %in% names(df_model)]

tab_list <- lapply(var_prefix, function(v) table(df_model[[v]]))
names(tab_list) <- var_prefix

get_n_for_coef <- function(coef_name) {
  v <- var_prefix[startsWith(coef_name, var_prefix)][1]
  if (is.na(v)) return(NA_real_)
  
  lvl <- substr(coef_name, nchar(v) + 1, nchar(coef_name))
  if (lvl == "" || is.na(lvl)) return(NA_real_)
  
  tab <- tab_list[[v]]
  if (lvl %in% names(tab)) as.numeric(tab[[lvl]]) else 0
}

n_vec <- sapply(coef_cols, get_n_for_coef)

# ⭐ shrink 強度（可調）
k <- 20
shrink <- sqrt(pmax(n_vec, 0) / (pmax(n_vec, 0) + k))
shrink[is.na(shrink)] <- 0

# ⭐ 套用 shrink
match_scores <- sweep(match_scores, 2, shrink, `*`)

# -------------------------------------------------
# Step 7. 正規化結果、縮放後可視化
# -------------------------------------------------
match_scores <- t(scale(t(match_scores)))  # 標準化各專員分佈


# --- Step 1. 產生「所有專員」清單 ---
agent_names <- rownames(match_scores)

# --- Step 2. 建立一個 function：輸出單一專員的矩陣 ---
get_agent_matrix <- function(agent) {
  mat <- matrix(match_scores[agent, ], nrow = 1)
  rownames(mat) <- agent
  colnames(mat) <- colnames(match_scores)
  return(mat)
}

# --- Step 3. 建立互動式圖表（初始顯示第一位專員）---
first_agent <- agent_names[1]
first_matrix <- get_agent_matrix(first_agent)

p <- plot_ly(
  z = first_matrix,
  x = colnames(first_matrix),
  y = rownames(first_matrix),
  type = "heatmap",
  colorscale = list(
    c(0, "#4575b4"),  # 柔和藍
    c(0.5, "white"),# 淡黃白中間
    c(1, "#d73027")   # 柔和紅
  ),
  colorbar = list(title = "Match Score"),
  hoverinfo = "x+y+z"
)

# --- Step 7. 下拉選單 ---
dropdown_buttons <- lapply(agent_names, function(agent) {
  mat <- get_agent_matrix(agent)
  list(
    method = "restyle",
    args = list(
      list(
        z = list(mat),
        y = list(rownames(mat))
      )
    ),
    label = agent
  )
})

# --- Step 8. Layout ---
p <- p %>%
  layout(
    title = "服務專員 Match Score 熱圖",
    xaxis = list(title = "工作型態／車型", tickangle = -45),
    yaxis = list(title = "服務專員", autorange = "reversed"),
    updatemenus = list(
      list(
        y = 1.15,
        buttons = dropdown_buttons,
        direction = "down",
        showactive = TRUE,
        x = 0.1,
        xanchor = "left",
        yanchor = "top"
      )
    )
  )

p


#問卷題目對低滿意案件的關聯
df_low <- df %>%
  filter(是否回函 == 1) %>%          # 只取有回函者
  mutate(
    低滿意 = as.numeric(低滿意),      # 確保是 0/1
    整體滿意度 = 滿意度總分 / 100     # 若滿分是1000，改成 /100 或 /10 視比例
  )

df_low <- df_low %>%
  mutate(
    據點名稱 = as.character(據點名稱),
    據點名稱 = str_trim(據點名稱),
    據點名稱 = na_if(據點名稱, "")
  ) %>%
  filter(!is.na(據點名稱))

# 找出題目欄位（以題號開頭者，如 "01-"、"02-"）
question_list <- grep("^[0-9]{2}-", colnames(df_low), value = TRUE)

# 把題目欄位的值全部轉成數字型別
df_low[question_list] <- lapply(df_low[question_list], function(x) as.numeric(as.character(x)))

# 移除沒有變異的題目
question_list <- question_list[
  sapply(df_low[, question_list], function(x) length(unique(x[!is.na(x)])) > 1)
]

df_low <- df_low %>%
  mutate(
    across(
      all_of(question_list),
      parse_num_safe
    )
  )

cor(df_low$低滿意, df_low$整體滿意度, method = "spearman")

model <- lm(整體滿意度 ~ 低滿意, data = df_low)
summary(model)

ggplot(df_low, aes(x = factor(低滿意), y = 整體滿意度)) +
  geom_boxplot(fill = "tomato", alpha = 0.6) +
  labs(
    title = "低滿意 vs 整體滿意度",
    x = "是否出現低滿意（1=有，0=無）",
    y = "整體滿意度"
  )

cor_importance <- cor(
  df_low[, question_list],
  df_low$低滿意,
  method = "spearman",
  use = "pairwise.complete.obs"
)
names(cor_importance) <- question_list

# 取前 15 題（由影響大到小）
head(sort(cor_importance, decreasing = TRUE), 15)

plot_data <- sort(cor_importance, decreasing = TRUE)[1:15] %>%
  as.data.frame() %>%
  rownames_to_column("題目") %>%
  rename(相關性 = 2)

# 在這裡做排序，決定顯示順序
plot_data$題目 <- forcats::fct_reorder(plot_data$題目, abs(plot_data$相關性), .desc = TRUE)

ggplot(plot_data, aes(
  x = reorder(題目, 相關性, FUN = function(x) -x),  # ← 將排序反轉
  y = 相關性,
  fill = 相關性
)) +
  geom_col() +
  coord_flip() +
  scale_fill_gradient2(
    low = "#FF7F7F", mid = "white", high = "#7FB3FF",
    midpoint = 0,
    name = "相關性"
  ) +
  labs(
    title = "影響低滿意發生之關鍵題目",
    x = "題目",
    y = "與低滿意之相關（ρ）"
  ) +
  theme_minimal(base_size = 14)

# ---- 定義 Spearman 分廠繪圖函數 ----
plot_branch_correlation <- function(branch_name, df_branch, question_list) {
  
  cor_importance <- cor(
    df_branch[, question_list],
    df_branch$低滿意,
    method = "spearman",
    use = "pairwise.complete.obs"
  )
  
  cor_importance <- as.numeric(cor_importance)
  names(cor_importance) <- question_list
  
  cor_importance <- cor_importance[!is.na(cor_importance)]
  
  if (length(cor_importance) == 0) {
    return(
      ggplot() +
        annotate("text", x = 1, y = 1, label = paste0(branch_name, "\n無可用相關題目")) +
        theme_void()
    )
  }
  
  # 取前 15 題（影響力最大）
  plot_data <- sort(cor_importance, decreasing = TRUE) %>%
    head(15) %>%
    as.data.frame() %>%
    rownames_to_column("題目")
  
  names(plot_data)[2] <- "相關性"
  
  plot_data$題目 <- factor(
    plot_data$題目,
    levels = plot_data$題目[order(abs(plot_data$相關性), decreasing = TRUE)]
  )
  
  ggplot(plot_data, aes(x = 題目, y = 相關性, fill = 相關性)) +
    geom_col() +
    coord_flip() +
    scale_fill_gradient2(
      low = "#FF7F7F", mid = "white", high = "#7FB3FF",
      midpoint = 0,
      name = "相關性 (ρ)"
    ) +
    labs(
      title = paste0(branch_name, " - 影響低滿意發生之關鍵題目"),
      subtitle = paste0("樣本數 n = ", nrow(df_branch)),
      x = "題目",
      y = "與低滿意之相關（ρ）"
    ) +
    theme_minimal(base_size = 14) +
    theme(
      text = element_text(family = "Microsoft JhengHei"),
      plot.title = element_text(face = "bold", size = 16, hjust = 0.5),
      plot.subtitle = element_text(size = 12, hjust = 0.5)
    )
}

# ---- 主程式：依據據點分廠繪圖 ----
plots_by_branch <- split(df_low, df_low$據點名稱) %>%
  lapply(function(branch_df) {
    plot_branch_correlation(unique(branch_df$據點名稱), branch_df, question_list)
  })

# ---- 檢視結果 ----
# 顯示某個廠的圖，例如：
plots_by_branch[["LS新莊廠"]]
plots_by_branch[["LS濱江廠"]]
plots_by_branch[["LS中和廠"]]
plots_by_branch[["LS士林廠"]]
plots_by_branch[["LS三重廠"]]

# ===== 參數 =====
set.seed(42)
N <- 500                                 # 你要標幾筆
out_path <- "CSI_label_task_500.csv"     # 匯出檔名

# ===== 基礎清理：只取有文字的 =====
df_text <- df %>%
  dplyr::filter(!is.na(`38-寶貴建議`), `38-寶貴建議` != "") %>%
  dplyr::mutate(
    建議文字 = trimws(as.character(`38-寶貴建議`)),
    低滿意 = dplyr::case_when(
      !is.null(`低滿意`) ~ as.integer(`低滿意`),
      TRUE ~ NA_integer_
    ),
    滿意度總分 = suppressWarnings(as.numeric(滿意度總分))
  ) %>%
  dplyr::filter(!is.na(建議文字), nchar(建議文字) >= 3)

# 去重（同一句只留一筆）
df_text <- df_text %>%
  dplyr::distinct(建議文字, .keep_all = TRUE)

# ===== 分層抽樣：盡量各取一半（若某層不足就全取）=====
have_flag <- !all(is.na(df_text$低滿意))
if (have_flag) {
  df0 <- dplyr::filter(df_text, 低滿意 == 0)
  df1 <- dplyr::filter(df_text, 低滿意 == 1)
  n0 <- min(nrow(df0), floor(N/2))
  n1 <- min(nrow(df1), N - n0)
  samp <- dplyr::bind_rows(
    dplyr::slice_sample(df0, n = n0),
    dplyr::slice_sample(df1, n = n1)
  )
  if (nrow(samp) < N) {
    # 若總數仍不足，從剩餘樣本補齊
    rest <- dplyr::anti_join(df_text, samp, by = "建議文字")
    need <- N - nrow(samp)
    if (need > 0 && nrow(rest) > 0) {
      samp <- dplyr::bind_rows(samp, dplyr::slice_sample(rest, n = min(need, nrow(rest))))
    }
  }
} else {
  # 沒有低滿意旗標就單純抽樣
  samp <- dplyr::slice_sample(df_text, n = min(N, nrow(df_text)))
}

# ===== 組標註表 =====
label_task <- samp %>%
  dplyr::transmute(
    text_id = sprintf("T%06d", dplyr::row_number()),
    建議文字,
    # 標註欄（請你回填）：Complaint/Neutral/Praise 或你想用的體系
    label = "",
    
    # 參考用背景欄位（可保留或刪除）
    低滿意 = low <- if (have_flag) as.integer(低滿意) else NA_integer_,
    滿意度總分 = 滿意度總分,
    據點 = dplyr::coalesce(據點名稱, 據點代碼, NA_character_),
    服務專員 = dplyr::coalesce(服務專員姓名, NA_character_),
    入廠日 = as.character(suppressWarnings(lubridate::as_date(入廠日)))
  )

#readr::write_csv(label_task, out_path)
#cat("已匯出：", out_path, "（", nrow(label_task), "筆）\n", sep = "")

#reticulate::conda_list()
#reticulate::use_condaenv("nlp", required = TRUE)
#jieba <- import("jieba")
#seg <- jieba$cut("斷詞測試，服務態度很差，等待太久")
#unlist(seg)
#words
#transformers <- import("transformers")
#cat("🎉 Transformers 成功載入！\n")

df_SAsummary <- df%>%
  mutate(
    填寫問卷時間 = ymd_hms(填寫問卷時間)
  ) %>%
  # 篩選時間區間
  filter(填寫問卷時間 >= ymd("2024-12-21") & 填寫問卷時間 <= ymd("2025-10-20")) %>%
  group_by(服務專員姓名) %>%
  summarise(
    樣本數 = n(),
    高滿意件數 = sum(高滿意 == 1, na.rm = TRUE),
    落實度Aplus件數 = sum(`A+落實` == 1, na.rm = TRUE),
    低滿意件數 = sum(低滿意 == 1, na.rm = TRUE),
    平均滿意度 = mean(滿意度總分, na.rm = TRUE) 
  ) %>%
  mutate(
    A級比例 = 高滿意件數 / 樣本數,
    Aplus比例 = 落實度Aplus件數 / 樣本數,
    低滿意比例 = 低滿意件數 / 樣本數
  )
# 輸出 Excel
#write_xlsx(df_SAsummary, "CS_Agent_Performance_2024-12-21_to_2025-10-20.xlsx")

cat("✅ 已完成專員滿意度統計報表輸出！\n")

# 1️⃣ 加入年月欄位
df_sa_month <- df %>%
  mutate(年月 = format(入廠日, "%Y-%m")) %>%   # ← 新增年月欄位
  group_by(據點代碼,服務專員姓名, 年月) %>%
  summarise(
    接車台數 = n(),
    .groups = "drop"
  )

# 2️⃣ 各月平均
df_avg_month <- df_sa_month %>%
  group_by(據點代碼,年月) %>%
  summarise(平均接車台數 = mean(接車台數, na.rm = TRUE),
            專員數 = n_distinct(服務專員姓名),
            .groups = "drop")

# 4️⃣ 視覺化
# 取出廠別清單
branches <- unique(df_avg_month$據點代碼)

# 建立 plotly 基礎圖層（所有據點）
p <- plot_ly()

for (b in branches) {
  df_b <- df_avg_month %>% filter(據點代碼 == b)
  p <- add_trace(
    p,
    data = df_b,
    x = ~年月,
    y = ~平均接車台數,
    type = "scatter",
    mode = "lines+markers",
    name = b,
    hovertemplate = paste(
      "<b>據點：</b>", b, "<br>",
      "<b>年月：</b>%{x}<br>",
      "<b>平均每專員接車：</b>%{y:.1f} 台<br>",
      "<b>專員數：</b>%{customdata}<extra></extra>"
    ),
    customdata = df_b$專員數
  )
}

# 建立 dropdown 選單
buttons <- list(
  list(method = "update",
       args = list(list(visible = rep(TRUE, length(branches))),
                   list(title = "各廠SA平均每月接車台數（全部）")),
       label = "全部據點")
)

for (i in seq_along(branches)) {
  vis <- rep(FALSE, length(branches))
  vis[i] <- TRUE
  buttons[[length(buttons) + 1]] <- list(
    method = "update",
    args = list(list(visible = vis),
                list(title = paste0("據點 ", branches[i], " 平均每月接車台數"))),
    label = branches[i]
  )
}

# 最終版圖
p <- p %>%
  layout(
    title = list(
      text = "各廠每位專員平均每月接車台數推移（不計回函）",
      x = 0.5
    ),
    xaxis = list(title = "年月", tickangle = 45),
    yaxis = list(title = "平均每專員接車台數", tickformat = ","),
    updatemenus = list(
      list(
        active = 0,
        type = "dropdown",
        x = 1.05,
        y = 0.8,
        buttons = buttons
      )
    ),
    legend = list(title = list(text = "據點代碼"), orientation = "h", x = 0.3, y = -0.25),
    hovermode = "x unified"
  )

p

# Step 1️⃣ 過濾有回函
df_csi <- df %>%
  filter(是否回函 == 1 & !is.na(填寫問卷時間))
# Step 2️⃣ 定義一個「自訂月份分群」函數
#   → 每月從21號開始，到下月20號為止
assign_csi_month <- function(date) {
  if (is.na(date)) return(NA_character_)
  
  date <- as.Date(date)
  month <- month(date)
  year <- year(date)
  
  # 若日期 <= 20日 → 屬於上個月
  if (day(date) <= 20) {
    month <- month - 1
    if (month == 0) {
      month <- 12
      year <- year - 1
    }
  }
  
  sprintf("%04d-%02d", year, month)
}

df_csi <- df_csi %>%
  mutate(CSI月份 = sapply(填寫問卷時間, assign_csi_month))

# Step 3️⃣ 每廠每月統計指標
df_csi_summary <- df_csi %>%
  group_by(據點代碼, CSI月份) %>%
  summarise(
    樣本數 = n(),
    Aplus落實 = sum(`A+落實`, na.rm = TRUE),
    高滿意 = sum(高滿意, na.rm = TRUE),
    低滿意 = sum(低滿意, na.rm = TRUE),
    落實度 = Aplus落實 / 樣本數,
    滿意度 = 高滿意 / 樣本數,
    低滿意比例 = 低滿意 / 樣本數,
    .groups = "drop"
  )

# Step 4️⃣ 檢查結果
df_csi_summary %>%
  arrange(desc(CSI月份))

# 確保欄位型態一致
df_avg_month <- df_avg_month %>%
  rename(年月 = 年月, 據點 = 據點代碼)

df_csi_summary <- df_csi_summary %>%
  rename(年月 = CSI月份, 據點 = 據點代碼)

# 合併兩邊資料
df_model2 <- df_csi_summary %>%
  left_join(df_avg_month, by = c("據點", "年月"))

# 檢查
glimpse(df_model2)

df_model_lag <- df_model2 %>%
  group_by(據點) %>%
  arrange(年月) %>%
  mutate(
    上月平均接車台數 = lag(平均接車台數, n = 0),
    上月樣本數 = lag(樣本數, n = 0)
  ) %>%
  ungroup()

# 要分析的三個指標
target_vars <- c("落實度", "滿意度", "低滿意比例")

# 結果表
result_lag <- purrr::map_dfr(target_vars, function(var) {
  formula_str <- paste0(var, " ~ 上月平均接車台數 + 上月樣本數")
  model <- lm(as.formula(formula_str), data = df_model_lag)
  summary_m <- summary(model)
  
  tibble(
    指標 = var,
    R平方 = summary_m$r.squared,
    調整後R平方 = summary_m$adj.r.squared,
    F統計量 = summary_m$fstatistic[1],
    p值 = pf(summary_m$fstatistic[1], summary_m$fstatistic[2], summary_m$fstatistic[3], lower.tail = FALSE),
    接車估計 = coef(summary_m)[2],
    接車_p值 = coef(summary_m)[,4][2],
    樣本估計 = coef(summary_m)[3],
    樣本_p值 = coef(summary_m)[,4][3]
  )
})

print(result_lag)

plot_ly(
  result_lag,
  x = ~指標,
  y = ~R平方,
  type = "bar",
  name = "R²",
  marker = list(color = "#1f77b4")
) %>%
  add_trace(
    y = ~調整後R平方,
    name = "調整後R²",
    marker = list(color = "#ff7f0e")
  ) %>%
  layout(
    title = list(text = "各CSI指標與接車量Lag模型解釋力", x = 0.5),
    barmode = "group",
    xaxis = list(title = "指標"),
    yaxis = list(title = "R平方值", range = c(0, 1)),
    legend = list(orientation = "h", x = 0.3, y = -0.2)
  )

#入廠時間對於CSI成績的變化

df_trend_week <- df %>%
  filter(!is.na(入廠日) & 是否回函 == 1) %>%
  mutate(
    週期 = floor_date(as.Date(入廠日), "week", week_start = 1)  # ← 週期從週一開始
  ) %>%
  group_by( 週期) %>%
  summarise(
    樣本數 = n(),
    Aplus落實 = sum(`A+落實`, na.rm = TRUE),
    高滿意 = sum(高滿意, na.rm = TRUE),
    低滿意 = sum(低滿意, na.rm = TRUE),
    落實度比例 = round(Aplus落實 / 樣本數, 4),
    滿意度比例 = round(高滿意 / 樣本數, 4),
    低滿意比例 = round(低滿意 / 樣本數, 4),
    .groups = "drop"
  )

p <- plot_ly(df_trend_week, x = ~週期) %>%
  add_trace(
    y = ~落實度比例, name = "落實度",
    type = 'scatter', mode = 'lines+markers',
    line = list(color = "#1f77b4"),
    hovertemplate = "週期: %{x}<br>落實度: %{y:.2%}<extra></extra>"
  ) %>%
  add_trace(
    y = ~滿意度比例, name = "滿意度",
    type = 'scatter', mode = 'lines+markers',
    line = list(color = "#2ca02c"),
    hovertemplate = "週期: %{x}<br>滿意度: %{y:.2%}<extra></extra>"
  ) %>%
  add_trace(
    y = ~低滿意比例, name = "低滿意比例",
    type = 'scatter', mode = 'lines+markers',
    line = list(color = "#d62728"),
    hovertemplate = "週期: %{x}<br>低滿意比例: %{y:.2%}<extra></extra>"
  ) %>%
  layout(
    title = list(text = "顧客入廠日期對 CSI 指標變化（週平均）", x = 0.5),
    xaxis = list(title = "入廠週期", tickangle = 45),
    yaxis = list(title = "比例", tickformat = ".2%"),
    hovermode = "x unified",
    legend = list(orientation = "h", x = 0.3, y = -0.25)
  )

p

df_weekday <- df_csi %>%
  filter(是否回函 == 1, !is.na(入廠日)) %>%
  mutate(
    星期 = wday(入廠日, week_start = 1, label = TRUE, abbr = FALSE)  # 星期一=1
  ) %>%
  group_by(據點代碼, 星期) %>%
  summarise(
    樣本數 = n(),
    落實度 = sum(`A+落實`, na.rm = TRUE) / 樣本數,
    滿意度 = sum(高滿意, na.rm = TRUE) / 樣本數,
    低滿意比例 = sum(低滿意, na.rm = TRUE) / 樣本數,
    .groups = "drop"
  )
branches <- unique(df_weekday$據點代碼)

p <- plot_ly()

for (b in branches) {
  df_b <- df_weekday %>% filter(據點代碼 == b)
  
  p <- p %>%
    add_trace(
      data = df_b, x = ~星期, y = ~落實度,
      type = "scatter", mode = "lines+markers",
      name = paste(b, "落實度"),
      line = list(color = "#1f77b4"),
      visible = ifelse(b == branches[1], TRUE, FALSE),
      customdata = df_b$樣本數,
      hovertemplate = paste(
        "<b>據點：</b>", b, "<br>",
        "星期：%{x}<br>",
        "落實度：%{y:.2%}<br>",
        "樣本數：%{customdata}<extra></extra>"
      )
    ) %>%
    add_trace(
      data = df_b, x = ~星期, y = ~滿意度,
      type = "scatter", mode = "lines+markers",
      name = paste(b, "滿意度"),
      line = list(color = "#2ca02c"),
      visible = ifelse(b == branches[1], TRUE, FALSE),
      customdata = df_b$樣本數,
      hovertemplate = paste(
        "<b>據點：</b>", b, "<br>",
        "星期：%{x}<br>",
        "滿意度：%{y:.2%}<br>",
        "樣本數：%{customdata}<extra></extra>"
      )
    ) %>%
    add_trace(
      data = df_b, x = ~星期, y = ~低滿意比例,
      type = "scatter", mode = "lines+markers",
      name = paste(b, "低滿意比例"),
      line = list(color = "#d62728"),
      visible = ifelse(b == branches[1], TRUE, FALSE),
      customdata = df_b$樣本數,
      hovertemplate = paste(
        "<b>據點：</b>", b, "<br>",
        "星期：%{x}<br>",
        "低滿意比例：%{y:.2%}<br>",
        "樣本數：%{customdata}<extra></extra>"
      )
    )
}

# dropdown切換
buttons <- list()
for (i in seq_along(branches)) {
  vis <- rep(FALSE, length(branches) * 3)
  vis[(i - 1) * 3 + 1:3] <- TRUE
  buttons[[i]] <- list(
    method = "update",
    args = list(list(visible = vis),
                list(title = paste0("據點 ", branches[i], " 星期分布趨勢"))),
    label = branches[i]
  )
}

p <- p %>%
  layout(
    title = list(text = "各據點星期一～六 CSI 指標分布趨勢", x = 0.5),
    xaxis = list(title = "星期", tickangle = 0),
    yaxis = list(title = "比例（%）", tickformat = ".2%"),
    updatemenus = list(list(active = 0, type = "dropdown", x = 1.05, y = 0.8, buttons = buttons)),
    hovermode = "x unified",
    legend = list(orientation = "h", x = 0.3, y = -0.25)
  )

p

plot_ly(df_weekday, x = ~星期, y = ~滿意度, color = ~據點代碼, type = "violin") %>%
  layout(title = "各據點星期滿意度分布", yaxis = list(title = "滿意度", tickformat = ".2%"))

df_weekday_agent <- df_csi %>%
  filter(是否回函 == 1, !is.na(入廠日)) %>%
  mutate(
    星期 = wday(入廠日, week_start = 1, label = TRUE, abbr = FALSE)
  ) %>%
  group_by(服務專員姓名, 星期) %>%
  summarise(
    樣本數 = n(),
    落實度 = sum(`A+落實`, na.rm = TRUE) / 樣本數,
    滿意度 = sum(高滿意, na.rm = TRUE) / 樣本數,
    低滿意比例 = sum(低滿意, na.rm = TRUE) / 樣本數,
    .groups = "drop"
  )
agents <- unique(df_weekday_agent$服務專員姓名)

# 建立圖形容器
p <- plot_ly()

for (a in agents) {
  df_a <- df_weekday_agent %>% filter(服務專員姓名 == a)
  
  # 三條線：落實度、滿意度、低滿意比例
  p <- p %>%
    add_trace(
      data = df_a, x = ~星期, y = ~落實度,
      type = "scatter", mode = "lines+markers",
      name = paste(a, "落實度"),
      line = list(color = "#1f77b4", width = 2, dash = "dot"),
      visible = ifelse(a == agents[1], TRUE, FALSE),
      customdata = df_a$樣本數,
      hovertemplate = paste(
        "<b>專員：</b>", a, "<br>",
        "星期：%{x}<br>",
        "落實度：%{y:.2%}<br>",
        "樣本數：%{customdata}<extra></extra>"
      )
    ) %>%
    add_trace(
      data = df_a, x = ~星期, y = ~滿意度,
      type = "scatter", mode = "lines+markers",
      name = paste(a, "滿意度"),
      line = list(color = "#2ca02c", width = 2),
      visible = ifelse(a == agents[1], TRUE, FALSE),
      customdata = df_a$樣本數,
      hovertemplate = paste(
        "<b>專員：</b>", a, "<br>",
        "星期：%{x}<br>",
        "滿意度：%{y:.2%}<br>",
        "樣本數：%{customdata}<extra></extra>"
      )
    ) %>%
    add_trace(
      data = df_a, x = ~星期, y = ~低滿意比例,
      type = "scatter", mode = "lines+markers",
      name = paste(a, "低滿意比例"),
      line = list(color = "#d62728", width = 2),
      visible = ifelse(a == agents[1], TRUE, FALSE),
      customdata = df_a$樣本數,
      hovertemplate = paste(
        "<b>專員：</b>", a, "<br>",
        "星期：%{x}<br>",
        "低滿意比例：%{y:.2%}<br>",
        "樣本數：%{customdata}<extra></extra>"
      )
    )
}

# 建立 dropdown（選專員）
buttons_agent <- list()
for (i in seq_along(agents)) {
  vis <- rep(FALSE, length(agents) * 3)
  vis[(i - 1) * 3 + 1:3] <- TRUE
  buttons_agent[[i]] <- list(
    method = "update",
    args = list(list(visible = vis),
                list(title = paste0("專員 ", agents[i], " CSI 星期分布"))),
    label = agents[i]
  )
}

# 版面設定
p <- p %>%
  layout(
    title = list(text = "各專員 CSI 指標分布（星期一～六）", x = 0.5),
    xaxis = list(title = "星期"),
    yaxis = list(title = "比例（%）", tickformat = ".2%"),
    hovermode = "x unified",
    updatemenus = list(list(
      active = 0,
      type = "dropdown",
      x = 1.05,
      y = 0.8,
      buttons = buttons_agent,
      showactive = TRUE,
      bgcolor = "white",
      bordercolor = "gray",
      borderwidth = 1
    )),
    legend = list(orientation = "h", x = 0.3, y = -0.25)
  )

p

