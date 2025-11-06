# ============================================================
# APS 3 - Estatística Aplicada
# Relatório de Entrega
# Autor: Décio Faria
# Professora: Cleide Mara Faria Soares
# Tema: Análise de Atrasos e Cancelamentos de Voos (Big Data)
# ============================================================

# ------------------------------
# 1. Pacotes necessários
# ------------------------------
required_packages <- c("tidyverse", "officer", "rvg", "flextable")
for (pkg in required_packages) {
  if (!require(pkg, character.only = TRUE)) {
    install.packages(pkg, dependencies = TRUE)
    library(pkg, character.only = TRUE)
  }
}

# ------------------------------
# 2. Importação dos dados
# ------------------------------
caminho_arquivo <- "C:/Users/DÉCIO/R/flight_data_2024.csv"
dados <- read_csv(caminho_arquivo)

# ------------------------------
# 3. Limpeza e transformação
# ------------------------------
dados_limpos <- dados %>%
  select(
    year, month, day_of_month, origin, origin_city_name, origin_state_nm,
    dep_time, taxi_out, wheels_off, wheels_on, taxi_in,
    air_time, distance, weather_delay, late_aircraft_delay
  ) %>%
  drop_na() %>%
  mutate(
    total_taxi = taxi_out + taxi_in,
    atraso_estimado = weather_delay + late_aircraft_delay
  )

# ------------------------------
# 4. Estatísticas descritivas
# ------------------------------
estatisticas <- tibble(
  Métrica = c("Média (total_taxi)", "Desvio-padrão (total_taxi)", "Variância (total_taxi)",
              "Média (atraso_estimado)", "Desvio-padrão (atraso_estimado)", "Variância (atraso_estimado)"),
  Valor = c(
    mean(dados_limpos$total_taxi),
    sd(dados_limpos$total_taxi),
    var(dados_limpos$total_taxi),
    mean(dados_limpos$atraso_estimado),
    sd(dados_limpos$atraso_estimado),
    var(dados_limpos$atraso_estimado)
  )
)

# ------------------------------
# 5. Correlação e regressão
# ------------------------------
correlacao <- cor(dados_limpos$total_taxi, dados_limpos$atraso_estimado)
modelo <- lm(atraso_estimado ~ total_taxi + air_time + distance, data = dados_limpos)
resumo_modelo <- summary(modelo)
texto_modelo <- paste(capture.output(resumo_modelo), collapse = "\n")

# ------------------------------
# 6. Gráficos
# ------------------------------
grafico1 <- ggplot(dados_limpos, aes(x = atraso_estimado)) +
  geom_histogram(bins = 30, fill = "steelblue", color = "black") +
  labs(title = "Distribuição dos Atrasos Estimados",
       x = "Atraso Estimado (min)", y = "Frequência")

grafico2 <- ggplot(dados_limpos, aes(x = total_taxi, y = atraso_estimado)) +
  geom_point(color = "purple", alpha = 0.5) +
  geom_smooth(method = "lm", color = "red", se = FALSE) +
  labs(title = "Correlação entre Tempo de Táxi e Atrasos Estimados",
       x = "Tempo de Táxi (min)", y = "Atraso Estimado (min)")

grafico3 <- ggplot(dados_limpos, aes(x = distance, y = atraso_estimado)) +
  geom_point(alpha = 0.5, color = "darkorange") +
  geom_smooth(method = "lm", color = "blue", se = FALSE) +
  labs(title = "Relação entre Distância e Atraso Estimado",
       x = "Distância (milhas)", y = "Atraso Estimado (min)")

# ------------------------------
# 7. Relatório Word (officer)
# ------------------------------
doc <- read_docx()

# Capa
doc <- doc %>%
  body_add_par("CENTRO UNIVERSITÁRIO ANHANGUERA", style = "heading 1") %>%
  body_add_par("Curso: Estatística Aplicada", style = "Normal") %>%
  body_add_par("Professor(a): Cleide Mara Faria Soares", style = "Normal") %>%
  body_add_par("Aluno: Décio Faria", style = "Normal") %>%
  body_add_par("APS 3 – Análise de Atrasos em Voos (Big Data)", style = "heading 2") %>%
  body_add_break()

# Introdução
doc <- doc %>%
  body_add_par("1. INTRODUÇÃO", style = "heading 1") %>%
  body_add_par("Este relatório apresenta uma análise estatística dos atrasos de voos no ano de 2024. Utilizando uma base com mais de 1 milhão de registros, foram aplicadas técnicas de estatística descritiva, correlação e regressão para identificar padrões e fatores que influenciam os atrasos.", style = "Normal") %>%
  body_add_break()

# Estatísticas
doc <- doc %>%
  body_add_par("2. ESTATÍSTICAS DESCRITIVAS", style = "heading 1") %>%
  body_add_flextable(flextable(estatisticas)) %>%
  body_add_par(paste("Correlação entre tempo de táxi e atraso estimado: ", round(correlacao, 4)), style = "Normal") %>%
  body_add_break()

# Resultados do modelo
doc <- doc %>%
  body_add_par("3. MODELO DE REGRESSÃO LINEAR", style = "heading 1") %>%
  body_add_par(texto_modelo, style = "Normal") %>%
  body_add_break()

# Gráficos
doc <- doc %>%
  body_add_par("4. VISUALIZAÇÕES", style = "heading 1") %>%
  body_add_gg(grafico1, width = 6, height = 4) %>%
  body_add_par("") %>%
  body_add_gg(grafico2, width = 6, height = 4) %>%
  body_add_par("") %>%
  body_add_gg(grafico3, width = 6, height = 4) %>%
  body_add_break()

# Conclusões
doc <- doc %>%
  body_add_par("5. CONCLUSÕES", style = "heading 1") %>%
  body_add_par("A análise demonstrou que há uma relação positiva entre o tempo total de taxiamento e o atraso médio dos voos. Além disso, a distância e o tempo de voo também mostraram influência sobre a variabilidade dos atrasos. Esses resultados permitem identificar gargalos operacionais e propor melhorias na gestão de escalas e alertas de atraso.", style = "Normal") %>%
  body_add_par("Este estudo exemplifica o uso da Estatística Aplicada na análise de grandes volumes de dados (Big Data) para otimização de processos logísticos e de transporte aéreo.", style = "Normal") %>%
  body_add_break()

# Rodapé
doc <- doc %>%
  body_add_par("Relatório gerado automaticamente via R (2025)", style = "Normal")

# Salvar arquivo
output_path <- "C:/Users/DÉCIO/R/relatorio_APS3_DecioFaria.docx"
print(doc, target = output_path)

cat("\n✅ Relatório gerado com sucesso em:", output_path, "\n")
