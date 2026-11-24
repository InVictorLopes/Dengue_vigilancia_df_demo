# fct_download.R -- tabelas do botao "Baixar dados" (CSV ou XLSX)
#
# Uma linha por unidade territorial (RA ou setor censitario), com a populacao
# e as colunas do indicador escolhido. Os indicadores do periodo usam o mesmo
# calculo do mapa (calc_ind_territorio()); a estratificacao vem pronta do ETL
# e nao depende do periodo nem dos filtros.

INDICADORES_DOWNLOAD <- list(
  "Indicadores do período (seguem os filtros)" = c(
    "Incidência (casos/população)" = "inc",
    "Mortalidade (óbitos por dengue/população)" = "mort",
    "Letalidade (óbitos por dengue/casos)" = "let",
    "Proporção de cura (curados/casos encerrados)" = "prop_cura",
    "Proporção de casos graves (graves/confirmados)" = "prop_grave",
    "Todos os indicadores do período" = "todos"
  ),
  "Estratificação de risco (série histórica)" = c(
    "Estratificação de risco (I1–I4, escore e estrato)" = "estrat"
  )
)

INDICADORES_PERIODO <- c("inc", "mort", "let", "prop_cura", "prop_grave")

rotulo_indicador_download <- function(indicador) {
  todos <- unlist(unname(INDICADORES_DOWNLOAD))
  names(todos)[match(indicador, todos)]
}

#' Tabela para download: uma linha por RA/setor
#' @noRd
tabela_download <- function(dados, periodo, f, indicador, unidade = "ra") {
  if (indicador == "estrat") {
    e <- dados$estratificacao[[unidade]]
    if (is.null(e)) stop("Estratificação indisponível -- regere os dados (dev/setup_databases.R).")
    e <- e[order(-e$escore, na.last = TRUE), , drop = FALSE]
    serie <- paste0(e$ano_inicial[1], "–", e$ano_final[1])
    base <- if (unidade == "setor") {
      data.frame("Setor censitário" = e$codigo, "Região Administrativa" = e$ra_nome, check.names = FALSE)
    } else {
      data.frame("Região Administrativa" = e$nome, check.names = FALSE)
    }
    tab <- data.frame(
      "População (Censo 2022)" = e$populacao,
      "Casos prováveis" = e$casos_totais,
      "Casos graves" = e$graves_totais,
      "Óbitos pelo agravo" = e$obitos_totais,
      "I1 magnitude (incidência média anual, /100 mil)" = round(e$I1, 2),
      "I2 recorrência (anos com incidência ≥ 300/100 mil)" = e$I2,
      "I3 situação recente (incidência, /100 mil)" = round(e$I3, 2),
      "I4 gravidade (incidência média de graves, /100 mil)" = round(e$I4, 2),
      "Escore composto" = round(e$escore, 4),
      "Estrato de risco" = ifelse(is.na(e$estrato), "Sem dados", e$estrato),
      "Série histórica" = serie,
      check.names = FALSE
    )
    return(cbind(base, tab))
  }

  if (unidade == "setor") {
    base <- data.frame("Setor censitário" = dados$setores$codigo,
                       "Região Administrativa" = dados$setores$ra_nome, check.names = FALSE)
    pop <- dados$setores$populacao
  } else {
    base <- data.frame("Região Administrativa" = dados$bairros$name_subdistrict, check.names = FALSE)
    pop <- dados$bairros$pop
  }
  base[["População (Censo 2022)"]] <- pop

  # mesma conta do mapa: taxa do periodo dividida pelos anos do periodo
  # (minimo 1 ano) -- ver especificar_indicador()
  anos <- max(1, as.numeric(periodo[2] - periodo[1]) / 365)
  ind <- function(i) calc_ind_territorio(dados, periodo, f, i, unidade)
  pct <- function(x) round(100 * x, 2)

  partes <- list(
    inc = function() {
      t <- ind("inc")
      data.frame("Casos no período" = t$casos,
                 "Incidência no período (por 100 mil hab.)" = round(t$ind, 2),
                 "Incidência anual (por 100 mil hab./ano)" = round(t$ind / anos, 2),
                 check.names = FALSE)
    },
    mort = function() {
      t <- ind("mort")
      data.frame("Óbitos por dengue no período" = t$obitos,
                 "Mortalidade no período (por 100 mil hab.)" = round(t$ind, 3),
                 "Mortalidade anual (por 100 mil hab./ano)" = round(t$ind / anos, 3),
                 check.names = FALSE)
    },
    let = function() {
      t <- ind("let")
      data.frame("Óbitos por dengue no período" = t$obitos,
                 "Casos no período" = t$casos,
                 "Letalidade (%)" = pct(t$ind),
                 check.names = FALSE)
    },
    prop_cura = function() {
      t <- ind("prop_cura")
      data.frame("Curados" = t$curados,
                 "Casos encerrados no período" = t$encerrados,
                 "Proporção de cura (%)" = pct(t$ind),
                 check.names = FALSE)
    },
    prop_grave = function() {
      t <- ind("prop_grave")
      data.frame("Casos graves" = t$graves,
                 "Casos confirmados no período" = t$confirmados,
                 "Proporção de casos graves (%)" = pct(t$ind),
                 check.names = FALSE)
    }
  )

  ids <- if (indicador == "todos") INDICADORES_PERIODO else indicador
  tab <- do.call(cbind, c(list(base), lapply(ids, function(i) partes[[i]]())))
  # "todos": casos e obitos aparecem em mais de um indicador
  tab <- tab[, !duplicated(names(tab)), drop = FALSE]
  # NaN (0/0) e Inf (sem populacao) viram celula vazia
  for (n in names(tab)) if (is.numeric(tab[[n]])) tab[[n]][!is.finite(tab[[n]])] <- NA
  tab
}

#' Grava a tabela em CSV (padrao Excel pt-BR: ";" e virgula decimal, UTF-8
#' com BOM) ou XLSX (aba "Dados" + aba "Informações" com periodo e filtros)
#' @noRd
escrever_download <- function(tab, info, arquivo, formato = c("xlsx", "csv")) {
  formato <- match.arg(formato)
  if (formato == "csv") {
    con <- file(arquivo, open = "w", encoding = "UTF-8")
    on.exit(close(con))
    cat("﻿", file = con)
    utils::write.table(tab, con, sep = ";", dec = ",", row.names = FALSE, na = "", qmethod = "double")
    return(invisible(arquivo))
  }

  wb <- openxlsx::createWorkbook()
  openxlsx::addWorksheet(wb, "Dados")
  openxlsx::writeDataTable(wb, "Dados", tab, tableStyle = "TableStyleLight9", withFilter = TRUE)
  openxlsx::freezePane(wb, "Dados", firstRow = TRUE)
  openxlsx::setColWidths(wb, "Dados", cols = seq_len(ncol(tab)), widths = pmin(45, pmax(12, nchar(names(tab)) + 2)))

  openxlsx::addWorksheet(wb, "Informações")
  openxlsx::writeData(wb, "Informações", info, headerStyle = openxlsx::createStyle(textDecoration = "bold"))
  openxlsx::setColWidths(wb, "Informações", cols = 1:2, widths = c(32, 100))
  openxlsx::addStyle(wb, "Informações", openxlsx::createStyle(wrapText = TRUE, valign = "top"),
                     rows = seq_len(nrow(info)) + 1, cols = 2, gridExpand = TRUE)
  openxlsx::saveWorkbook(wb, arquivo, overwrite = TRUE)
  invisible(arquivo)
}
