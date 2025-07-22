library(shiny)
library(ggplot2)
library(dplyr)
library(leaflet)
library(sf)
library(scales)
library(plotly)
library(tidyr)
library(lmtest) 
library(car)
library(openxlsx)
library(shinyjs)
library(htmlwidgets)
library(webshot2)


shinyServer(function(input, output, session) {
  useShinyjs()
  #======================= Beranda ===================
  # 1. == Jumlah Distrik ==
  output$jumlah_distrik <- renderValueBox({
    valueBox(
      value = nrow(data),
      subtitle = "Jumlah Distrik",
      icon = icon("map")
    )
  })
  
  # 2. == Rata-rata Pendidikan Rendah ==
  output$jumlah_variabel <- renderValueBox({
    rata_pendidikan_rendah <- round(mean(data$LOWEDU, na.rm = TRUE), 2)
    valueBox(
      value = paste0(rata_pendidikan_rendah, " %"),
      subtitle = "Rata-rata Pendidikan Rendah",
      icon = icon("chart-line")
    )
  })
  
  # 3. == Pendidikan Rendah Tertinggi ==
  output$total_populasi <- renderValueBox({
    max_pendidikan_rendah <- round(max(data$LOWEDU, na.rm = TRUE), 2)
    valueBox(
      value = paste0(max_pendidikan_rendah, " %"),
      subtitle = "Pendidikan Rendah Tertinggi",
      icon = icon("exclamation-triangle")
    )
  })
  
  library(officer)
  library(flextable)
  
  output$unduh_beranda <- downloadHandler(
    filename = function() {
      paste("Halaman-Beranda_", Sys.Date(), ".docx", sep = "")
    },
    content = function(file) {
      doc <- read_docx()
      
      # Tambahkan Judul
      doc <- doc %>%
        body_add_par("Dashboard Analisis Pendidikan Rendah di Indonesia", style = "heading 1") %>%
        body_add_par("Dashboard ini menyajikan analisis statistik terhadap tingkat pendidikan rendah (LOWEDU) di Indonesia berdasarkan data SUSENAS 2017.", style = "Normal") %>%
        body_add_par("Analisis ini menelaah keterkaitan antara pendidikan rendah dengan berbagai indikator sosial dan ekonomi, meliputi persentase lansia (ELDERLY), kepala rumah tangga perempuan (FHEAD), ukuran rata-rata rumah tangga (FAMILYSIZE), kemiskinan (POVERTY), akses terhadap listrik (NOELECTRIC), dan status kepemilikan tempat tinggal (RENTED).", style = "Normal") %>%
        body_add_par("Dashboard ini dikembangkan sebagai Projek Ujian Akhir Semester Mata Kuliah Komputasi Statistik dan sebagai alat bantu interaktif untuk mendukung pemahaman konsep statistik deskriptif, inferensia, serta analisis regresi linier dalam konteks pembangunan sosial dan pengentasan ketimpangan sosial.", style = "Normal")
      
      # Tambahkan Metadata
      doc <- doc %>%
        body_add_par("Informasi Metadata", style = "heading 2") %>%
        body_add_par("- Nama Dataset: Social Vulnerability Data (Indonesia), Kurniawan et al. (2022), Elsevier Data in Brief", style = "Normal") %>%
        body_add_par("- Jumlah Distrik: 511 kabupaten/kota (SUSENAS 2017 & Peta Spasial 2013)", style = "Normal") %>%
        body_add_par("- Variabel Utama: LOWEDU, POVERTY, ELDERLY, FHEAD, RENTED", style = "Normal") %>%
        body_add_par("- Sumber Data: SUSENAS 2017, Proyeksi Penduduk BPS 2017, Peta Jarak Spasial", style = "Normal") %>%
        body_add_par("- Format Data: CSV dan GeoJSON, termasuk matriks jarak", style = "Normal") %>%
        body_add_par("- Tujuan: Analisis faktor sosial yang berkontribusi terhadap pendidikan rendah", style = "Normal") %>%
        body_add_par("- Tautan Publikasi: https://doi.org/10.1016/j.dib.2021.107743", style = "Normal") %>%
        body_add_par("- Akses Dataset: SOVI Data & Matriks Jarak", style = "Normal") %>%
        body_add_par("- Lisensi: CC-BY", style = "Normal")
      
      # Tambahkan Tabel Deskripsi Variabel
      deskripsi_tbl <- data.frame(
        Label = c("LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED"),
        Nama_Variabel = c(
          "Pendidikan Rendah", "Penduduk Lansia", "Kepala Rumah Tangga Perempuan",
          "Ukuran Rumah Tangga", "Kemiskinan", "Tanpa Listrik", "Menyewa Rumah"
        ),
        Deskripsi = c(
          "Persentase penduduk usia 15 tahun ke atas dengan pendidikan rendah.",
          "Persentase penduduk usia 65 tahun ke atas.",
          "Persentase rumah tangga dengan kepala keluarga perempuan.",
          "Rata-rata jumlah anggota rumah tangga dalam satu rumah tangga.",
          "Persentase penduduk miskin.",
          "Persentase rumah tangga tanpa sumber listrik.",
          "Persentase rumah tangga yang menyewa tempat tinggal."
        ),
        stringsAsFactors = FALSE
      )
      
      # Buat flextable dengan styling rapi
      ft <- flextable(deskripsi_tbl) %>%
        autofit() %>%
        set_table_properties(width = 1, layout = "autofit") %>%
        align(align = "left", part = "all") %>%
        fontsize(size = 10, part = "all") %>%
        bold(part = "header") %>%
        theme_booktabs()
      
      # Tambahkan ke dokumen
      doc <- doc %>%
        body_add_par("Deskripsi Variabel:", style = "heading 2") %>%
        body_add_flextable(ft)
      
      # Simpan file
      print(doc, target = file)
    }
  )
  
  
  #======================= Eksplorasi Data : Tabel===================

  #1. Tabel Data Asli 
  output$tabel_data_asli <- DT::renderDataTable({
    data_murni <- data[, c("DISTRICTCODE", "nmprov", "nmkab", 
                           "LOWEDU", "ELDERLY", "FHEAD", 
                           "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED")]
    
    names(data_murni)[names(data_murni) == "nmkab"] <- "KAB"
    names(data_murni)[names(data_murni) == "nmprov"] <- "PROV"
    
    is_num <- sapply(data_murni, is.numeric)
    data_murni[is_num] <- lapply(data_murni[is_num], function(x) round(x, 5))
    
    DT::datatable(
      data_murni,
      options = list(scrollX = TRUE, pageLength = 5),
      rownames = FALSE
    )
  })

  #2. Download Data Asli 
  output$unduh_data_asli <- downloadHandler(
    filename = function() {
      paste0("data_awal_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      df_download <- data[, c("DISTRICTCODE", "nmprov", "nmkab",
                              "LOWEDU", "ELDERLY", "FHEAD",
                              "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED")]
      
      names(df_download)[names(df_download) == "nmkab"] <- "KAB"
      names(df_download)[names(df_download) == "nmprov"] <- "PROV"
      
      openxlsx::write.xlsx(df_download, file, rowNames = FALSE)
    }
  )
  
  output$unduh_word_data_asli <- downloadHandler(
    filename = function() {
      paste0("Halaman-Tabel_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      # Fungsi untuk membuat ringkasan statistik
      summary_text <- function(varname) {
        v <- data[[varname]]
        paste0(
          varname, ": Min = ", round(min(v, na.rm = TRUE), 2),
          ", Mean = ", round(mean(v, na.rm = TRUE), 2),
          ", Max = ", round(max(v, na.rm = TRUE), 2),
          ", SD = ", round(sd(v, na.rm = TRUE), 2)
        )
      }
      
      # Daftar ringkasan variabel
      ringkasan <- c(
        summary_text("LOWEDU"),
        summary_text("ELDERLY"),
        summary_text("FHEAD"),
        summary_text("FAMILYSIZE"),
        summary_text("POVERTY"),
        summary_text("NOELECTRIC"),
        summary_text("RENTED")
      )
      
      # Buat dokumen Word hanya berisi ringkasan
      doc <- read_docx()
      doc <- body_add_par(doc, "Ringkasan Statistik", style = "heading 1")
      
      for (line in ringkasan) {
        doc <- body_add_par(doc, line, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )

  #3. Fungsi pembuat ringkasan deskriptif
  summary_box <- function(var, nama, ikon) {
    nilai <- data[[var]]
    info <- paste0(
      "Min: ", round(min(nilai, na.rm = TRUE), 2), " | ",
      "Mean: ", round(mean(nilai, na.rm = TRUE), 2), " | ",
      "Max: ", round(max(nilai, na.rm = TRUE), 2), " | ",
      "SD: ", round(sd(nilai, na.rm = TRUE), 2)
    )
    valueBox(
      value = nama,
      subtitle = info,
      icon = icon(ikon),
    )
  }
  
  output$stat_lowedu <- renderValueBox({
    summary_box("LOWEDU", "LOWEDU", "university")
  })
  output$stat_elderly <- renderValueBox({
    summary_box("ELDERLY", "ELDERLY", "users")
  })
  output$stat_fhead <- renderValueBox({
    summary_box("FHEAD", "FHEAD", "user-tie")
  })
  output$stat_familysize <- renderValueBox({
    summary_box("FAMILYSIZE", "FAMILYSIZE", "people-roof")
  })
  output$stat_poverty <- renderValueBox({
    summary_box("POVERTY", "POVERTY", "house-crack")
  })
  output$stat_noelectric <- renderValueBox({
    summary_box("NOELECTRIC", "NOELECTRIC", "bolt")
  })
  output$stat_rented <- renderValueBox({
    summary_box("RENTED", "RENTED", "key")
  })
  
  
  
  #======================= Eksplorasi Data : Grafik ===================
  
  # Fungsi reaktif untuk menyaring data sesuai provinsi / kabupaten
  data_terfilter <- reactive({
    df <- data
    
    if (!is.null(input$filter_prov) && !"ALL" %in% input$filter_prov) {
      df <- df[df$nmprov %in% input$filter_prov, ]
    }
    
    if (!is.null(input$filter_kab) && length(input$filter_kab) > 0) {
      df <- df[df$nmkab %in% input$filter_kab, ]
    }
    
    df
  })
  
  
  # === 1. Update pilihan kabupaten berdasarkan provinsi ===
  observeEvent(input$filter_prov, {
    req(input$filter_prov)
    
    if ("ALL" %in% input$filter_prov) {
      # Jika ALL dipilih, kosongkan dan disable kabupaten
      updateSelectInput(session, "filter_kab", 
                        choices = sort(unique(data$nmkab)), 
                        selected = character(0))
      shinyjs::disable("filter_kab")
    } else {
      # Jika bukan ALL, update pilihan kabupaten berdasarkan provinsi
      kabupaten_tersedia <- sort(unique(data$nmkab[data$nmprov %in% input$filter_prov]))
      updateSelectInput(session, "filter_kab",
                        choices = kabupaten_tersedia,
                        selected = NULL)
      shinyjs::enable("filter_kab")
    }
  })
  
  
  # === 2. Jika kabupaten dipilih, matikan input provinsi ===
  observe({
    if (!is.null(input$filter_kab) && length(input$filter_kab) > 0) {
      shinyjs::disable("filter_prov")
    } else {
      shinyjs::enable("filter_prov")
    }
  })
  
  # === 3. Reset Filter ===
  observeEvent(input$reset_filter, {
    updateSelectInput(session, "filter_prov", selected = "ALL")
    updateSelectInput(session, "filter_kab", selected = character(0))
    updateSelectInput(session, "var_statistik", selected = "ILLITERATE")
    
    shinyjs::enable("filter_prov")
    shinyjs::disable("filter_kab")
  })
  
  
  # === 3. Grafik ===
  # Plot Histogram
  output$plot_histogram <- renderPlotly({
    req(input$var_statistik)
    df <- data_terfilter()
    var <- input$var_statistik
    
    p <- ggplot(df, aes_string(x = var)) +
      geom_histogram(bins = 30, fill = "#4C6C29", color = "white") +
      theme_minimal() +
      labs(x = var, y = "Frekuensi")
    
    ggplotly(p)
  })
  
  # Plot Boxplot
  output$plot_boxplot <- renderPlotly({
    req(input$var_statistik)
    df <- data_terfilter()
    var <- input$var_statistik
    
    p <- ggplot(df, aes(y = .data[[var]])) +
      geom_boxplot(fill = "#2e7d32", color = "black") +
      theme_minimal() +
      labs(y = var)
    
    ggplotly(p)
  })
  
  # Interpretasi grafik
  output$interpretasi_grafik <- renderUI({
    req(input$var_statistik)
    
    var <- input$var_statistik
    df <- data_terfilter()[[var]]
    
    rata <- round(mean(df, na.rm = TRUE), 2)
    medi <- round(median(df, na.rm = TRUE), 2)
    skew <- ifelse(rata > medi, "positif (right-skewed)",
                   ifelse(rata < medi, "negatif (left-skewed)", "simetris"))
    
    q1 <- round(quantile(df, 0.25, na.rm = TRUE), 2)
    q3 <- round(quantile(df, 0.75, na.rm = TRUE), 2)
    iqr <- round(q3 - q1, 2)
    
    # Ambil filter provinsi dan kabupaten
    prov <- input$filter_prov
    kab <- input$filter_kab
    
    konteks <- if (!is.null(prov) && length(prov) > 0) {
      paste("pada provinsi", paste(prov, collapse = ", "))
    } else if (!is.null(kab) && length(kab) > 0) {
      paste("pada kabupaten", paste(kab, collapse = ", "))
    } else {
      "untuk seluruh distrik"
    }
    
    HTML(paste0(
      "<p style='text-align:justify'>",
      "Visualisasi histogram dan boxplot memberikan gambaran sebaran variabel <b>", var, "</b> ", konteks, ". ",
      "Nilai rata-rata tercatat sebesar <b>", rata, "</b> dan median sebesar <b>", medi, "</b>, menunjukkan distribusi yang <b>", skew, "</b>. ",
      "Boxplot menunjukkan nilai Q1 sebesar <b>", q1, "</b> dan Q3 sebesar <b>", q3, "</b>, menghasilkan IQR (interquartile range) sebesar <b>", iqr, "</b>. ",
      "Outlier dapat diidentifikasi sebagai titik yang berada di luar rentang IQR ± 1.5. ",
      "Interpretasi ini berguna untuk memahami sebaran nilai dan potensi pencilan di daerah yang diamati.",
      "</p>"
    ))
  })
  
  output$unduh_histogram <- downloadHandler(
    filename = function() {
      paste0("histogram_", input$var_statistik, "_", Sys.Date(), ".png")
    },
    content = function(file) {
      df <- data_terfilter()
      var <- input$var_statistik
      png(file, width = 800, height = 600)
      print(
        ggplot(df, aes_string(x = var)) +
          geom_histogram(bins = 30, fill = "#4C6C29", color = "white") +
          theme_minimal() +
          labs(x = var, y = "Frekuensi")
      )
      dev.off()
    }
  )
  
  output$unduh_boxplot <- downloadHandler(
    filename = function() {
      paste0("boxplot_", input$var_statistik, "_", Sys.Date(), ".png")
    },
    content = function(file) {
      df <- data_terfilter()
      var <- input$var_statistik
      png(file, width = 800, height = 600)
      print(
        ggplot(df, aes(y = .data[[var]])) +
          geom_boxplot(fill = "#2e7d32", color = "black") +
          theme_minimal() +
          labs(y = var)
      )
      dev.off()
    }
  )
  
  output$unduh_interpretasi_grafik <- downloadHandler(
    filename = function() {
      paste0("interpretasi_", input$var_statistik, "_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      var <- input$var_statistik
      df <- data_terfilter()[[var]]
      
      rata <- round(mean(df, na.rm = TRUE), 2)
      medi <- round(median(df, na.rm = TRUE), 2)
      skew <- ifelse(rata > medi, "positif (right-skewed)",
                     ifelse(rata < medi, "negatif (left-skewed)", "simetris"))
      
      q1 <- round(quantile(df, 0.25, na.rm = TRUE), 2)
      q3 <- round(quantile(df, 0.75, na.rm = TRUE), 2)
      iqr <- round(q3 - q1, 2)
      
      prov <- input$filter_prov
      kab <- input$filter_kab
      konteks <- if (!is.null(prov) && length(prov) > 0 && !"ALL" %in% prov) {
        paste("pada provinsi", paste(prov, collapse = ", "))
      } else if (!is.null(kab) && length(kab) > 0) {
        paste("pada kabupaten", paste(kab, collapse = ", "))
      } else {
        "untuk seluruh distrik"
      }
      
      kalimat <- paste0(
        "Visualisasi histogram dan boxplot memberikan gambaran sebaran variabel '", var, "' ", konteks, ". ",
        "Rata-rata = ", rata, ", median = ", medi, " → distribusi ", skew, ". ",
        "Q1 = ", q1, ", Q3 = ", q3, ", IQR = ", iqr, "."
      )
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Interpretasi Visualisasi", style = "heading 1")
      doc <- body_add_par(doc, kalimat, style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  
  # Value Box
  output$stat_grafik <- renderValueBox({
    req(input$var_statistik)
    
    var <- input$var_statistik
    df <- data_terfilter()[[var]]
    
    mean_val <- round(mean(df, na.rm = TRUE), 2)
    median_val <- round(median(df, na.rm = TRUE), 2)
    q1 <- round(quantile(df, 0.25, na.rm = TRUE), 2)
    q3 <- round(quantile(df, 0.75, na.rm = TRUE), 2)
    iqr <- round(q3 - q1, 2)
    skew <- ifelse(mean_val > median_val, "Positif", 
                   ifelse(mean_val < median_val, "Negatif", "Simetris"))
    
    valueBox(
      value = tags$b(var),
      subtitle = paste0("Mean: ", mean_val,
                        " | Median: ", median_val,
                        " | Q1: ", q1,
                        " | Q3: ", q3,
                        " | IQR: ", iqr,
                        " | Skew: ", skew),
      icon = icon("chart-bar"),
      color = "olive",
      width = 12
    )
  })
  
  output$unduh_halaman_grafik <- downloadHandler(
    filename = function() {
      paste0("Halaman-Grafik_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      library(ggplot2)
      library(magrittr)
      
      # Data dan variabel
      var <- input$var_statistik
      df_all <- data_terfilter()
      df <- df_all[[var]]
      
      # Statistik ringkas
      mean_val <- round(mean(df, na.rm = TRUE), 2)
      median_val <- round(median(df, na.rm = TRUE), 2)
      q1 <- round(quantile(df, 0.25, na.rm = TRUE), 2)
      q3 <- round(quantile(df, 0.75, na.rm = TRUE), 2)
      iqr <- round(q3 - q1, 2)
      skew <- ifelse(mean_val > median_val, "positif (right-skewed)",
                     ifelse(mean_val < median_val, "negatif (left-skewed)", "simetris"))
      
      # Interpretasi konteks wilayah
      prov <- input$filter_prov
      kab <- input$filter_kab
      konteks <- if (!is.null(prov) && length(prov) > 0 && !"ALL" %in% prov) {
        paste("pada provinsi", paste(prov, collapse = ", "))
      } else if (!is.null(kab) && length(kab) > 0) {
        paste("pada kabupaten", paste(kab, collapse = ", "))
      } else {
        "untuk seluruh distrik"
      }
      
      interpretasi <- paste0(
        "Visualisasi histogram dan boxplot memberikan gambaran sebaran variabel '", var, "' ", konteks, ". ",
        "Rata-rata = ", mean_val, ", median = ", median_val, " → distribusi ", skew, ". ",
        "Q1 = ", q1, ", Q3 = ", q3, ", IQR = ", iqr, "."
      )
      
      # Buat sementara file gambar histogram & boxplot
      hist_file <- tempfile(fileext = ".png")
      box_file <- tempfile(fileext = ".png")
      
      ggsave(hist_file,
             ggplot(df_all, aes_string(x = var)) +
               geom_histogram(bins = 30, fill = "#4C6C29", color = "white") +
               theme_minimal() +
               labs(title = paste("Histogram:", var), x = var, y = "Frekuensi"),
             width = 6, height = 4, dpi = 300)
      
      ggsave(box_file,
             ggplot(df_all, aes(y = .data[[var]])) +
               geom_boxplot(fill = "#2e7d32", color = "black") +
               theme_minimal() +
               labs(title = paste("Boxplot:", var), y = var),
             width = 4, height = 4, dpi = 300)
      
      # Buat dokumen
      doc <- read_docx()
      
      doc <- doc %>%
        body_add_par("Halaman Grafik", style = "heading 1") %>%
        body_add_par(paste0("Variabel: ", var), style = "heading 2") %>%
        body_add_par(paste0("Wilayah: ", konteks), style = "Normal") %>%
        body_add_par(paste0("Rata-rata: ", mean_val, ", Median: ", median_val,
                            ", Q1: ", q1, ", Q3: ", q3, ", IQR: ", iqr, ", Skew: ", skew), style = "Normal") %>%
        body_add_par("Interpretasi:", style = "heading 2") %>%
        body_add_par(interpretasi, style = "Normal") %>%
        body_add_par("Histogram", style = "heading 2") %>%
        body_add_img(src = hist_file, width = 6, height = 4) %>%
        body_add_par("Boxplot", style = "heading 2") %>%
        body_add_img(src = box_file, width = 4, height = 4)
      
      print(doc, target = file)
    }
  )
  
  
  #======================= Peta ===================
  
  # Palet warna dinamis
  pal_peta <- reactive({
    req(input$var_peta)
    validate(need(input$var_peta %in% names(geo_map), "Variabel tidak valid"))
    colorNumeric(
      palette = "YlGn",
      domain = geo_map[[input$var_peta]],
      na.color = "transparent"
    )
  })
  
  # Render Leaflet
  output$peta_populasi <- renderLeaflet({
    req(input$var_peta)
    var <- input$var_peta
    df <- data_peta()
    pal <- colorNumeric("YlGn", domain = df[[var]], na.color = "transparent")
    
    leaflet(df) %>%
      addProviderTiles(providers$CartoDB.Positron) %>%
      addPolygons(
        fillColor = ~pal(df[[var]]),
        weight = 1,
        color = "white",
        opacity = 1,
        fillOpacity = 0.8,
        label = ~paste0(nmkab, ": ", round(df[[var]], 2)),
        labelOptions = labelOptions(direction = "auto", style = list("font-size" = "12px")),
        highlight = highlightOptions(weight = 2, color = "#666", fillOpacity = 0.9, bringToFront = TRUE)
      ) %>%
      addLegend(
        pal = pal,
        values = df[[var]],
        title = var,
        position = "bottomright",
        opacity = 0.7
      )
  })
  
  
  # Filter Peta 
  output$filter_kab_ui_peta <- renderUI({
    req(input$filter_prov_peta)
    if (input$filter_prov_peta == "ALL") return(NULL)
    
    kabupaten_terpilih <- geo_map %>%
      filter(nmprov == input$filter_prov_peta) %>%
      pull(nmkab) %>%
      unique() %>%
      sort()
    
    selectInput("filter_kab_peta", "Pilih Kabupaten:", 
                choices = c("ALL", kabupaten_terpilih),
                selected = "ALL")
  })
  
  # Filter data untuk peta
  data_peta <- reactive({
    req(input$var_peta)
    df <- geo_map
    
    if (!is.null(input$filter_prov_peta) && input$filter_prov_peta != "ALL") {
      df <- df %>% filter(nmprov == input$filter_prov_peta)
      if (!is.null(input$filter_kab_peta) && input$filter_kab_peta != "ALL") {
        df <- df %>% filter(nmkab == input$filter_kab_peta)
      }
    }
    df
  })
  

  # Interpretasi Peta Otomatis
  output$interpretasi_peta <- renderUI({
    req(input$var_peta)
    var <- input$var_peta
    
    df <- data_peta()  # Data yang sudah difilter prov/kab
    
    validate(
      need(var %in% names(df), "Variabel tidak tersedia"),
      need(is.numeric(df[[var]]), "Variabel bukan numerik"),
      need(nrow(df) > 0, "Tidak ada data yang ditampilkan")
    )
    
    nilai <- df[[var]]
    rata <- round(mean(nilai, na.rm = TRUE), 2)
    maks <- round(max(nilai, na.rm = TRUE), 2)
    minim <- round(min(nilai, na.rm = TRUE), 2)
    
    HTML(paste0(
      "<p style='text-align:justify'>",
      "Peta ini menunjukkan sebaran nilai variabel <b>", var, "</b> ",
      "di wilayah yang telah difilter. ",
      "Warna yang lebih gelap menunjukkan nilai yang lebih tinggi dari variabel tersebut. ",
      "Nilai rata-rata tercatat sebesar <b>", rata, "</b>, dengan nilai minimum <b>", minim,
      "</b> dan maksimum <b>", maks, "</b>. ",
      "Visualisasi ini berguna untuk mengenali daerah-daerah dengan konsentrasi nilai ekstrem.",
      "</p>"
    ))
  })
  
  
  output$unduh_peta_png <- downloadHandler(
    filename = function() {
      paste0("peta_", input$var_peta, "_", Sys.Date(), ".png")
    },
    content = function(file) {
      df <- data_peta()
      var <- input$var_peta
      pal <- colorNumeric("YlGn", domain = df[[var]], na.color = "transparent")
      
      map <- leaflet(df) %>%
        addProviderTiles(providers$CartoDB.Positron) %>%
        addPolygons(
          fillColor = ~pal(df[[var]]),
          weight = 1,
          color = "white",
          opacity = 1,
          fillOpacity = 0.8,
          label = ~paste0(nmkab, ": ", round(df[[var]], 2)),
          highlight = highlightOptions(weight = 2, color = "#666", fillOpacity = 0.9, bringToFront = TRUE)
        ) %>%
        addLegend(pal = pal, values = df[[var]], title = var, position = "bottomright")
      
      temp_html <- tempfile(fileext = ".html")
      saveWidget(map, file = temp_html, selfcontained = TRUE)
      webshot2::webshot(temp_html, file = file, vwidth = 1000, vheight = 800)
    }
  )
  
  output$unduh_interpretasi_peta <- downloadHandler(
    filename = function() {
      paste0("interpretasi_peta_", input$var_peta, "_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      var <- input$var_peta
      df <- data_peta()
      nilai <- df[[var]]
      rata <- round(mean(nilai, na.rm = TRUE), 2)
      maks <- round(max(nilai, na.rm = TRUE), 2)
      minim <- round(min(nilai, na.rm = TRUE), 2)
      
      kalimat <- paste0(
        "Peta ini menunjukkan sebaran nilai variabel '", var, "' di wilayah yang telah difilter. ",
        "Warna gelap merepresentasikan nilai tinggi. ",
        "Rata-rata = ", rata, ", minimum = ", minim, ", maksimum = ", maks, "."
      )
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Interpretasi Peta", style = "heading 1")
      doc <- body_add_par(doc, kalimat, style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  output$unduh_halaman_peta <- downloadHandler(
    filename = function() {
      paste0("Halaman-Peta_", input$var_peta, "_", Sys.Date(), ".docx")
    },
    content = function(file) {
      req(input$var_peta)
      
      var <- input$var_peta
      df <- data_peta()
      nilai <- df[[var]]
      
      # Hitung statistik
      rata <- round(mean(nilai, na.rm = TRUE), 2)
      maks <- round(max(nilai, na.rm = TRUE), 2)
      minim <- round(min(nilai, na.rm = TRUE), 2)
      
      # Interpretasi
      kalimat <- paste0(
        "Peta ini menunjukkan sebaran nilai variabel '", var, "' di wilayah yang telah difilter. ",
        "Warna gelap merepresentasikan nilai tinggi. ",
        "Rata-rata = ", rata, ", minimum = ", minim, ", maksimum = ", maks, "."
      )
      
      # Buat peta Leaflet dan simpan screenshot sementara
      pal <- colorNumeric("YlGn", domain = df[[var]], na.color = "transparent")
      
      map <- leaflet(df) %>%
        addProviderTiles(providers$CartoDB.Positron) %>%
        addPolygons(
          fillColor = ~pal(df[[var]]),
          weight = 1,
          color = "white",
          opacity = 1,
          fillOpacity = 0.8,
          label = ~paste0(nmkab, ": ", round(df[[var]], 2)),
          highlight = highlightOptions(weight = 2, color = "#666", fillOpacity = 0.9, bringToFront = TRUE)
        ) %>%
        addLegend(pal = pal, values = df[[var]], title = var, position = "bottomright")
      
      # Simpan sebagai HTML dan screenshot ke PNG
      temp_html <- tempfile(fileext = ".html")
      temp_png <- tempfile(fileext = ".png")
      
      saveWidget(map, file = temp_html, selfcontained = TRUE)
      webshot2::webshot(temp_html, file = temp_png, vwidth = 1000, vheight = 800)
      
      # Buat dokumen Word
      doc <- read_docx()
      doc <- doc %>%
        body_add_par("Halaman Peta", style = "heading 1") %>%
        body_add_par(paste0("Variabel: ", var), style = "heading 2") %>%
        body_add_par(paste0("Statistik Ringkas — Rata-rata: ", rata, 
                            ", Minimum: ", minim, ", Maksimum: ", maks), style = "Normal") %>%
        body_add_par("Interpretasi:", style = "heading 2") %>%
        body_add_par(kalimat, style = "Normal") %>%
        body_add_par("Visualisasi Peta", style = "heading 2") %>%
        body_add_img(temp_png, width = 6.5, height = 5.2)
      
      print(doc, target = file)
    }
  )
  
  #======================= Manajemen Data ===================
  #1. Reactive value untuk menyimpan hasil manajemen data
  hasil_manajemen <- reactiveVal(data)
  cutpoints_list <- reactiveVal(list())
  
  #2. Operasi manajemen data
  observeEvent(input$proses_kat, {
    req(input$var_kategorik, input$metode_kat)
    
    if (input$banyak_kategori <= 1) {
      showModal(modalDialog(
        title = "Jumlah Kategori Tidak Valid",
        "Jumlah kategori harus lebih dari 1. Silakan masukkan angka minimal 2.",
        easyClose = TRUE,
        footer = modalButton("Tutup")
      ))
      return()
    }
    
    data_baru <- hasil_manajemen()
    kolom_lama <- grep("_(SL|Q|J)$", names(data_baru), value = TRUE)
    data_baru <- data_baru[, !names(data_baru) %in% kolom_lama]
    
    cut_list <- list()
    
    for (var in input$var_kategorik) {
      var_data <- data[[var]]
      
      for (metode in input$metode_kat) {
        
        if (metode == "Interval Sama Lebar") {
          kategori <- cut(var_data,
                          breaks = input$banyak_kategori,
                          labels = 1:input$banyak_kategori,
                          include.lowest = TRUE)
          nama_baru <- paste0(var, "_SL")
          data_baru[[nama_baru]] <- as.numeric(as.character(kategori))
          cut_list[[nama_baru]] <- levels(cut(var_data,
                                              breaks = input$banyak_kategori,
                                              include.lowest = TRUE))
        }
        
        if (metode == "Kuartil") {
          q <- quantile(var_data, probs = seq(0, 1, length.out = input$banyak_kategori + 1), na.rm = TRUE)
          if (length(unique(q)) == length(q)) {
            kategori <- cut(var_data,
                            breaks = q,
                            labels = 1:input$banyak_kategori,
                            include.lowest = TRUE)
            nama_baru <- paste0(var, "_Q")
            data_baru[[nama_baru]] <- as.numeric(as.character(kategori))
            cut_list[[nama_baru]] <- levels(cut(var_data,
                                                breaks = q,
                                                include.lowest = TRUE))
          }
        }
        
        if (metode == "Natural Breaks") {
          jenks <- classInt::classIntervals(var_data, n = input$banyak_kategori, style = "jenks")
          brks <- jenks$brks
          kategori <- cut(var_data,
                          breaks = brks,
                          labels = 1:(length(brks) - 1),
                          include.lowest = TRUE)
          nama_baru <- paste0(var, "_J")
          data_baru[[nama_baru]] <- as.numeric(as.character(kategori))
          cut_list[[nama_baru]] <- levels(cut(var_data,
                                              breaks = brks,
                                              include.lowest = TRUE))
        }
        
      }
    }
    
    hasil_manajemen(data_baru)
    cutpoints_list(cut_list)
  })
  
  #3. Tabel hasil manajemen data
  output$tabel_manajemen_data <- DT::renderDataTable({
    data_tampil <- hasil_manajemen()
    
    kolom_kategori <- grep("_(SL|Q|J)$", names(data_tampil), value = TRUE)
    
    kolom_relevan <- c("nmkab", "LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED", kolom_kategori)
    kolom_relevan <- kolom_relevan[kolom_relevan %in% names(data_tampil)]
    
    is_num <- sapply(data_tampil, is.numeric)
    data_tampil[is_num] <- lapply(data_tampil[is_num], function(x) round(x, 5))
    
    DT::datatable(
      data_tampil[, kolom_relevan, drop = FALSE],
      options = list(scrollX = TRUE, pageLength = 5),
      rownames = FALSE,
      colnames = gsub("^nmkab$", "KAB", kolom_relevan)
    )
  })
  
  #4. Download Excel manajemen data
  output$unduh_kategori <- downloadHandler(
    filename = function() {
      paste0("data_kategorisasi_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      df <- hasil_manajemen()
      if (!is.data.frame(df)) {
        showModal(modalDialog(
          title = "Gagal Menyimpan",
          "Data manajemen bukan data.frame yang valid.",
          easyClose = TRUE,
          footer = modalButton("Tutup")
        ))
        return(NULL)
      }
      
      names(df)[names(df) == "nmkab"] <- "KAB"
      var_utama <- c("DISTRICTCODE", "KAB", "LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED")
      var_kat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
      variabel_simpan <- intersect(names(df), c(var_utama, var_kat))
      df_simpan <- df[, variabel_simpan]
      
      tryCatch({
        openxlsx::write.xlsx(df_simpan, file, rowNames = FALSE)
      }, error = function(e) {
        showModal(modalDialog(
          title = "Terjadi Kesalahan",
          paste("Gagal menyimpan file Excel:", e$message),
          easyClose = TRUE,
          footer = modalButton("Tutup")
        ))
      })
    }
  )
  
  #5. Cut Point Info
  output$cutpoint_info <- renderUI({
    req(cutpoints_list())
    cp <- cutpoints_list()
    tagList(
      lapply(names(cp), function(nama_var) {
        tags$div(
          tags$b(nama_var), tags$br(),
          paste(cp[[nama_var]], collapse = " - ")
        )
      })
    )
  })
  
  #6. Interpretasi Cutpoint
  output$interpretasi_cutpoint <- renderUI({
    req(cutpoints_list())
    cp <- cutpoints_list()
    if (length(cp) == 0) return(NULL)
    
    vars_sl <- names(cp)[grepl("_sl$", names(cp), ignore.case = TRUE)]
    vars_q  <- names(cp)[grepl("_q$", names(cp), ignore.case = TRUE)]
    vars_j  <- names(cp)[grepl("_j$", names(cp), ignore.case = TRUE)]
    
    isi <- list()
    
    if (length(vars_sl) > 0) {
      total_kat <- unique(sapply(cp[vars_sl], length))
      teks_sl <- paste0(
        "Variabel ", paste(vars_sl, collapse = ", "),
        " dikategorikan menggunakan metode Sama Lebar menjadi ", total_kat[1], " kategori. ",
        "Lebar interval dihitung dengan formula (Max - Min) / k, di mana k adalah jumlah kategori. ",
        "Metode ini membagi rentang nilai menjadi bagian yang sama panjang tanpa mempertimbangkan jumlah data dalam tiap kategori. ",
        "Karena itu, metode ini cukup sensitif terhadap pencilan."
      )
      isi[[length(isi)+1]] <- tags$p(style = "text-align: justify;", teks_sl)
    }
    
    if (length(vars_q) > 0) {
      total_kat <- unique(sapply(cp[vars_q], length))
      teks_q <- paste0(
        "Variabel ", paste(vars_q, collapse = ", "),
        " dikategorikan menggunakan metode Kuartil menjadi ", total_kat[1], " kategori. ",
        "Interval dibentuk berdasarkan nilai kuantil ke-i yang membagi data menjadi kategori dengan jumlah observasi yang relatif seimbang. ",
        "Metode ini lebih tahan terhadap pencilan karena tidak bergantung pada nilai ekstrem, melainkan pada posisi data dalam distribusi."
      )
      isi[[length(isi)+1]] <- tags$p(style = "text-align: justify;", teks_q)
    }
    
    if (length(vars_j) > 0) {
      total_kat <- unique(sapply(cp[vars_j], length))
      teks_j <- paste0(
        "Variabel ", paste(vars_j, collapse = ", "),
        " dikategorikan menggunakan metode Natural Breaks (Jenks) menjadi ", total_kat[1], " kategori. ",
        "Metode ini membagi data dengan cara meminimalkan variasi di dalam setiap kelompok dan memaksimalkan perbedaan antar kelompok. ",
        "Natural Breaks sangat cocok untuk data spasial atau distribusi yang tidak merata, karena membentuk batasan berdasarkan pola alami dalam data."
      )
      isi[[length(isi)+1]] <- tags$p(style = "text-align: justify;", teks_j)
    }
    
    tagList(isi)
  })
  
  output$unduh_interpretasi_kat <- downloadHandler(
    filename = function() {
      paste0("interpretasi_kategorisasi_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      cp <- cutpoints_list()
      if (length(cp) == 0) return()
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Interpretasi Hasil Kategorisasi", style = "heading 1")
      
      if (length(cp) == 0) {
        doc <- body_add_par(doc, "Belum ada hasil kategorisasi.", style = "Normal")
      } else {
        vars_sl <- names(cp)[grepl("_sl$", names(cp), ignore.case = TRUE)]
        vars_q  <- names(cp)[grepl("_q$", names(cp), ignore.case = TRUE)]
        vars_j  <- names(cp)[grepl("_j$", names(cp), ignore.case = TRUE)]
        
        if (length(vars_sl) > 0) {
          total_kat <- unique(sapply(cp[vars_sl], length))
          teks_sl <- paste0(
            "Variabel ", paste(vars_sl, collapse = ", "),
            " dikategorikan dengan metode Sama Lebar menjadi ", total_kat[1], " kategori. ",
            "Interval dibuat sama panjang dan sensitif terhadap pencilan."
          )
          doc <- body_add_par(doc, teks_sl, style = "Normal")
        }
        
        if (length(vars_q) > 0) {
          total_kat <- unique(sapply(cp[vars_q], length))
          teks_q <- paste0(
            "Variabel ", paste(vars_q, collapse = ", "),
            " dikategorikan dengan metode Kuartil menjadi ", total_kat[1], " kategori. ",
            "Interval berdasarkan distribusi kuantil dan lebih tahan terhadap pencilan."
          )
          doc <- body_add_par(doc, teks_q, style = "Normal")
        }
        
        if (length(vars_j) > 0) {
          total_kat <- unique(sapply(cp[vars_j], length))
          teks_j <- paste0(
            "Variabel ", paste(vars_j, collapse = ", "),
            " dikategorikan dengan metode Natural Breaks menjadi ", total_kat[1], " kategori. ",
            "Interval berdasarkan variasi alami dalam data."
          )
          doc <- body_add_par(doc, teks_j, style = "Normal")
        }
      }
      
      print(doc, target = file)
    }
  )
  
  output$unduh_halaman_manajemen <- downloadHandler(
    filename = function() {
      paste0("Halaman-Manajemen_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      df <- hasil_manajemen()
      cp <- cutpoints_list()
      if (is.null(df) || length(cp) == 0) return()
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Halaman Manajemen Data", style = "heading 1")
      
      # === Tabel hasil manajemen data ===
      kolom_kategori <- grep("_(SL|Q|J)$", names(df), value = TRUE)
      kolom_relevan <- c("nmkab", "LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", 
                         "POVERTY", "NOELECTRIC", "RENTED", kolom_kategori)
      kolom_relevan <- kolom_relevan[kolom_relevan %in% names(df)]
      
      # Tabel ringkas max 10 baris (atau semua jika <=10)
      df_tabel <- df[, kolom_relevan]
      df_tabel <- head(df_tabel, 10)
      names(df_tabel)[names(df_tabel) == "nmkab"] <- "KAB"
      
      doc <- body_add_par(doc, "Contoh Tabel Hasil Kategorisasi (10 baris pertama):", style = "heading 2")
      doc <- body_add_table(doc, df_tabel, style = "table_template") # gunakan style default jika tidak ada template
      
      # === Cut point info ===
      doc <- body_add_par(doc, "Cut Point Hasil Kategorisasi", style = "heading 2")
      for (nama_var in names(cp)) {
        cut_vals <- cp[[nama_var]]
        doc <- body_add_par(doc, paste0(nama_var, ": ", paste(cut_vals, collapse = " | ")), style = "Normal")
      }
      
      # === Interpretasi Otomatis ===
      doc <- body_add_par(doc, "Interpretasi", style = "heading 2")
      
      vars_sl <- names(cp)[grepl("_sl$", names(cp), ignore.case = TRUE)]
      vars_q  <- names(cp)[grepl("_q$", names(cp), ignore.case = TRUE)]
      vars_j  <- names(cp)[grepl("_j$", names(cp), ignore.case = TRUE)]
      
      if (length(vars_sl) > 0) {
        total_kat <- unique(sapply(cp[vars_sl], length))
        teks_sl <- paste0(
          "Variabel ", paste(vars_sl, collapse = ", "),
          " dikategorikan menggunakan metode Sama Lebar menjadi ", total_kat[1], " kategori. ",
          "Metode ini membagi rentang nilai menjadi bagian yang sama panjang tanpa mempertimbangkan jumlah data dalam tiap kategori. ",
          "Karena itu, metode ini cukup sensitif terhadap pencilan."
        )
        doc <- body_add_par(doc, teks_sl, style = "Normal")
      }
      
      if (length(vars_q) > 0) {
        total_kat <- unique(sapply(cp[vars_q], length))
        teks_q <- paste0(
          "Variabel ", paste(vars_q, collapse = ", "),
          " dikategorikan menggunakan metode Kuartil menjadi ", total_kat[1], " kategori. ",
          "Interval dibentuk berdasarkan nilai kuantil yang membagi data menjadi jumlah yang relatif seimbang. ",
          "Metode ini lebih tahan terhadap pencilan karena berdasarkan posisi data, bukan nilai ekstrem."
        )
        doc <- body_add_par(doc, teks_q, style = "Normal")
      }
      
      if (length(vars_j) > 0) {
        total_kat <- unique(sapply(cp[vars_j], length))
        teks_j <- paste0(
          "Variabel ", paste(vars_j, collapse = ", "),
          " dikategorikan menggunakan metode Natural Breaks (Jenks) menjadi ", total_kat[1], " kategori. ",
          "Metode ini meminimalkan variasi di dalam kelompok dan memaksimalkan variasi antar kelompok, cocok untuk distribusi tidak merata."
        )
        doc <- body_add_par(doc, teks_j, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  #======================= Uji Asumsi : Normalitas ===================
  # Uji Shapiro-Wilk
  output$hasil_shapiro <- renderPrint({
    req(input$var_normalitas)
    x <- data[[input$var_normalitas]]
    shapiro.test(x)
  })
  
  # Histogram dan Q-Q Plot
  output$plot_normalitas <- renderPlot({
    req(input$var_normalitas)
    x <- data[[input$var_normalitas]]
    hist(x, breaks = 30, col = "#4C6C29", border = "white",
         main = paste("Histogram", input$var_normalitas),
         xlab = input$var_normalitas)
  })
  
  output$qqplot_normalitas <- renderPlot({
    req(input$var_normalitas)
    x <- data[[input$var_normalitas]]
    qqnorm(x, main = paste("Q-Q Plot", input$var_normalitas))
    qqline(x, col = "red", lwd = 2)
  })
  
  # Interpretasi Otomatis
  output$interpretasi_normalitas <- renderUI({
    req(input$var_normalitas)
    x <- data[[input$var_normalitas]]
    hasil <- shapiro.test(x)
    pval <- round(hasil$p.value, 4)
    
    kesimpulan <- if (pval >= 0.05) {
      "karena p-value ≥ 0.05, maka tidak cukup bukti untuk menolak H0. Data terdistribusi normal."
    } else {
      "karena p-value < 0.05, maka terdapat bukti untuk menolak H0. Data tidak terdistribusi normal."
    }
    
    HTML(paste0(
      "<p style='text-align:justify'>",
      "Hasil uji Shapiro-Wilk pada variabel <b>", input$var_normalitas, "</b> menunjukkan nilai p-value sebesar <b>", pval, "</b>. ",
      "Dengan tingkat signifikansi 5%, ", kesimpulan,
      "</p>"
    ))
  })
  
  output$unduh_normalitas <- downloadHandler(
    filename = function() {
      paste0("hasil_uji_normalitas_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      library(ggplot2)
      
      var <- input$var_normalitas
      x <- data[[var]]
      hasil <- shapiro.test(x)
      
      # Histogram
      p1 <- ggplot(data.frame(x), aes(x)) +
        geom_histogram(bins = 30, fill = "#4C6C29", color = "white") +
        labs(title = paste("Histogram", var), x = var, y = "Frekuensi") +
        theme_minimal()
      
      # Simpan Q-Q Plot secara manual ke PNG
      tmp_q <- tempfile(fileext = ".png")
      png(tmp_q, width = 800, height = 500, res = 120)
      qqnorm(x, main = paste("Q-Q Plot", var))
      qqline(x, col = "red", lwd = 2)
      dev.off()
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Hasil Uji Normalitas (Shapiro-Wilk)", style = "heading 1")
      doc <- body_add_par(doc, paste("Variabel:", var), style = "Normal")
      doc <- body_add_par(doc, paste("Statistik:", round(hasil$statistic, 4)), style = "Normal")
      doc <- body_add_par(doc, paste("p-value:", round(hasil$p.value, 4)), style = "Normal")
      
      if (hasil$p.value >= 0.05) {
        doc <- body_add_par(doc, "Kesimpulan: Tidak cukup bukti untuk menolak H0. Data kemungkinan berdistribusi normal.", style = "Normal")
      } else {
        doc <- body_add_par(doc, "Kesimpulan: Terdapat bukti untuk menolak H0. Data kemungkinan tidak normal.", style = "Normal")
      }
      
      doc <- body_add_par(doc, " ", style = "Normal")
      doc <- body_add_par(doc, "Histogram:", style = "heading 2")
      doc <- body_add_gg(doc, value = p1, width = 6, height = 4)
      
      doc <- body_add_par(doc, "Q-Q Plot:", style = "heading 2")
      doc <- body_add_img(doc, src = tmp_q, width = 6, height = 4)
      
      print(doc, target = file)
    }
  )
  
  
  #======================= Uji Asumsi : Homogenitas ===================
  # Filter variabel kategorik
  output$pilihan_group_uji <- renderUI({
    df <- hasil_manajemen()
    var_kat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(var_kat) == 0) {
      return(helpText("Tidak ditemukan variabel kategorik. Silakan kategorisasi terlebih dahulu."))
    }
    
    selectInput("group_uji", "Variabel Kategorik (Hasil Kategorisasi):",
                choices = var_kat,
                selected = var_kat[1])
  })
  

  output$hasil_levene <- renderPrint({
    req(input$var_homogenitas, input$group_uji)
    
    df <- hasil_manajemen()
    
    if (!(input$var_homogenitas %in% names(df))) {
      cat("Variabel numerik tidak ditemukan dalam data.\n")
      return()
    }
    if (!(input$group_uji %in% names(df))) {
      cat("Variabel kategorik tidak ditemukan dalam data.\n")
      return()
    }
    
    n_kategori <- length(unique(df[[input$group_uji]]))
    if (n_kategori < 2) {
      cat("Variabel kategorik harus memiliki minimal 2 kategori.\n")
      return()
    }
    
    df_valid <- df %>%
      dplyr::select(all_of(c(input$var_homogenitas, input$group_uji))) %>%
      dplyr::filter(!is.na(.data[[input$var_homogenitas]]) & !is.na(.data[[input$group_uji]]))
    
    if (nrow(df_valid) == 0) {
      cat("Tidak ada data yang valid untuk uji Levene setelah menghapus NA.\n")
      return()
    }
    
    # Kode asli
    formula <- as.formula(paste0(input$var_homogenitas, " ~ as.factor(", input$group_uji, ")"))
    car::leveneTest(formula, data = df)
  })
  
  
  #interpretasi levene
  output$interpretasi_levene <- renderUI({
    req(input$var_homogenitas, input$group_uji)
    df <- hasil_manajemen()
    
    if (!(input$var_homogenitas %in% names(df)) || !(input$group_uji %in% names(df))) {
      return(HTML("<p style='color:red;'>Variabel tidak ditemukan dalam data.</p>"))
    }
    
    n_kategori <- length(unique(df[[input$group_uji]]))
    if (n_kategori < 2) {
      return(HTML("<p style='color:red;'>Variabel kategorik harus memiliki minimal 2 kategori.</p>"))
    }
    
    df_valid <- df %>%
      dplyr::select(all_of(c(input$var_homogenitas, input$group_uji))) %>%
      dplyr::filter(!is.na(.data[[input$var_homogenitas]]) & !is.na(.data[[input$group_uji]]))
    
    if (nrow(df_valid) == 0) {
      return(HTML("<p style='color:red;'>Tidak ada data yang valid untuk interpretasi setelah menghapus NA.</p>"))
    }
    
    # Analisis
    formula <- as.formula(paste0(input$var_homogenitas, " ~ as.factor(", input$group_uji, ")"))
    hasil <- car::leveneTest(formula, data = df_valid)
    pval <- hasil$`Pr(>F)`[1]
    
    kesimpulan <- if (pval < 0.05) {
      "Karena p-value < 0.05, maka H₀ ditolak. Terdapat perbedaan varians yang signifikan antar kelompok. Data tidak homogen."
    } else {
      "Karena p-value ≥ 0.05, maka H₀ tidak ditolak. Tidak terdapat perbedaan varians yang signifikan antar kelompok. Data dapat dianggap homogen."
    }
    
    HTML(paste0(
      "<p style='text-align:justify'>",
      "Uji Levene digunakan untuk menguji kesamaan varians dari variabel <b>", input$var_homogenitas, "</b> ",
      "berdasarkan kategori dalam <b>", input$group_uji, "</b>. ",
      "Hasil uji menunjukkan p-value sebesar <b>", round(pval, 4), "</b>. ",
      kesimpulan,
      "</p>"
    ))
  })
  
  # Box Plot: Uji homogenitas
  output$plot_levene <- renderPlot({
    req(input$var_homogenitas, input$group_uji)
    df <- hasil_manajemen()
    
    if (!(input$var_homogenitas %in% names(df)) || !(input$group_uji %in% names(df))) return()
    
    df_valid <- df %>%
      dplyr::select(all_of(c(input$var_homogenitas, input$group_uji))) %>%
      dplyr::filter(!is.na(.data[[input$var_homogenitas]]) & !is.na(.data[[input$group_uji]]))
    
    if (nrow(df_valid) == 0) return()
    
    ggplot(df_valid, aes(x = factor(.data[[input$group_uji]]), y = .data[[input$var_homogenitas]])) +
      geom_boxplot(fill = "#2e7d32", color = "black") +
      labs(x = input$group_uji, y = input$var_homogenitas) +
      theme_minimal()
    
  })
  
  output$unduh_levene <- downloadHandler(
    filename = function() {
      paste0("hasil_uji_homogenitas_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      library(ggplot2)
      
      df <- hasil_manajemen()
      var_y <- input$var_homogenitas
      var_kat <- input$group_uji
      
      if (!(var_y %in% names(df)) || !(var_kat %in% names(df))) return()
      
      df_valid <- df %>%
        dplyr::select(all_of(c(var_y, var_kat))) %>%
        dplyr::filter(!is.na(.data[[var_y]]) & !is.na(.data[[var_kat]]))
      
      if (nrow(df_valid) == 0) return()
      
      # Hasil uji
      formula <- as.formula(paste0(var_y, " ~ as.factor(", var_kat, ")"))
      hasil <- car::leveneTest(formula, data = df_valid)
      pval <- round(hasil$`Pr(>F)`[1], 4)
      
      kesimpulan <- if (pval < 0.05) {
        "Karena p-value < 0.05, maka H₀ ditolak. Terdapat perbedaan varians yang signifikan antar kelompok. Data tidak homogen."
      } else {
        "Karena p-value ≥ 0.05, maka H₀ tidak ditolak. Tidak terdapat perbedaan varians yang signifikan antar kelompok. Data dapat dianggap homogen."
      }
      
      # Plot boxplot
      plot_box <- ggplot(df_valid, aes(x = factor(.data[[var_kat]]), y = .data[[var_y]])) +
        geom_boxplot(fill = "#2e7d32", color = "black") +
        labs(title = paste("Boxplot", var_y, "berdasarkan", var_kat),
             x = var_kat, y = var_y) +
        theme_minimal()
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Hasil Uji Homogenitas (Levene)", style = "heading 1")
      doc <- body_add_par(doc, paste("Variabel target:", var_y), style = "Normal")
      doc <- body_add_par(doc, paste("Variabel kategorik:", var_kat), style = "Normal")
      doc <- body_add_par(doc, paste("p-value:", pval), style = "Normal")
      doc <- body_add_par(doc, kesimpulan, style = "Normal")
      
      doc <- body_add_par(doc, " ", style = "Normal")
      doc <- body_add_par(doc, "Visualisasi Boxplot:", style = "heading 2")
      doc <- body_add_gg(doc, value = plot_box, width = 6, height = 4)
      
      print(doc, target = file)
    }
  )
  
  
  # ======================= Uji Beda Rata-rata ===================
  observe({
    df <- hasil_manajemen()  # Sudah mencakup hasil kategorisasi
    var_numerik <- c("LOWEDU", "ILLITERATE", "POVERTY", "NOELECTRIC",
                     "ELDERLY", "FHEAD", "FAMILYSIZE", "RENTED") 
    var_kat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    updateSelectInput(session, "var_rerata1", choices = var_numerik)
    updateSelectInput(session, "var_rerata2", choices = var_numerik)
    updateSelectInput(session, "group_rerata", choices = var_kat)
    updateSelectInput(session, "paired_var1", choices = var_numerik)
    updateSelectInput(session, "paired_var2", choices = var_numerik)
  })
  
  hasil_uji_rerata <- eventReactive(input$proses_uji_rerata, {
    req(input$jenis_uji_rerata)
    
    if (input$jenis_uji_rerata == "satu" || input$jenis_uji_rerata == "pasangan") {
      df <- data  # Gunakan data asli untuk satu rata-rata & pasangan
    } else if (input$jenis_uji_rerata == "dua") {
      df <- hasil_manajemen()  # Hanya untuk uji dua rata-rata
      if (is.null(df) || nrow(df) == 0) {
        showModal(modalDialog(
          title = "Peringatan",
          "Silahkan lakukan kategorisasi data terlebih dahulu di menu Manajemen Data.",
          easyClose = TRUE,
          footer = NULL
        ))
        return(NULL)
      }
    }
    
    # ==== Uji Satu Rata-rata ====
    if (input$jenis_uji_rerata == "satu") {
      x <- df[[input$var_rerata1]]
      x <- na.omit(x)
      return(t.test(x, mu = input$mu, alternative = input$arah_uji1))
      
      # ==== Uji Dua Rata-rata ====
    } else if (input$jenis_uji_rerata == "dua") {
      x <- df[[input$var_rerata2]]
      g <- df[[input$group_rerata]]
      df_valid <- data.frame(x, g) %>% filter(!is.na(x) & !is.na(g))
      
      n_kat <- length(unique(df_valid$g))
      if (n_kat != 2) {
        showModal(modalDialog(
          title = "Peringatan",
          "Variabel grup harus memiliki tepat 2 kategori. Silakan ubah melalui menu Manajemen Data.",
          easyClose = TRUE,
          footer = NULL
        ))
        return(NULL)
      }
      
      return(t.test(x ~ g, data = df_valid,
                    var.equal = as.logical(input$equal_var),
                    alternative = input$arah_uji2))
      
      # ==== Uji Berpasangan ====
    } else if (input$jenis_uji_rerata == "pasangan") {
      # Validasi: tidak tersedia pasangan logis → beri peringatan
      showModal(modalDialog(
        title = "Informasi",
        "Data saat ini tidak mendukung uji rata-rata berpasangan karena tidak ada variabel sebelum/sesudah perlakuan yang logis.",
        easyClose = TRUE,
        footer = NULL
      ))
      return(NULL)
    }
  })
  
  output$hasil_uji_rerata <- renderPrint({
    req(hasil_uji_rerata())  # hanya render jika hasil tidak NULL
    hasil_uji_rerata()
  })
  
  output$interpretasi_uji_rerata <- renderUI({
    req(hasil_uji_rerata())
    hasil <- hasil_uji_rerata()
    pval <- hasil$p.value
    stat <- round(hasil$statistic, 3)
    ci <- paste0("[", signif(hasil$conf.int[1], 4), ", ", signif(hasil$conf.int[2], 4), "]")
    
    jenis <- switch(input$jenis_uji_rerata,
                    "satu" = "Uji Satu Rata-rata",
                    "dua" = "Uji Dua Rata-rata Independen",
                    "pasangan" = "Uji Rata-rata Berpasangan")
    
    arah_input <- input[[paste0("arah_uji", switch(input$jenis_uji_rerata, satu = "1", dua = "2", pasangan = "3"))]]
    
    arah <- switch(arah_input,
                   "two.sided" = "Dua Arah (μ₁ ≠ μ₂)",
                   "less" = "Satu Arah Kiri (μ₁ < μ₂)",
                   "greater" = "Satu Arah Kanan (μ₁ > μ₂)")
    
    hipotesis <- switch(arah_input,
                        "two.sided" = HTML("H<sub>0</sub>: μ₁ = μ₂ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: μ₁ ≠ μ₂"),
                        "less"      = HTML("H<sub>0</sub>: μ₁ ≥ μ₂ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: μ₁ < μ₂"),
                        "greater"   = HTML("H<sub>0</sub>: μ₁ ≤ μ₂ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: μ₁ > μ₂")
    )
    
    interpretasi_teks <- if (pval < 0.05) {
      "Hasil pengujian menunjukkan bahwa terdapat perbedaan yang signifikan secara statistik antara kelompok yang diuji. Artinya, terdapat cukup bukti untuk menolak hipotesis nol (H₀) dan menyimpulkan bahwa rata-rata kelompok berbeda."
    } else {
      "Hasil pengujian menunjukkan bahwa tidak terdapat perbedaan yang signifikan secara statistik. Artinya, tidak cukup bukti untuk menolak hipotesis nol (H₀), sehingga tidak dapat disimpulkan bahwa rata-rata kelompok berbeda."
    }
    
    kesimpulan <- if (pval < 0.05) {
      "<span style='color: #2e7d32; font-weight: 600;'>Terdapat perbedaan yang signifikan secara statistik.</span>"
    } else {
      "<span style='color: #c62828; font-weight: 600;'>Tidak terdapat perbedaan yang signifikan secara statistik.</span>"
    }
    
    catatan <- "<em>Catatan:</em>  <strong>Salah arah uji</strong> dapat menyebabkan kesalahan penarikan kesimpulan."
    
    HTML(paste0(
      "<div style='padding: 10px; text-align: justify;'>",
      "<p><strong>Jenis Uji:</strong> ", jenis, "<br/>",
      "<strong>Arah Uji:</strong> ", arah, "<br/>",
      "<strong>Statistik Uji (t):</strong> ", stat, "<br/>",
      "<strong>p-value:</strong> ", signif(pval, 4), "<br/>",
      "<strong>Confidence Interval:</strong> ", ci, "</p>",
      "<p><strong>Hipotesis:</strong><br/>", hipotesis, "</p>",
      "<p><strong>Kesimpulan:</strong><br/>", kesimpulan, "</p>",
      "<p><strong>Interpretasi:</strong> ", interpretasi_teks, "</p>",
      "<p style='font-size: 13px; color: #666;'>", catatan, "</p>",
      "</div>"
    ))
  })
  
  
  output$group_rerata_ui <- renderUI({
    df <- hasil_manajemen()
    var_kat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(var_kat) == 0) {
      return(helpText("Belum ada variabel kategorik. Silahkan lakukan kategorisasi di menu Manajemen Data. 
                      Variabel harus dikategorikan menjadi 2 kategori sebelum dilakukan uji"))
    }
    
    selectInput("group_rerata", "Pilih Variabel Grup:", choices = var_kat)
  })
  
  output$unduh_uji_rerata_docx <- downloadHandler(
    filename = function() {
      paste0("hasil_uji_beda_rerata_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      hasil <- hasil_uji_rerata()
      if (is.null(hasil)) return()
      
      jenis <- switch(input$jenis_uji_rerata,
                      "satu" = "Uji Satu Rata-rata",
                      "dua" = "Uji Dua Rata-rata Independen",
                      "pasangan" = "Uji Rata-rata Berpasangan")
      
      arah_input <- input[[paste0("arah_uji", switch(input$jenis_uji_rerata, satu = "1", dua = "2", pasangan = "3"))]]
      
      arah <- switch(arah_input,
                     "two.sided" = "Dua Arah (μ₁ ≠ μ₂)",
                     "less" = "Satu Arah Kiri (μ₁ < μ₂)",
                     "greater" = "Satu Arah Kanan (μ₁ > μ₂)")
      
      hipotesis <- switch(arah_input,
                          "two.sided" = "H₀: μ₁ = μ₂, H₁: μ₁ ≠ μ₂",
                          "less"      = "H₀: μ₁ ≥ μ₂, H₁: μ₁ < μ₂",
                          "greater"   = "H₀: μ₁ ≤ μ₂, H₁: μ₁ > μ₂")
      
      pval <- signif(hasil$p.value, 4)
      stat <- round(hasil$statistic, 3)
      ci <- paste0("[", signif(hasil$conf.int[1], 4), ", ", signif(hasil$conf.int[2], 4), "]")
      
      interpretasi <- if (pval < 0.05) {
        "Terdapat perbedaan yang signifikan secara statistik."
      } else {
        "Tidak terdapat perbedaan yang signifikan secara statistik."
      }
      
      catatan <- "Catatan: Salah arah uji dapat menyebabkan kesalahan penarikan kesimpulan."
      
      # Buat dokumen
      doc <- read_docx()
      doc <- body_add_par(doc, "Hasil Uji Beda Rata-rata", style = "heading 1")
      
      doc <- body_add_par(doc, paste("Jenis Uji:", jenis), style = "Normal")
      doc <- body_add_par(doc, paste("Arah Uji:", arah), style = "Normal")
      doc <- body_add_par(doc, paste("Hipotesis:", hipotesis), style = "Normal")
      doc <- body_add_par(doc, paste("Statistik Uji (t):", stat), style = "Normal")
      doc <- body_add_par(doc, paste("p-value:", pval), style = "Normal")
      doc <- body_add_par(doc, paste("Confidence Interval:", ci), style = "Normal")
      doc <- body_add_par(doc, paste("Kesimpulan:", interpretasi), style = "Normal")
      doc <- body_add_par(doc, catatan, style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  
  # ======================= Uji Proporsi ===============================
  observe({
    df <- hasil_manajemen()
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(kandidat) == 0) {
      updateSelectInput(session, "var_kat_proporsi", choices = character(0))
      updateSelectInput(session, "var_kat_proporsi2", choices = character(0))
      return()
    }
    
    kandidat_2kategori <- kandidat[
      sapply(kandidat, function(kol) length(na.omit(unique(df[[kol]]))) == 2)
    ]
    
    updateSelectInput(session, "var_kat_proporsi", choices = kandidat_2kategori)
    updateSelectInput(session, "var_kat_proporsi2", choices = kandidat_2kategori)
  })
  
  output$nilai_kat_proporsi_ui <- renderUI({
    df <- hasil_manajemen()
    
    if (is.null(input$var_kat_proporsi) || !(input$var_kat_proporsi %in% names(df))) {
      return(helpText("Belum ada variabel kategorik. Silakan lakukan kategorisasi terlebih dahulu."))
    }
    
    nilai <- sort(unique(na.omit(df[[input$var_kat_proporsi]])))
    if (length(nilai) != 2) {
      return(helpText("Variabel kategorik harus memiliki tepat 2 kategori."))
    }
    
    checkboxGroupInput("nilai_kat_proporsi", "Kategori 'Sukses':", choices = nilai, selected = tail(nilai, 1))
  })
  
  output$nilai_kat_proporsi2_ui <- renderUI({
    df <- hasil_manajemen()
    
    if (is.null(input$var_kat_proporsi2) || !(input$var_kat_proporsi2 %in% names(df))) {
      return(helpText("Belum ada variabel kategorik. Silakan lakukan kategorisasi terlebih dahulu."))
    }
    
    nilai <- sort(unique(na.omit(df[[input$var_kat_proporsi2]])))
    if (length(nilai) != 2) {
      return(helpText("Variabel kategorik harus memiliki tepat 2 kategori."))
    }
    
    checkboxGroupInput("nilai_kat_proporsi2", "Kategori 'Sukses':", choices = nilai, selected = tail(nilai, 1))
  })
  
  output$group_proporsi_ui <- renderUI({
    df <- hasil_manajemen()
    
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(kandidat) == 0) {
      return(helpText(""))
    }
    
    ok <- vapply(kandidat, function(kol) {
      length(unique(na.omit(df[[kol]]))) == 2
    }, logical(1))
    
    kandidat_2lvl <- kandidat[ok]
    
    if (length(kandidat_2lvl) == 0) {
      return(helpText("Tidak ditemukan variabel grup dengan tepat 2 kategori. Silakan buat 2-level melalui Manajemen Data."))
    }
    
    selectInput("group_proporsi", "Variabel Grup (2 Kategori):", choices = kandidat_2lvl)
  })
  
  hasil_uji_proporsi <- eventReactive(input$proses_uji_proporsi, {
    df <- hasil_manajemen()
    alpha <- input$alpha_proporsi
    
    if (input$jenis_uji_proporsi == "satu") {
      req(input$var_kat_proporsi, input$nilai_kat_proporsi)
      var <- input$var_kat_proporsi
      sukses <- input$nilai_kat_proporsi
      
      if (length(unique(na.omit(df[[var]]))) != 2) {
        showModal(modalDialog(title = "Peringatan", "Variabel harus memiliki tepat 2 kategori untuk uji proporsi.", easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      x <- sum(df[[var]] %in% sukses, na.rm = TRUE)
      n <- sum(!is.na(df[[var]]))
      if (n == 0) return(NULL)
      
      metode <- if (n < 30 || x < 5 || x > n - 5) "binom" else "prop"
      hasil <- if (metode == "binom") {
        binom.test(x = x, n = n, p = input$nilai_p0_proporsi, alternative = input$arah_uji_proporsi)
      } else {
        prop.test(x = x, n = n, p = input$nilai_p0_proporsi, alternative = input$arah_uji_proporsi, correct = FALSE)
      }
      
      return(list(jenis = "satu", hasil = hasil, x = x, n = n, metode = metode))
      
    } else {
      req(input$var_kat_proporsi2, input$nilai_kat_proporsi2, input$group_proporsi)
      var <- input$var_kat_proporsi2
      sukses <- input$nilai_kat_proporsi2
      grup <- input$group_proporsi
      
      if (length(unique(na.omit(df[[var]]))) != 2) {
        showModal(modalDialog(title = "Peringatan", "Variabel harus memiliki tepat 2 kategori untuk uji proporsi.", easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      df_valid <- df %>% filter(!is.na(.data[[var]]) & !is.na(.data[[grup]])) %>% mutate(status = ifelse(.data[[var]] %in% sukses, 1, 0))
      tab <- table(df_valid[[grup]], df_valid$status)
      
      if (nrow(tab) != 2 || ncol(tab) != 2) {
        showModal(modalDialog(title = "Peringatan", "Data tidak memenuhi format 2x2 untuk uji proporsi dua kelompok.", easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      x <- c(tab[1, "1"], tab[2, "1"])
      n <- c(sum(tab[1, ]), sum(tab[2, ]))
      metode <- if (any(x < 5 | (n - x) < 5)) "fisher" else "prop"
      
      hasil <- if (metode == "fisher") {
        fisher.test(matrix(c(tab[1, "1"], tab[1, "0"], tab[2, "1"], tab[2, "0"]), nrow = 2))
      } else {
        prop.test(x = x, n = n, alternative = input$arah_uji_proporsi2, correct = FALSE)
      }
      
      return(list(jenis = "dua", hasil = hasil, x = x, n = n, metode = metode))
    }
  })
  
  
  output$hasil_uji_proporsi <- renderPrint({
    h <- hasil_uji_proporsi()
    req(h)
    h$hasil
  })
  
  output$interpretasi_uji_proporsi <- renderUI({
    h <- hasil_uji_proporsi()
    req(h)
    hasil <- h$hasil
    ci <- if (!is.null(hasil$conf.int)) {
      paste0("[", signif(hasil$conf.int[1], 4), ", ", signif(hasil$conf.int[2], 4), "]")
    } else {
      "Tidak tersedia (uji Fisher tidak menghasilkan interval kepercayaan)"
    }
    
    pval <- hasil$p.value
    stat <- if (!is.null(hasil$statistic) && is.numeric(hasil$statistic)) {
      round(hasil$statistic, 3)
    } else {
      "Tidak tersedia (uji Fisher tidak menghasilkan statistik uji)"
    }
    
    alpha <- input$alpha_proporsi
    
    jenis <- if (h$jenis == "satu") "Uji Proporsi Satu Kelompok" else "Uji Proporsi Dua Kelompok"
    
    arah_input <- if (h$jenis == "satu") input$arah_uji_proporsi else input$arah_uji_proporsi2
    arah <- switch(arah_input,
                   "two.sided" = "Dua Arah (p ≠ p₀)",
                   "less"      = "Satu Arah Kiri (p < p₀ atau p₁ < p₂)",
                   "greater"   = "Satu Arah Kanan (p > p₀ atau p₁ > p₂)")
    
    hipotesis <- if (h$jenis == "satu") {
      switch(arah_input,
             "two.sided" = HTML("H<sub>0</sub>: p = p₀ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: p ≠ p₀"),
             "less"      = HTML("H<sub>0</sub>: p ≥ p₀ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: p < p₀"),
             "greater"   = HTML("H<sub>0</sub>: p ≤ p₀ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: p > p₀"))
    } else {
      switch(arah_input,
             "two.sided" = HTML("H<sub>0</sub>: p₁ = p₂ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: p₁ ≠ p₂"),
             "less"      = HTML("H<sub>0</sub>: p₁ ≥ p₂ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: p₁ < p₂"),
             "greater"   = HTML("H<sub>0</sub>: p₁ ≤ p₂ &nbsp;&nbsp;&nbsp;&nbsp; H<sub>1</sub>: p₁ > p₂"))
    }
    
    interpretasi_teks <- if (pval < alpha) {
      "Hasil pengujian menunjukkan bahwa terdapat perbedaan yang signifikan secara statistik antara proporsi yang diuji. Artinya, terdapat cukup bukti untuk menolak hipotesis nol (H₀) dan menyimpulkan bahwa proporsi berbeda."
    } else {
      "Hasil pengujian menunjukkan bahwa tidak terdapat perbedaan yang signifikan secara statistik. Artinya, tidak cukup bukti untuk menolak hipotesis nol (H₀), sehingga tidak dapat disimpulkan bahwa proporsi berbeda."
    }
    
    kesimpulan <- if (pval < alpha) {
      "<span style='color: #2e7d32; font-weight: 600;'>Terdapat perbedaan yang signifikan secara statistik.</span>"
    } else {
      "<span style='color: #c62828; font-weight: 600;'>Tidak terdapat perbedaan yang signifikan secara statistik.</span>"
    }
    
    metode_text <- switch(h$metode,
                          "prop" = "Uji proporsi normal (asymptotik)",
                          "binom" = "Uji binomial eksak",
                          "fisher" = "Uji Fisher exact"
    )
    
    catatan <- "<em>Catatan:</em>  <strong>Salah arah uji</strong> dapat menyebabkan kesalahan penarikan kesimpulan."
    
    label_stat <- if (h$metode == "fisher") {
      "Statistik Uji (Odds Ratio):"
    } else if (h$metode == "binom") {
      "Statistik Uji (Exact Binomial):"
    } else {
      "Statistik Uji (z):"
    }
    
    HTML(paste0(
      "<div style='padding: 10px; text-align: justify;'>",
      "<p><strong>Jenis Uji:</strong> ", jenis, "<br/>",
      "<strong>Arah Uji:</strong> ", arah, "<br/>",
      "<strong>", label_stat, "</strong> ", stat, "<br/>",
      "<strong>p-value:</strong> ", signif(pval, 4), "<br/>",
      "<strong>Confidence Interval:</strong> ", ci, "</p>",
      "<p><strong>Hipotesis:</strong><br/>", hipotesis, "</p>",
      "<p><strong>Kesimpulan:</strong><br/>", kesimpulan, "</p>",
      "<p><strong>Interpretasi:</strong> ", interpretasi_teks, "</p>",
      "<p><strong>Metode Uji:</strong> ", metode_text, "</p>",
      "<p style='font-size: 13px; color: #666;'>", catatan, "</p>",
      "</div>"
    ))
  })
  
  output$unduh_uji_proporsi_docx <- downloadHandler(
    filename = function() {
      paste0("hasil_uji_proporsi_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      h <- hasil_uji_proporsi()
      if (is.null(h)) return()
      
      hasil <- h$hasil
      jenis <- if (h$jenis == "satu") "Uji Proporsi Satu Kelompok" else "Uji Proporsi Dua Kelompok"
      
      arah_input <- if (h$jenis == "satu") input$arah_uji_proporsi else input$arah_uji_proporsi2
      arah <- switch(arah_input,
                     "two.sided" = "Dua Arah (p ≠ p₀)",
                     "less" = "Satu Arah Kiri (p < p₀ atau p₁ < p₂)",
                     "greater" = "Satu Arah Kanan (p > p₀ atau p₁ > p₂)")
      
      hipotesis <- if (h$jenis == "satu") {
        switch(arah_input,
               "two.sided" = "H₀: p = p₀, H₁: p ≠ p₀",
               "less" = "H₀: p ≥ p₀, H₁: p < p₀",
               "greater" = "H₀: p ≤ p₀, H₁: p > p₀")
      } else {
        switch(arah_input,
               "two.sided" = "H₀: p₁ = p₂, H₁: p₁ ≠ p₂",
               "less" = "H₀: p₁ ≥ p₂, H₁: p₁ < p₂",
               "greater" = "H₀: p₁ ≤ p₂, H₁: p₁ > p₂")
      }
      
      pval <- signif(hasil$p.value, 4)
      stat <- if (!is.null(hasil$statistic)) round(hasil$statistic, 3) else "Tidak tersedia"
      ci <- if (!is.null(hasil$conf.int)) paste0("[", signif(hasil$conf.int[1], 4), ", ", signif(hasil$conf.int[2], 4), "]") else "Tidak tersedia"
      
      interpretasi <- if (pval < input$alpha_proporsi) {
        "Terdapat perbedaan yang signifikan secara statistik."
      } else {
        "Tidak terdapat perbedaan yang signifikan secara statistik."
      }
      
      metode_text <- switch(h$metode,
                            "prop" = "Uji proporsi normal (asymptotik)",
                            "binom" = "Uji binomial eksak",
                            "fisher" = "Uji Fisher exact")
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Hasil Uji Proporsi", style = "heading 1")
      doc <- body_add_par(doc, paste("Jenis Uji:", jenis), style = "Normal")
      doc <- body_add_par(doc, paste("Arah Uji:", arah), style = "Normal")
      doc <- body_add_par(doc, paste("Hipotesis:", hipotesis), style = "Normal")
      doc <- body_add_par(doc, paste("Statistik Uji:", stat), style = "Normal")
      doc <- body_add_par(doc, paste("p-value:", pval), style = "Normal")
      doc <- body_add_par(doc, paste("Confidence Interval:", ci), style = "Normal")
      doc <- body_add_par(doc, paste("Metode Uji:", metode_text), style = "Normal")
      doc <- body_add_par(doc, paste("Kesimpulan:", interpretasi), style = "Normal")
      doc <- body_add_par(doc, "Catatan: Salah arah uji dapat menyebabkan kesalahan penarikan kesimpulan.", style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  
  # ======================= Uji Varians ===============================
  observe({
    df <- hasil_manajemen()
    var_analisis <- c("LOWEDU", "ILLITERATE", "POVERTY", "NOELECTRIC",
                      "ELDERLY", "FHEAD", "FAMILYSIZE", "RENTED")
    num_vars <- names(df)[names(df) %in% var_analisis]
    
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    # Jika kandidat kosong, maka grp_vars kosong juga
    if (length(kandidat) == 0) {
      grp_vars <- character(0)
    } else {
      grp_vars <- kandidat[
        sapply(df[, kandidat, drop = FALSE], function(x) length(unique(na.omit(x))) == 2)
      ]
    }
    
    updateSelectInput(session, "var_varian_satu", choices = num_vars)
    updateSelectInput(session, "var_varian_dua", choices = num_vars)
    updateSelectInput(session, "group_varian_dua", choices = grp_vars)
  })
  
  
  output$group_varian_dua_ui <- renderUI({
    df <- hasil_manajemen()
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(kandidat) == 0) {
      return(helpText("Tidak ditemukan variabel kategorik dengan tepat 2 kategori. Silakan buat melalui Manajemen Data."))
    }
    
    var_kategorik <- kandidat[
      sapply(df[, kandidat, drop = FALSE], function(x) length(unique(na.omit(x))) == 2)
    ]
    
    if (length(var_kategorik) == 0) {
      return(helpText("Tidak ditemukan variabel kategorik dengan tepat 2 kategori. Silakan buat melalui Manajemen Data."))
    }
    
    selectInput("group_varian_dua", "Variabel Kategorik (2 Kategori):", choices = var_kategorik)
  })
  
  hasil_uji_varians <- eventReactive(input$proses_uji_varians, {
    df <- hasil_manajemen()
    alpha <- input$alpha_varians
    
    if (input$jenis_uji_varians == "satu") {
      req(input$var_varian_satu, input$nilai_varian0)
      data <- na.omit(df[[input$var_varian_satu]])
      
      if (length(data) < 2) return(NULL)
      
      s2 <- var(data)
      n <- length(data)
      chi_sq <- (n - 1) * s2 / input$nilai_varian0
      pval <- switch(input$arah_uji_varian_satu,
                     "two.sided" = 2 * min(pchisq(chi_sq, df = n - 1), 1 - pchisq(chi_sq, df = n - 1)),
                     "less" = pchisq(chi_sq, df = n - 1),
                     "greater" = 1 - pchisq(chi_sq, df = n - 1))
      
      return(list(jenis = "satu", pval = pval, stat = chi_sq, df = n - 1, alpha = alpha, metode = "chisq"))
      
    } else {
      req(input$var_varian_dua, input$group_varian_dua)
      df <- df[!is.na(df[[input$var_varian_dua]]) & !is.na(df[[input$group_varian_dua]]), ]
      df[[input$group_varian_dua]] <- as.factor(df[[input$group_varian_dua]])
      
      if (length(levels(df[[input$group_varian_dua]])) != 2) return(NULL)
      
      formula <- as.formula(paste(input$var_varian_dua, "~", input$group_varian_dua))
      hasil <- var.test(formula, data = df, alternative = input$arah_uji_varian_dua)
      
      return(list(jenis = "dua", hasil = hasil, alpha = alpha, metode = "f"))
    }
  })
  
  output$hasil_uji_varians <- renderPrint({
    h <- hasil_uji_varians()
    req(h)
    
    if (h$jenis == "dua") {
      print(h$hasil)
    } else {
      cat("Statistik Uji (Chi-Square):", round(h$stat, 4), "\n",
          "Derajat bebas:", h$df, "\n",
          "p-value:", signif(h$pval, 4))
    }
  })
  
  output$interpretasi_uji_varians <- renderUI({
    h <- hasil_uji_varians()
    req(h)
    
    alpha <- h$alpha
    jenis <- if (h$jenis == "satu") "Uji Varians Satu Kelompok" else "Uji Varians Dua Kelompok"
    arah <- if (h$jenis == "satu") input$arah_uji_varian_satu else input$arah_uji_varian_dua
    label_arah <- switch(arah,
                         "two.sided" = "Dua Arah (≠)",
                         "less" = "Satu Arah Kiri (<)",
                         "greater" = "Satu Arah Kanan (>)")
    
    if (h$jenis == "satu") {
      pval <- h$pval
      stat <- round(h$stat, 3)
      df <- h$df
      hipotesis <- switch(arah,
                          "two.sided" = HTML("H<sub>0</sub>: \u03c3² = \u03c3₀² &nbsp;&nbsp;&nbsp; H<sub>1</sub>: \u03c3² ≠ \u03c3₀²"),
                          "less" = HTML("H<sub>0</sub>: \u03c3² ≥ \u03c3₀² &nbsp;&nbsp;&nbsp; H<sub>1</sub>: \u03c3² < \u03c3₀²"),
                          "greater" = HTML("H<sub>0</sub>: \u03c3² ≤ \u03c3₀² &nbsp;&nbsp;&nbsp; H<sub>1</sub>: \u03c3² > \u03c3₀²"))
    } else {
      pval <- h$hasil$p.value
      stat <- round(h$hasil$statistic, 3)
      df <- paste(h$hasil$parameter[1], ",", h$hasil$parameter[2])
      hipotesis <- switch(arah,
                          "two.sided" = HTML("H<sub>0</sub>: \u03c3₁² = \u03c3₂² &nbsp;&nbsp;&nbsp; H<sub>1</sub>: \u03c3₁² ≠ \u03c3₂²"),
                          "less" = HTML("H<sub>0</sub>: \u03c3₁² ≥ \u03c3₂² &nbsp;&nbsp;&nbsp; H<sub>1</sub>: \u03c3₁² < \u03c3₂²"),
                          "greater" = HTML("H<sub>0</sub>: \u03c3₁² ≤ \u03c3₂² &nbsp;&nbsp;&nbsp; H<sub>1</sub>: \u03c3₁² > \u03c3₂²"))
    }
    
    interpretasi <- if (pval < alpha) {
      "Terdapat perbedaan varians yang signifikan secara statistik. Artinya, terdapat cukup bukti untuk menolak H₀."
    } else {
      "Tidak terdapat perbedaan varians yang signifikan secara statistik. Artinya, tidak cukup bukti untuk menolak H₀."
    }
    
    kesimpulan <- if (pval < alpha) {
      "<span style='color: #2e7d32; font-weight: 600;'>Terdapat perbedaan yang signifikan secara statistik.</span>"
    } else {
      "<span style='color: #c62828; font-weight: 600;'>Tidak terdapat perbedaan yang signifikan secara statistik.</span>"
    }
    
    label_stat <- if (h$metode == "chisq") "Statistik Uji (Chi-Square):" else "Statistik Uji (F):"
    
    HTML(paste0(
      "<div style='padding: 10px; text-align: justify;'>",
      "<p><strong>Jenis Uji:</strong> ", jenis, "<br/>",
      "<strong>Arah Uji:</strong> ", label_arah, "<br/>",
      "<strong>", label_stat, "</strong> ", stat, "<br/>",
      "<strong>Derajat Bebas:</strong> ", df, "<br/>",
      "<strong>p-value:</strong> ", signif(pval, 4), "</p>",
      "<p><strong>Hipotesis:</strong><br/>", hipotesis, "</p>",
      "<p><strong>Kesimpulan:</strong><br/>", kesimpulan, "</p>",
      "<p><strong>Interpretasi:</strong> ", interpretasi, "</p>",
      "<p style='font-size: 13px; color: #666;'><em>Catatan:</em> Uji varians mengasumsikan data berdistribusi normal.</p>",
      "</div>"
    ))
  })
  
  output$unduh_uji_varians_docx <- downloadHandler(
    filename = function() {
      paste0("hasil_uji_varians_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      h <- hasil_uji_varians()
      if (is.null(h)) return()
      
      doc <- read_docx()
      
      jenis <- if (h$jenis == "satu") "Uji Varians Satu Kelompok" else "Uji Varians Dua Kelompok"
      arah <- if (h$jenis == "satu") input$arah_uji_varian_satu else input$arah_uji_varian_dua
      arah_text <- switch(arah,
                          "two.sided" = "Dua Arah (≠)",
                          "less" = "Satu Arah Kiri (<)",
                          "greater" = "Satu Arah Kanan (>)")
      
      hipotesis <- if (h$jenis == "satu") {
        switch(arah,
               "two.sided" = "H0: σ² = σ₀², H1: σ² ≠ σ₀²",
               "less"      = "H0: σ² ≥ σ₀², H1: σ² < σ₀²",
               "greater"   = "H0: σ² ≤ σ₀², H1: σ² > σ₀²")
      } else {
        switch(arah,
               "two.sided" = "H0: σ₁² = σ₂², H1: σ₁² ≠ σ₂²",
               "less"      = "H0: σ₁² ≥ σ₂², H1: σ₁² < σ₂²",
               "greater"   = "H0: σ₁² ≤ σ₂², H1: σ₁² > σ₂²")
      }
      
      if (h$jenis == "satu") {
        stat <- round(h$stat, 4)
        df <- h$df
        pval <- signif(h$pval, 4)
        metode <- "Chi-Square"
      } else {
        stat <- round(h$hasil$statistic, 4)
        df <- paste0("(", h$hasil$parameter[1], ", ", h$hasil$parameter[2], ")")
        pval <- signif(h$hasil$p.value, 4)
        metode <- "F"
      }
      
      kesimpulan <- if (pval < h$alpha) {
        "Terdapat perbedaan varians yang signifikan secara statistik."
      } else {
        "Tidak terdapat perbedaan varians yang signifikan secara statistik."
      }
      
      doc <- body_add_par(doc, "Hasil Uji Varians", style = "heading 1")
      doc <- body_add_par(doc, paste("Jenis Uji:", jenis), style = "Normal")
      doc <- body_add_par(doc, paste("Arah Uji:", arah_text), style = "Normal")
      doc <- body_add_par(doc, paste("Hipotesis:", hipotesis), style = "Normal")
      doc <- body_add_par(doc, paste("Metode Uji:", metode), style = "Normal")
      doc <- body_add_par(doc, paste("Statistik Uji:", stat), style = "Normal")
      doc <- body_add_par(doc, paste("Derajat Bebas:", df), style = "Normal")
      doc <- body_add_par(doc, paste("p-value:", pval), style = "Normal")
      doc <- body_add_par(doc, paste("Kesimpulan:", kesimpulan), style = "Normal")
      doc <- body_add_par(doc, "Catatan: Uji varians mengasumsikan data berasal dari populasi berdistribusi normal.", style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  
  # ======================= Uji ANOVA ===============================
  observe({
    df <- hasil_manajemen()
    var_analisis <- c("ILLITERATE", "LOWEDU", "POVERTY", "NOELECTRIC",
                      "ELDERLY", "FHEAD", "FAMILYSIZE", "RENTED")
    
    num_vars <- names(df)[names(df) %in% var_analisis]
    
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    kat_vars <- if (length(kandidat) > 0) {
      kandidat[sapply(df[kandidat], function(x) length(unique(na.omit(x))) >= 2)]
    } else {
      character(0)
    }
    
    updateSelectInput(session, "anova_var_numerik1", choices = num_vars)
    updateSelectInput(session, "anova_var_numerik2", choices = num_vars)
    updateSelectInput(session, "anova_group1", choices = kat_vars)
    updateSelectInput(session, "anova_group2a", choices = kat_vars)
    updateSelectInput(session, "anova_group2b", choices = kat_vars)
  })
  
  # Faktor untuk ANOVA Satu Arah
  output$anova_group1_ui <- renderUI({
    df <- hasil_manajemen()
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(kandidat) == 0) {
      return(helpText("Belum ada variabel kategorik. Silakan lakukan kategorisasi terlebih dahulu melalui menu Manajemen Data."))
    }
    
    kat_vars <- kandidat[sapply(df[kandidat], function(x) length(unique(na.omit(x))) >= 2)]
    
    if (length(kat_vars) == 0) {
      return(helpText("Variabel kategorik belum memenuhi syarat (minimal 2 kategori berbeda). Silakan sesuaikan kategorisasi."))
    }
    
    selectInput("anova_group1", "Faktor Kategorik:", choices = kat_vars)
  })
  
  # Faktor A untuk ANOVA Dua Arah
  output$anova_group2a_ui <- renderUI({
    df <- hasil_manajemen()
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(kandidat) == 0) {
      return(helpText("Belum ada variabel kategorik. Silakan lakukan kategorisasi terlebih dahulu melalui menu Manajemen Data."))
    }
    
    kat_vars <- kandidat[sapply(df[kandidat], function(x) length(unique(na.omit(x))) >= 2)]
    
    if (length(kat_vars) == 0) {
      return(helpText("Variabel kategorik belum memenuhi syarat (minimal 2 kategori berbeda). Silakan sesuaikan kategorisasi."))
    }
    
    selectInput("anova_group2a", "Faktor Kategorik A:", choices = kat_vars)
  })
  
  # Faktor B untuk ANOVA Dua Arah
  output$anova_group2b_ui <- renderUI({
    df <- hasil_manajemen()
    kandidat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (length(kandidat) == 0) {
      return(helpText("Belum ada variabel kategorik. Silakan lakukan kategorisasi terlebih dahulu melalui menu Manajemen Data."))
    }
    
    kat_vars <- kandidat[sapply(df[kandidat], function(x) length(unique(na.omit(x))) >= 2)]
    
    if (length(kat_vars) == 0) {
      return(helpText("Variabel kategorik belum memenuhi syarat (minimal 2 kategori berbeda). Silakan sesuaikan kategorisasi."))
    }
    
    selectInput("anova_group2b", "Faktor Kategorik B:", choices = kat_vars)
  })
  
  
  hasil_uji_anova <- eventReactive(input$proses_uji_anova, {
    df <- hasil_manajemen()
    alpha <- input$alpha_anova
    
    var_analisis <- c("ILLITERATE", "LOWEDU", "POVERTY", "NOELECTRIC")
    kandidat_kat <- grep("_(SL|Q|J)$", names(df), value = TRUE)
    
    if (input$jenis_uji_anova == "satu") {
      req(input$anova_var_numerik1, input$anova_group1)
      
      # Validasi numerik
      if (!(input$anova_var_numerik1 %in% var_analisis)) {
        showModal(modalDialog(title = "Peringatan",
                              "Variabel numerik tidak valid. Hanya variabel analisis yang diizinkan.",
                              easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      # Validasi kategorik
      if (!(input$anova_group1 %in% kandidat_kat)) {
        showModal(modalDialog(title = "Peringatan",
                              "Variabel kategorik tidak valid. Silakan lakukan kategorisasi terlebih dahulu.",
                              easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      formula <- as.formula(paste(input$anova_var_numerik1, "~", input$anova_group1))
      model <- aov(formula, data = df)
      return(list(jenis = "satu", model = model, alpha = alpha))
      
    } else {
      req(input$anova_var_numerik2, input$anova_group2a, input$anova_group2b)
      
      # Validasi numerik
      if (!(input$anova_var_numerik2 %in% var_analisis)) {
        showModal(modalDialog(title = "Peringatan",
                              "Variabel numerik tidak valid. Hanya variabel analisis yang diizinkan.",
                              easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      # Validasi kategorik
      if (!(input$anova_group2a %in% kandidat_kat) || !(input$anova_group2b %in% kandidat_kat)) {
        showModal(modalDialog(title = "Peringatan",
                              "Variabel kategorik tidak valid. Silakan lakukan kategorisasi terlebih dahulu.",
                              easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      # Validasi agar dua group tidak sama
      if (input$anova_group2a == input$anova_group2b) {
        showModal(modalDialog(title = "Peringatan",
                              "Faktor kategorik A dan B tidak boleh sama.",
                              easyClose = TRUE, footer = modalButton("Tutup")))
        return(NULL)
      }
      
      interaksi <- if (isTRUE(input$anova_interaksi)) "*" else "+"
      formula <- as.formula(paste(input$anova_var_numerik2, "~",
                                  paste(input$anova_group2a, interaksi, input$anova_group2b)))
      model <- aov(formula, data = df)
      return(list(jenis = "dua", model = model, alpha = alpha))
    }
  })
  
  
  output$hasil_uji_anova <- renderPrint({
    h <- hasil_uji_anova()
    req(h)
    summary(h$model)
  })
  
  output$interpretasi_uji_anova <- renderUI({
    h <- hasil_uji_anova()
    req(h)
    hasil <- summary(h$model)
    alpha <- h$alpha
    
    pval <- summary(h$model)[[1]][["Pr(>F)"]][1]
    
    jenis <- if (h$jenis == "satu") "ANOVA Satu Arah" else
      if (isTRUE(input$anova_interaksi)) "ANOVA Dua Arah dengan Interaksi" else "ANOVA Dua Arah"
    
    interpretasi <- if (pval < alpha) {
      "Terdapat perbedaan rata-rata yang signifikan antar kelompok."
    } else {
      "Tidak terdapat perbedaan rata-rata yang signifikan antar kelompok."
    }
    
    kesimpulan <- if (pval < alpha) {
      "<span style='color: #2e7d32; font-weight: 600;'>Signifikan: Tolak H<sub>0</sub>.</span>"
    } else {
      "<span style='color: #c62828; font-weight: 600;'>Tidak signifikan: Gagal tolak H<sub>0</sub>.</span>"
    }
    
    HTML(paste0(
      "<div style='padding: 10px; text-align: justify;'>",
      "<p><strong>Jenis Uji:</strong> ", jenis, "<br/>",
      "<strong>p-value:</strong> ", signif(pval, 4), "</p>",
      "<p><strong>Kesimpulan:</strong><br/>", kesimpulan, "</p>",
      "<p><strong>Interpretasi:</strong> ", interpretasi, "</p>",
      "<p style='font-size: 13px; color: #666;'><em>Catatan:</em> ANOVA mengasumsikan normalitas dan homogenitas varians.</p>",
      "</div>"
    ))
  })
  
  output$unduh_uji_anova_docx <- downloadHandler(
    filename = function() {
      paste0("hasil_uji_anova_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      h <- hasil_uji_anova()
      if (is.null(h)) return()
      
      hasil <- summary(h$model)
      jenis <- if (h$jenis == "satu") "ANOVA Satu Arah" else {
        if (isTRUE(input$anova_interaksi)) "ANOVA Dua Arah dengan Interaksi" else "ANOVA Dua Arah"
      }
      
      pval <- hasil[[1]][["Pr(>F)"]][1]
      alpha <- h$alpha
      signifikan <- if (pval < alpha) TRUE else FALSE
      kesimpulan <- if (signifikan) {
        "Terdapat perbedaan rata-rata yang signifikan antar kelompok."
      } else {
        "Tidak terdapat perbedaan rata-rata yang signifikan antar kelompok."
      }
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Hasil Uji ANOVA", style = "heading 1")
      doc <- body_add_par(doc, paste("Jenis Uji:", jenis), style = "Normal")
      doc <- body_add_par(doc, paste("Tingkat Signifikansi (α):", alpha), style = "Normal")
      doc <- body_add_par(doc, paste("p-value:", signif(pval, 4)), style = "Normal")
      doc <- body_add_par(doc, paste("Kesimpulan:", kesimpulan), style = "Normal")
      doc <- body_add_par(doc, "Catatan: ANOVA mengasumsikan normalitas dan homogenitas varians.", style = "Normal")
      
      doc <- body_add_par(doc, "Rangkuman Tabel ANOVA:", style = "heading 2")
      # Tambahkan tabel hasil ANOVA
      hasil_df <- as.data.frame(hasil[[1]])
      hasil_df <- tibble::rownames_to_column(hasil_df, var = "Sumber Variasi")
      doc <- body_add_table(doc, hasil_df, style = "table_template")
      
      print(doc, target = file)
    }
  )


  # ======================== Regresi Linier ========================
  
  model_regresi <- eventReactive(input$proses_regresi, {
    req(input$x_rlb)
    data <- hasil_manajemen()
    
    # Buat formula regresi
    form <- as.formula(paste("LOWEDU ~", paste(input$x_rlb, collapse = "+")))
    model <- lm(form, data = data)
    return(model)
  })
  
  output$output_regresi <- renderPrint({
    model <- model_regresi()
    summary(model)
  })
  
  output$interpretasi_regresi <- renderUI({
    model <- model_regresi()
    r2 <- round(summary(model)$r.squared, 4)
    pval <- summary(model)$coefficients[,"Pr(>|t|)"]
    
    signif_count <- sum(pval < 0.05)
    HTML(paste0(
      "<p><strong>R-squared (R²):</strong> ", r2, "</p>",
      "<p><strong>Jumlah Prediktor Signifikan (p < 0.05):</strong> ", signif_count, "</p>",
      "<p><strong>Interpretasi:</strong><br>",
      "Nilai R² sebesar ", r2, " menunjukkan bahwa model mampu menjelaskan sekitar ", round(r2 * 100, 1), 
      "% variasi yang terjadi pada variabel respon <b>LOWEDU</b> (persentase penduduk berpendidikan rendah). ",
      if (r2 >= 0.5) {
        "Artinya, model memiliki daya jelas yang cukup baik."
      } else {
        "Namun, nilai ini masih tergolong rendah sehingga terdapat variabel lain di luar model yang turut memengaruhi."
      },
      "</p>"
    ))
  })
  
  
  output$uji_normalitas <- renderPrint({
    model <- model_regresi()
    residual <- residuals(model)
    
    if (length(residual) > 5000) {
      cat("Data terlalu besar untuk uji Shapiro-Wilk (maksimum 5000 sampel).")
      return()
    }
    
    shapiro.test(residual)
  })
  
  output$interpretasi_normalitas <- renderUI({
    model <- model_regresi()
    res <- residuals(model)
    if (length(res) > 5000) return(NULL)
    
    uji <- shapiro.test(res)
    pval <- uji$p.value
    HTML(paste0(
      "<p><strong>Hipotesis:</strong><br>",
      "H₀: Residual model berdistribusi normal<br>",
      "H₁: Residual model tidak berdistribusi normal</p>",
      
      "<p><strong>p-value:</strong> ", signif(pval, 4), "</p>",
      
      "<p><strong>Interpretasi:</strong><br>",
      if (pval > 0.05) {
        "<span style='color:#2e7d32;'>Karena p-value > 0.05, maka tidak cukup bukti untuk menolak H₀. 
      Dengan demikian, residual model <b>berdistribusi normal</b> dan asumsi normalitas <b>terpenuhi</b>.</span>"
      } else {
        "<span style='color:#c62828;'>Karena p-value ≤ 0.05, maka terdapat cukup bukti untuk menolak H₀. 
      Dengan demikian, residual <b>tidak berdistribusi normal</b> dan asumsi normalitas <b>tidak terpenuhi</b>.</span>"
      }, "</p>"
    ))
  })
  
  
  output$uji_vif <- renderPrint({
    model <- model_regresi()
    car::vif(model)
  })
  
  output$interpretasi_vif <- renderUI({
    model <- model_regresi()
    v <- car::vif(model)
    max_vif <- max(v)
    
    HTML(paste0(
      "<p><strong>Interpretasi:</strong><br>",
      "Multikolinearitas terjadi jika terdapat hubungan yang sangat kuat antar variabel X (prediktor), yang menyebabkan ketidakstabilan estimasi koefisien regresi.",
      "<br>Nilai <b>VIF (Variance Inflation Factor)</b> digunakan untuk mendeteksi hal ini. Jika VIF > 10, maka ada indikasi kuat adanya multikolinearitas.</p>",
      
      "<p><strong>VIF Tertinggi:</strong> ", round(max_vif, 2), "</p>",
      
      if (max_vif > 10) {
        "<p><span style='color:#c62828;'>Terdapat multikolinearitas kuat antar variabel prediktor. Sebaiknya lakukan pengecekan ulang dan pertimbangkan untuk menghapus salah satu variabel yang berkorelasi tinggi.</span></p>"
      } else {
        "<p><span style='color:#2e7d32;'>Tidak ditemukan multikolinearitas yang signifikan. Semua variabel prediktor dapat digunakan dalam model regresi.</span></p>"
      }
    ))
  })
  
  
  output$uji_breusch <- renderPrint({
    model <- model_regresi()
    lmtest::bptest(model)
  })
  
  output$interpretasi_breusch <- renderUI({
    bp <- lmtest::bptest(model_regresi())
    pval <- bp$p.value
    
    HTML(paste0(
      "<p><strong>Hipotesis:</strong><br>",
      "H₀: Residual memiliki varians yang homogen (homoskedastisitas)<br>",
      "H₁: Residual memiliki varians yang tidak homogen (heteroskedastisitas)</p>",
      
      "<p><strong>p-value:</strong> ", signif(pval, 4), "</p>",
      
      "<p><strong>Interpretasi:</strong><br>",
      if (pval > 0.05) {
        "<span style='color:#2e7d32;'>Karena p-value > 0.05, maka tidak cukup bukti untuk menolak H₀. 
      Residual model <b>memiliki varians homogen</b>. Asumsi homoskedastisitas terpenuhi.</span>"
      } else {
        "<span style='color:#c62828;'>Karena p-value ≤ 0.05, maka terdapat heteroskedastisitas. 
      Asumsi varians homogen <b>tidak terpenuhi</b>, sehingga estimasi koefisien bisa menjadi tidak efisien.</span>"
      }, "</p>"
    ))
  })
  
  
  output$uji_moran <- renderPrint({
    model <- model_regresi()
    res <- residuals(model)
    
    # Baca matriks jarak
    dist_mat <- read.csv("Data/distance.csv", row.names = 1)
    mat <- as.matrix(dist_mat)
    
    # Buat matriks bobot: jika jarak < threshold, dianggap neighbor
    threshold <- quantile(mat, 0.1, na.rm = TRUE)  # bisa kamu ubah jadi nilai tetap, misal 1.5
    weight_mat <- (mat <= threshold) * 1
    diag(weight_mat) <- 0  # nolkan diagonal
    
    # Ubah ke listw
    nb_list <- spdep::mat2listw(weight_mat, style = "W")
    
    # Uji Moran
    spdep::moran.test(res, nb_list)
  })
  
  
  output$interpretasi_moran <- renderUI({
    model <- model_regresi()
    res <- residuals(model)
    
    dist_mat <- read.csv("Data/distance.csv", row.names = 1)
    mat <- as.matrix(dist_mat)
    
    threshold <- quantile(mat, 0.1, na.rm = TRUE)
    weight_mat <- (mat <= threshold) * 1
    diag(weight_mat) <- 0
    
    nb_list <- spdep::mat2listw(weight_mat, style = "W")
    mtest <- spdep::moran.test(res, nb_list)
    pval <- mtest$p.value
    
    HTML(paste0(
      "<p><strong>Hipotesis:</strong><br>",
      "H₀: Tidak ada autokorelasi spasial pada residual<br>",
      "H₁: Terdapat autokorelasi spasial pada residual</p>",
      
      "<p><strong>p-value:</strong> ", signif(pval, 4), "</p>",
      
      "<p><strong>Interpretasi:</strong><br>",
      if (pval < 0.05) {
        "<span style='color:#c62828;'>Terdapat cukup bukti untuk menolak H₀. 
      <b>Ada autokorelasi spasial</b> pada residual, artinya nilai residual di suatu wilayah cenderung dipengaruhi oleh wilayah sekitar. 
      Ini menunjukkan ketergantungan spasial yang bisa mengganggu validitas model.</span>"
      } else {
        "<span style='color:#2e7d32;'>Tidak ada cukup bukti untuk menolak H₀. Residual model <b>tidak memiliki pola spasial</b>, sehingga asumsi ini <b>terpenuhi</b>.</span>"
      }, "</p>"
    ))
  })
  
  output$unduh_regresi_docx <- downloadHandler(
    filename = function() {
      paste0("laporan_regresi_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      library(car)
      library(lmtest)
      library(spdep)
      
      model <- model_regresi()
      res <- residuals(model)
      
      # ===== Ringkasan Model =====
      summ <- capture.output(summary(model))
      r2 <- round(summary(model)$r.squared, 4)
      pval <- summary(model)$coefficients[,"Pr(>|t|)"]
      signif_count <- sum(pval < 0.05)
      
      # ===== Uji Normalitas =====
      if (length(res) <= 5000) {
        norm <- shapiro.test(res)
        norm_text <- paste0("p-value = ", signif(norm$p.value, 4),
                            if (norm$p.value > 0.05) {
                              " → Residual berdistribusi normal (asumsi terpenuhi)."
                            } else {
                              " → Residual tidak normal (asumsi tidak terpenuhi)."
                            })
      } else {
        norm_text <- "Data terlalu besar untuk uji Shapiro-Wilk."
      }
      
      # ===== VIF =====
      vif_values <- vif(model)
      max_vif <- max(vif_values)
      vif_text <- paste0("VIF tertinggi: ", round(max_vif, 2),
                         if (max_vif > 10) {
                           " → Terdapat multikolinearitas kuat."
                         } else {
                           " → Tidak ada indikasi multikolinearitas berarti."
                         })
      
      # ===== Breusch-Pagan =====
      bp <- bptest(model)
      bp_text <- paste0("p-value = ", signif(bp$p.value, 4),
                        if (bp$p.value > 0.05) {
                          " → Homoskedastisitas terpenuhi."
                        } else {
                          " → Terdapat heteroskedastisitas."
                        })
      
      # ===== Moran's I =====
      dist_mat <- read.csv("Data/distance.csv", row.names = 1)
      mat <- as.matrix(dist_mat)
      threshold <- quantile(mat, 0.1, na.rm = TRUE)
      weight_mat <- (mat <= threshold) * 1
      diag(weight_mat) <- 0
      nb_list <- mat2listw(weight_mat, style = "W")
      mtest <- moran.test(res, nb_list)
      moran_text <- paste0("p-value = ", signif(mtest$p.value, 4),
                           if (mtest$p.value < 0.05) {
                             " → Terdapat autokorelasi spasial."
                           } else {
                             " → Tidak terdapat autokorelasi spasial."
                           })
      
      # ===== Bangun Dokumen =====
      doc <- read_docx()
      doc <- body_add_par(doc, "Laporan Analisis Regresi Linier", style = "heading 1")
      
      doc <- body_add_par(doc, "Ringkasan Model", style = "heading 2")
      for (line in summ) {
        doc <- body_add_par(doc, line, style = "Normal")
      }
      doc <- body_add_par(doc, paste0("R-squared: ", r2), style = "Normal")
      doc <- body_add_par(doc, paste0("Jumlah prediktor signifikan (p < 0.05): ", signif_count), style = "Normal")
      
      doc <- body_add_par(doc, "Uji Normalitas Residual (Shapiro-Wilk)", style = "heading 2")
      doc <- body_add_par(doc, norm_text, style = "Normal")
      
      doc <- body_add_par(doc, "Uji Multikolinearitas (VIF)", style = "heading 2")
      for (v in names(vif_values)) {
        doc <- body_add_par(doc, paste0(v, ": ", round(vif_values[[v]], 2)), style = "Normal")
      }
      doc <- body_add_par(doc, vif_text, style = "Normal")
      
      doc <- body_add_par(doc, "Uji Heteroskedastisitas (Breusch-Pagan)", style = "heading 2")
      doc <- body_add_par(doc, bp_text, style = "Normal")
      
      doc <- body_add_par(doc, "Uji Autokorelasi Spasial (Moran's I)", style = "heading 2")
      doc <- body_add_par(doc, moran_text, style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  # ======================== Estimasi Regresi Linier ========================
  # Model final tetap
  model_final2 <- lm(LOWEDU ~ ELDERLY + FHEAD + POVERTY + RENTED, data = data)
  
  # Reaktif menghitung estimasi
  estimasi_nilai <- eventReactive(input$hitung_estimasi, {
    req(input$input_elderly, input$input_fhead, input$input_poverty, input$input_rented)
    
    new_data <- data.frame(
      ELDERLY = input$input_elderly,
      FHEAD = input$input_fhead,
      POVERTY = input$input_poverty,
      RENTED = input$input_rented
    )
    
    pred <- predict(model_final2, newdata = new_data)
    return(round(pred, 2))
  })
  
  # Tampilkan hasil
  output$hasil_estimasi <- renderUI({
    req(estimasi_nilai())
    
    HTML(paste0(
      "<h4>Hasil Estimasi:</h4>",
      "<p>Dengan input variabel sebagai berikut:</p>",
      "<ul>",
      "<li><strong>ELDERLY:</strong> ", input$input_elderly, "</li>",
      "<li><strong>FHEAD:</strong> ", input$input_fhead, "</li>",
      "<li><strong>POVERTY:</strong> ", input$input_poverty, "</li>",
      "<li><strong>RENTED:</strong> ", input$input_rented, "</li>",
      "</ul>",
      "<p><strong>Perkiraan nilai LOWEDU (pendidikan rendah):</strong> ",
      "<span style='color:green; font-size: 18px;'>", estimasi_nilai(), "</span></p>"
    ))
  })
  
  # Interpretasi hasil estimasi
  output$hasil_estimasi <- renderUI({
    req(estimasi_nilai())
    
    HTML(paste0(
      "<h4><strong>Estimasi Persentase Pendidikan Rendah (LOWEDU)</strong></h4>",
      "<p><strong>Hasil Estimasi:</strong> ",
      "<span style='color:#2e7d32; font-size: 18px; font-weight:bold;'>", estimasi_nilai(), "%</span></p>",
      
      "<p><strong>Interpretasi:</strong> Dengan kondisi karakteristik wilayah seperti yang diinput, ",
      "diperkirakan sekitar <strong>", estimasi_nilai(), "%</strong> dari penduduk usia 15 tahun ke atas memiliki tingkat pendidikan rendah. ",
      "</p>"
    ))
  })

  output$interpretasi_estimasi <- renderUI({
    coef <- round(coef(model_final2), 4)
    
    persamaan <- paste0("LOWEDU = ", coef[1], " + ",
                        coef[2], "*ELDERLY + ",
                        coef[3], "*FHEAD + ",
                        coef[4], "*POVERTY + ",
                        coef[5], "*RENTED")
    
    HTML(paste0(
      "<p><strong>Persamaan Model:</strong></p>",
      "<pre>", persamaan, "</pre>",
      "<div style='text-align: justify;'>",
      "<p><strong>Catatan:</strong></p>",
      "<ul>",
      "<li>Model regresi ini menggunakan 4 variabel prediktor utama yang signifikan: <em>ELDERLY</em>, <em>FHEAD</em>, <em>POVERTY</em>, dan <em>RENTED</em>.</li>",
      "<li>Setiap koefisien menunjukkan estimasi perubahan rata-rata LOWEDU (persentase pendidikan rendah) terhadap perubahan satu satuan pada variabel terkait, dengan asumsi variabel lainnya tetap konstan.</li>",
      "<li>Jika terdapat koefisien bernilai negatif, hal tersebut bukan berarti hubungan sebab-akibat yang langsung, namun merupakan hasil dari pola data pada dataset yang dianalisis.</li>",
      "<li>Model ini merupakan model terbaik yang dipilih berdasarkan uji statistik menyeluruh, dan dapat digunakan untuk memprediksi proporsi penduduk berpendidikan rendah di suatu wilayah.</li>",
      "</ul>"
      
    ))
  })
  
  
  output$unduh_estimasi <- downloadHandler(
    filename = function() {
      paste0("estimasi_LOWEEDU_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      doc <- read_docx()
      doc <- body_add_par(doc, "Hasil Estimasi Persentase Pendidikan Rendah (LOWEDU)", style = "heading 1")
      doc <- body_add_par(doc, "Model regresi linier digunakan untuk mengestimasi nilai LOWEDU berdasarkan 4 variabel prediktor:", style = "Normal")
      doc <- body_add_par(doc, "- ELDERLY: Persentase penduduk usia 65 tahun ke atas", style = "Normal")
      doc <- body_add_par(doc, "- FHEAD: Persentase rumah tangga dengan kepala perempuan", style = "Normal")
      doc <- body_add_par(doc, "- POVERTY: Persentase penduduk miskin", style = "Normal")
      doc <- body_add_par(doc, "- RENTED: Persentase rumah tangga yang menyewa tempat tinggal", style = "Normal")
      
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "Input Pengguna:", style = "heading 2")
      doc <- body_add_par(doc, paste0("ELDERLY: ", input$input_elderly, "%"), style = "Normal")
      doc <- body_add_par(doc, paste0("FHEAD: ", input$input_fhead, "%"), style = "Normal")
      doc <- body_add_par(doc, paste0("POVERTY: ", input$input_poverty, "%"), style = "Normal")
      doc <- body_add_par(doc, paste0("RENTED: ", input$input_rented, "%"), style = "Normal")
      
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "Hasil Estimasi:", style = "heading 2")
      doc <- body_add_par(doc, paste0("Perkiraan persentase penduduk berpendidikan rendah (LOWEDU): ", estimasi_nilai(), "%"), style = "Normal")
      
      doc <- body_add_par(doc, "", style = "Normal")
      doc <- body_add_par(doc, "Interpretasi:", style = "heading 2")
      doc <- body_add_par(doc, paste0(
        "Berdasarkan input yang diberikan, diperkirakan sebesar ", estimasi_nilai(),
        "% penduduk usia 15 tahun ke atas di wilayah tersebut memiliki tingkat pendidikan rendah. ",
        "Estimasi ini dihitung menggunakan model regresi terbaik yang mempertimbangkan variabel demografis dan sosial yang signifikan."
      ), style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  output$catatan_ringkasan_model <- renderPrint({
    summary(model_final2)
  })
  
  output$catatan_interpretasi_model <- renderUI({
    r2 <- summary(model_final2)$r.squared
    r2_percent <- round(r2 * 100, 2)  # Konversi ke persen
    pval <- summary(model_final2)$coefficients[, "Pr(>|t|)"]
    signif_count <- sum(pval < 0.05)
    
    HTML(paste0(
      "<p><strong>R-squared (R²):</strong> ", r2, " (", r2_percent, "%)</p>",
      "<p><strong>Jumlah variabel prediktor signifikan (p < 0.05):</strong> ", signif_count, " dari ", length(pval), "</p>",
      "<p><strong>Interpretasi:</strong> ",
      "Model regresi ini mampu menjelaskan sekitar <strong>", r2_percent, "%</strong> variasi yang terjadi pada variabel respon (<em>LOWEDU</em>). ",
      if (r2 >= 0.5) {
        "Ini menunjukkan bahwa model memiliki daya prediksi yang cukup baik dalam memodelkan faktor-faktor yang memengaruhi tingkat pendidikan rendah."
      } else {
        "Nilai R² yang relatif rendah menunjukkan bahwa sebagian besar variasi <em>LOWEDU</em> belum dapat dijelaskan oleh model ini."
      },
      "</p>"
    ))
  })
  
  
  output$catatan_uji_normalitas <- renderPrint({
    shapiro.test(residuals(model_final2))
  })
  
  output$catatan_interpretasi_normalitas <- renderUI({
    pval <- shapiro.test(residuals(model_final2))$p.value
    HTML(paste0(
      "<p><strong>p-value:</strong> ", signif(pval, 4), "</p>",
      "<p><strong>Interpretasi:</strong> ",
      if (pval > 0.05) {
        "Residual terdistribusi normal. Asumsi normalitas terpenuhi."
      } else {
        "Residual tidak normal. Asumsi normalitas tidak terpenuhi."
      }, "</p>"
    ))
  })
  
  output$catatan_uji_vif <- renderPrint({
    car::vif(model_final2)
  })
  
  output$catatan_interpretasi_vif <- renderUI({
    v <- car::vif(model_final2)
    max_vif <- max(v)
    HTML(paste0(
      "<p><strong>VIF Tertinggi:</strong> ", round(max_vif, 2), "</p>",
      "<p><strong>Interpretasi:</strong> ",
      if (max_vif > 10) {
        "Terdapat multikolinearitas tinggi antar prediktor. Sebaiknya periksa ulang pemilihan variabel."
      } else if (max_vif > 5) {
        "Indikasi multikolinearitas sedang ditemukan."
      } else {
        "Tidak ada indikasi multikolinearitas serius."
      }, "</p>"
    ))
  })
  
  output$catatan_uji_breusch <- renderPrint({
    lmtest::bptest(model_final2)
  })
  
  output$catatan_interpretasi_breusch <- renderUI({
    pval <- lmtest::bptest(model_final2)$p.value
    HTML(paste0(
      "<p><strong>p-value:</strong> ", signif(pval, 4), "</p>",
      "<p><strong>Interpretasi:</strong> ",
      if (pval > 0.05) {
        "Tidak ditemukan heteroskedastisitas. Asumsi homoskedastisitas terpenuhi."
      } else {
        "Terdapat heteroskedastisitas. Asumsi tidak terpenuhi."
      }, "</p>"
    ))
  })
  
  output$catatan_uji_moran <- renderPrint({
    res <- residuals(model_final2)
    mat <- as.matrix(read.csv("Data/distance.csv", row.names = 1))
    threshold <- quantile(mat, 0.1, na.rm = TRUE)
    weight_mat <- (mat <= threshold) * 1
    diag(weight_mat) <- 0
    nb_list <- spdep::mat2listw(weight_mat, style = "W")
    spdep::moran.test(res, nb_list)
  })
  
  output$catatan_interpretasi_moran <- renderUI({
    res <- residuals(model_final2)
    mat <- as.matrix(read.csv("Data/distance.csv", row.names = 1))
    threshold <- quantile(mat, 0.1, na.rm = TRUE)
    weight_mat <- (mat <= threshold) * 1
    diag(weight_mat) <- 0
    nb_list <- spdep::mat2listw(weight_mat, style = "W")
    pval <- spdep::moran.test(res, nb_list)$p.value
    
    HTML(paste0(
      "<p><strong>p-value:</strong> ", signif(pval, 4), "</p>",
      "<p><strong>Interpretasi:</strong> ",
      if (pval < 0.05) {
        "Terdapat autokorelasi spasial pada residual. Hal ini menunjukkan adanya pola spasial yang belum ditangkap oleh model."
      } else {
        "Tidak terdapat autokorelasi spasial. Asumsi bebas spasial terpenuhi."
      }, "</p>"
    ))
  })
  
  output$catatan_rangkuman_model <- renderUI({
    HTML(paste0(
      "<p style='text-align:justify;'><strong>Model final</strong> yang digunakan dalam perhitungan estimasi adalah regresi linear dengan variabel respon ",
      "<em>LOWEDU</em> (persentase penduduk 15 tahun ke atas berpendidikan rendah), dan empat variabel prediktor yaitu: ",
      "<strong>ELDERLY</strong>, <strong>FHEAD</strong>, <strong>POVERTY</strong>, dan <strong>RENTED</strong>.</p>",
      
      "<p style='text-align:justify;'>Model ini telah melalui berbagai uji validasi:</p>",
      "<ul>",
      "<li><strong>Uji Normalitas:</strong> Lolos (residual berdistribusi normal).</li>",
      "<li><strong>Uji Multikolinearitas:</strong> Lolos (VIF < 10 untuk semua variabel).</li>",
      "<li><strong>Uji Heteroskedastisitas:</strong> Lolos (tidak terdapat heteroskedastisitas).</li>",
      "<li><strong>Uji Autokorelasi Spasial (Moran's I):</strong> Tidak lolos (terdapat indikasi autokorelasi spasial pada residual).</li>",
      "</ul>",
      
      "<p style='text-align:justify;'>Meskipun terdapat autokorelasi spasial, model ini tetap digunakan karena memiliki performa terbaik secara statistik dibandingkan kombinasi lainnya. ",
      "Namun, pengguna disarankan untuk mempertimbangkan aspek spasial lebih lanjut dalam analisis lanjutan (misalnya regresi spasial).</p>"
    ))
  })
  
  output$unduh_catatan_model <- downloadHandler(
    filename = function() {
      paste0("catatan_model_regresi_", Sys.Date(), ".docx")
    },
    content = function(file) {
      library(officer)
      
      doc <- read_docx()
      
      doc <- body_add_par(doc, "Ringkasan Uji Model Regresi", style = "heading 1")
      doc <- body_add_par(doc, "Model regresi final menggunakan variabel:", style = "Normal")
      doc <- body_add_par(doc, "Y = LOWEDU", style = "Normal")
      doc <- body_add_par(doc, "X = ELDERLY, FHEAD, POVERTY, RENTED", style = "Normal")
      
      doc <- body_add_par(doc, "Hasil Uji Asumsi:", style = "heading 2")
      doc <- body_add_par(doc, "- Uji Normalitas: Lolos (p-value > 0.05)", style = "Normal")
      doc <- body_add_par(doc, "- Uji Multikolinearitas: Lolos (VIF < 10)", style = "Normal")
      doc <- body_add_par(doc, "- Uji Heteroskedastisitas: Lolos (p-value > 0.05)", style = "Normal")
      doc <- body_add_par(doc, "- Uji Autokorelasi Spasial (Moran’s I): Tidak lolos (p-value < 0.05)", style = "Normal")
      
      doc <- body_add_par(doc, "Catatan:", style = "heading 2")
      doc <- body_add_par(doc, "Meskipun model belum lolos uji spasial, model ini dipilih karena secara umum memenuhi asumsi klasik dan memiliki performa terbaik dalam menjelaskan variabel respon.", style = "Normal")
      
      print(doc, target = file)
    }
  )
  
  
})
