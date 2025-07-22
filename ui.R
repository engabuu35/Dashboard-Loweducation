library(shiny)
library(shinydashboard)
library(plotly)

dashboardPage(
  skin = NULL,
  dashboardHeader(
    title = ("Loweducation"),
    tags$li(class = "dropdown",
            tags$style(HTML(".main-header .navbar { background-color: #2e7d32 !important;}"))
    )
  ),
  
  dashboardSidebar( width = 230,
                    sidebarMenu(id = "tabs",
                                
                                menuItem("Beranda", tabName = "beranda", icon = icon("home")),
                                menuItem("Eksplorasi Data", icon = icon("chart-bar"),
                                         menuSubItem("Tabel Data", tabName = "tabel", icon = icon("table")),
                                         menuSubItem("Grafik", tabName = "grafik", icon = icon("chart-area")),
                                         menuSubItem("Peta", tabName = "peta", icon = icon("map"))
                                ),
                                menuItem("Manajemen Data", tabName = "manajemen_data", icon = icon("database")),
                                menuItem("Uji Asumsi", icon = icon("check-circle"),
                                         menuSubItem("Normalitas", tabName = "normalitas", icon = icon("wave-square")),
                                         menuSubItem("Homogenitas", tabName = "homogenitas", icon = icon("balance-scale"))
                                ),
                                menuItem("Statistik Inferensia", icon = icon("flask"),
                                         menuSubItem("Uji Beda Rata-rata", tabName = "uji_rerata", icon = icon("equals")),
                                         menuSubItem("Uji Proporsi", tabName = "uji_proporsi", icon = icon("percentage")),
                                         menuSubItem("Uji Variansi", tabName = "uji_varians", icon = icon("compress-arrows-alt")),
                                         menuSubItem("Uji ANOVA", tabName = "anova", icon = icon("project-diagram"))
                                ),
                                menuItem("Regresi Linier", icon = icon("chart-area"),
                                         menuSubItem("Analisis", tabName = "rlb", icon = icon("line-chart")),
                                         menuSubItem("Estimasi", tabName = "estimasi", icon = icon("calculator")),
                                         menuSubItem("Catatan Estimasi", tabName = "catatan", icon = icon("pen"))
                                ),
                                menuItem("Tentang", tabName = "info", icon = icon("info"))
                                
                    )
  ),
  
  dashboardBody(
    useShinyjs(),
    tags$head(
      tags$link(
        href = "https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;600;700&display=swap",
        rel = "stylesheet"
      ),
      tags$link(rel = "stylesheet", type = "text/css", href = "style2.css")),
    
    tabItems(
      tabItem(tabName = "beranda",
              fluidPage(
                tags$div(style = "text-align: justify;",
                         
                         h2("Dashboard Analisis Pendidikan Rendah di Indonesia", 
                            style = "font-weight: bold; color: #4C6C29;"),
                         
                         p("Dashboard ini menyajikan analisis statistik terhadap tingkat pendidikan rendah (LOWEDU) di Indonesia berdasarkan data SUSENAS 2017. 
                          Analisis ini menelaah keterkaitan antara pendidikan rendah dengan berbagai indikator sosial dan ekonomi, 
                          meliputi persentase lansia (ELDERLY), kepala rumah tangga perempuan (FHEAD), ukuran rata-rata rumah tangga (FAMILYSIZE), 
                          kemiskinan (POVERTY), akses terhadap listrik (NOELECTRIC), dan status kepemilikan tempat tinggal (RENTED). 
                          Dashboard ini dikembangkan sebagai Projek Ujian Akhir Semester Mata Kuliah Komputasi Statistik dan sebagai alat bantu interaktif untuk mendukung pemahaman konsep statistik deskriptif, 
                          inferensia, serta analisis regresi linier dalam konteks pembangunan sosial dan pengentasan ketimpangan sosial.",
                           style = "text-align: justify; font-size: 14px;"),
                         br(),
                         
                         fluidRow(
                           valueBoxOutput("jumlah_distrik", width = 4),
                           valueBoxOutput("jumlah_variabel", width = 4),
                           valueBoxOutput("total_populasi", width = 4)
                         ),
                         
                         br(),
                         
                         box(
                           title = "Informasi Metadata",
                           width = 14,
                           solidHeader = TRUE,
                           status = "success",
                           
                           div(style = "font-size: 13px; font-family: 'Poppins', sans-serif; line-height: 1.6;",
                               
                               tags$ul(
                                 style = "list-style-type: disc; padding-left: 20px;",
                                 tags$li("Nama Dataset: Social Vulnerability Data (Indonesia), disusun oleh Kurniawan et al. (2022), tersedia melalui jurnal Elsevier Data in Brief."),
                                 tags$li("Jumlah Distrik: 511 kabupaten/kota (berbasis data SUSENAS 2017 dan peta spasial 2013)."),
                                 tags$li("Variabel Utama: LOWEDU, POVERTY, ELDERLY, FHEAD, RENTED."),
                                 tags$li("Sumber Data: SUSENAS 2017, Proyeksi Penduduk 2017 BPS, dan peta jarak spasial antar distrik."),
                                 tags$li("Format Data: CSV dan GeoJSON, termasuk matriks jarak antar wilayah administratif."),
                                 tags$li("Tujuan: Menganalisis faktor-faktor sosial yang berkontribusi terhadap rendahnya tingkat pendidikan di Indonesia."),
                                 tags$li(HTML('Tautan Publikasi: <a href="https://doi.org/10.1016/j.dib.2021.107743" target="_blank">Data in Brief (Elsevier)</a>')),
                                 tags$li(HTML('Akses Dataset: 
                                  <a href="https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/sovi_data.csv" target="_blank">SOVI Data</a>, 
                                  <a href="https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/distance.csv" target="_blank">Matriks Jarak</a>')),
                                 tags$li("Lisensi: CC-BY (Creative Commons Attribution)")
                               ),
                               
                               tags$hr(style = "border-top: 1px solid #ccc;"),
                               tags$h5("Deskripsi Variabel:", style = "font-weight: bold; margin-top: 15px;"),
                               
                               tags$div(style = "overflow-x:auto;",
                                        tags$table(
                                          style = "width:100%; font-size:13px; border-collapse: collapse;",
                                          border = 1,
                                          tags$thead(
                                            style = "background-color:#e8f5e9; font-weight:bold;",
                                            tags$tr(
                                              tags$th("Label"),
                                              tags$th("Nama Variabel"),
                                              tags$th("Deskripsi")
                                            )
                                          ),
                                          tags$tbody(
                                            style = "text-align: left;",
                                            tags$tr(tags$td("LOWEDU"), tags$td("Pendidikan Rendah"), tags$td("Persentase penduduk usia 15 tahun ke atas dengan pendidikan rendah.")),
                                            tags$tr(tags$td("ELDERLY"), tags$td("Penduduk Lansia"), tags$td("Persentase penduduk usia 65 tahun ke atas.")),
                                            tags$tr(tags$td("FHEAD"), tags$td("Kepala Rumah Tangga Perempuan"), tags$td("Persentase rumah tangga dengan kepala keluarga perempuan.")),
                                            tags$tr(tags$td("FAMILYSIZE"), tags$td("Ukuran Rumah Tangga"), tags$td("Rata-rata jumlah anggota rumah tangga dalam satu rumah tangga.")),
                                            tags$tr(tags$td("POVERTY"), tags$td("Kemiskinan"), tags$td("Persentase penduduk miskin.")),
                                            tags$tr(tags$td("NOELECTRIC"), tags$td("Tanpa Listrik"), tags$td("Persentase rumah tangga tanpa sumber listrik.")),
                                            tags$tr(tags$td("RENTED"), tags$td("Menyewa Rumah"), tags$td("Persentase rumah tangga yang menyewa tempat tinggal."))
                                          )
                                        )
                               )
                               
                           )
                         ),
                         div(style = "text-align: right;",
                             downloadButton("unduh_beranda", "Unduh Halaman Ini (.docx)", class = "btn-success")
                         )
                         
                )
              )
      ),
      
      tabItem(tabName = "tabel",
              fluidRow(
                box(title = "Tabel Data", width = 12, status = "primary", solidHeader = TRUE,
                    DT::dataTableOutput("tabel_data_asli"),
                    div(
                      style = "margin-top: 10px; text-align: right; width: 100%;",
                      downloadButton("unduh_data_asli", "Unduh Excel (.xlsx)", class = "btn btn-success"),
                      downloadButton("unduh_word_data_asli", "Unduh Ringkasan (.docx)", class = "btn btn-primary", style = "margin-left: 10px;")
                    )
                ),
                fluidRow(
                  valueBoxOutput("stat_lowedu", width = 3),
                  valueBoxOutput("stat_elderly", width = 3),
                  valueBoxOutput("stat_fhead", width = 3),
                  valueBoxOutput("stat_rented", width = 3)
                ),
                fluidRow(
                  div(
                    style = "display: flex; justify-content: center; gap: 20px;",
                    valueBoxOutput("stat_familysize", width = NULL),
                    valueBoxOutput("stat_poverty", width = NULL),
                    valueBoxOutput("stat_noelectric", width = NULL)
                  )
                )
              )
      ),

      tabItem(tabName = "grafik",
              fluidRow(
                box(
                  title = "Filter",
                  width = 6,
                  status = "success",
                  solidHeader = TRUE,
                  
                  selectInput("filter_prov", "Pilih Provinsi:", 
                              choices = c("ALL", sort(unique(data$nmprov))), 
                              selected = "ALL", multiple = TRUE),
                  
                  selectInput("filter_kab", "Pilih Kabupaten:", 
                              choices = sort(unique(data$nmkab)), 
                              selected = "NULL", multiple = TRUE),
                  
                  selectInput("var_statistik", "Pilih Variabel Analisis:", 
                              choices = c("LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED"), 
                              selected = "LOWEDU"),
                  actionButton("reset_filter", "Reset Filter", icon = icon("undo"), class = "btn btn-warning")
                ),
                
                box(
                  title = "Histogram",
                  width = 6,
                  status = "success",
                  solidHeader = TRUE,
                  
                  plotlyOutput("plot_histogram", height = "210px"),
                  div(style = "text-align: right; margin-top: 10px;",
                      downloadButton("unduh_histogram", "Unduh Histogram (.png)", class = "btn btn-success")
                  )
                )
              ),
              
              fluidRow(
                column(
                  width = 6,
                  box(
                    title = "Boxplot",
                    width = NULL,
                    status = "success",
                    solidHeader = TRUE,
                    plotlyOutput("plot_boxplot", height = "290px"),
                    div(style = "text-align: right; margin-top: 10px;",
                        downloadButton("unduh_boxplot", "Unduh Boxplot (.png)", class = "btn btn-success")
                    )
                  )
                ),
                
                column(
                  width = 6,
                  box(
                    title = "Interpretasi",
                    width = NULL,
                    status = "info",
                    solidHeader = TRUE,
                    uiOutput("interpretasi_grafik"),
                    div(style = "text-align: right; margin-top: 10px;",
                        downloadButton("unduh_interpretasi_grafik", "Unduh Interpretasi (.docx)", class = "btn btn-primary"),
                        downloadButton("unduh_halaman_grafik", "Unduh Halaman ini (.docx)", class = "btn-primary")
                    )
                  )
                ),
                valueBoxOutput("stat_grafik", width = 6)
              )
      ),
      

      tabItem(tabName = "peta",
              fluidRow(
                box(
                  width = 9,
                  title = "Peta Sebaran Variabel",
                  solidHeader = TRUE,
                  status = "success",
                  leafletOutput("peta_populasi", height = "550px"),
                  
                ),
                box(
                  width = 3,
                  title = "Filter Variabel",
                  solidHeader = TRUE,
                  status = "warning",
                  selectInput("var_peta", "Pilih Variabel:",
                              choices = c("LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED"),
                              selected = "LOWEDU"),
                  selectInput("filter_prov_peta", "Pilih Provinsi:",
                              choices = c("ALL", sort(unique(geo_map$nmprov))),
                              selected = "ALL"),
                  uiOutput("filter_kab_ui_peta"),
                  div(style = "text-align: right; margin-top: 10px;",
                      downloadButton("unduh_peta_png", "Unduh Peta (.png)", class = "btn btn-success"),
                  )
                ),
                
                box(
                  width = 3,
                  title = "Interpretasi",
                  solidHeader = TRUE,
                  status = "warning",
                  uiOutput("interpretasi_peta"),
                  div(style = "text-align: right; margin-top: 10px;",
                      downloadButton("unduh_interpretasi_peta", "Unduh Interpretasi (.docx)", class = "btn btn-primary", style = "margin-left:10px;")
                  ),
                  div(style = "text-align: right; margin-top: 5px;",
                      downloadButton("unduh_halaman_peta", "Unduh Halaman ini (.docx)", class = "btn btn-info")
                  )
                  
                )
              )
      ),
      
      
      tabItem(tabName = "manajemen_data",
              fluidRow(
                box(
                  width = 6,
                  title = "Kategorisasi Variabel Kontinu",
                  solidHeader = TRUE,
                  status = "success",
                  style = "padding-bottom: 9px;",
                  fluidRow(
                    column(4,
                           selectInput("var_kategorik", 
                                       label = "Pilih Variabel:",
                                       choices = c("LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED"),
                                       multiple = TRUE,
                                       selected = "LOWEDU")
                    ),
                    column(4,
                           numericInput("banyak_kategori", 
                                        label = "Jumlah Kategori:", 
                                        value = 2, min = 2, max = 10)
                    ),
                    column(4,
                           checkboxGroupInput("metode_kat", 
                                              label = "Metode Kategorisasi:",
                                              choices = c("Interval Sama Lebar", "Kuartil", "Natural Breaks"),
                                              selected = "Interval Sama Lebar")
                    )
                  ),
                  actionButton("proses_kat", label = tagList(icon("calculator"), "Proses Kategorisasi"), class = "btn btn-success")
                ),
                
                box(
                  width = 6,
                  title = "Info",
                  solidHeader = TRUE,
                  status = "info",
                  p("Kolom hasil disimpan dengan format: nama_variabel + metode."),
                  tags$b("Keterangan:"),
                  tags$ul(
                    tags$li("SL: Sama Lebar"),
                    tags$li("Q: Kuartil"),
                    tags$li("J: Natural Breaks (Jenks)"),
                    br(),
                  )
                )
              ),
              
              fluidRow(
                box(
                  width = 12,
                  title = "Tabel Hasil Manajemen Data",
                  solidHeader = TRUE,
                  status = "success",
                  DT::dataTableOutput("tabel_manajemen_data"),
                  div(
                    style = "margin-top: 10px; text-align: right; width: 100%;",
                    downloadButton("unduh_kategori", "Unduh Excel (.xlsx)", class = "btn btn-success")
                  )
                )
              ),
              
              fluidRow(
                box(
                  width = 6,
                  title = "Cut Point Hasil Kategorisasi",
                  solidHeader = TRUE,
                  status = "primary",
                  uiOutput("cutpoint_info")
                ),
                box(
                  width = 6,
                  title = "Penjelasan Hasil Kategorisasi",
                  solidHeader = TRUE,
                  status = "primary",
                  uiOutput("interpretasi_cutpoint"),
                  downloadButton("unduh_interpretasi_kat", "Unduh Interpretasi (.docx)", class = "btn btn-primary"),
                  downloadButton("unduh_halaman_manajemen", "Unduh Halaman ini (.docx)", class = "btn btn-info")
                )
              )
      ),
      
      # Uji Asumsi : Normalitas
      tabItem(tabName = "normalitas",
              fluidRow(
                column(width = 4,
                       style = "padding-left:0;padding-right:0;",
                       box(
                         width = 12,
                         title = "Input Uji Normalitas",
                         solidHeader = TRUE,
                         status = "success",
                         selectInput("var_normalitas", "Pilih Variabel untuk Uji Normalitas:",
                                     choices = c("LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED"),
                                     selected = "LOWEDU")
                       ),
                       box(
                         width = 12,
                         title = "Hasil Uji Shapiro-Wilk",
                         solidHeader = TRUE,
                         status = "warning",
                         verbatimTextOutput("hasil_shapiro")
                       ),
                       box(
                         width = 12,
                         title = "Interpretasi",
                         solidHeader = TRUE,
                         status = "info",
                         HTML(
                           "<b>Hipotesis:</b><br/>
                            H₀: Data berdistribusi normal<br/>
                            H₁: Data tidak berasal  berdistribusi normal<br/><br/>"
                         ),
                         uiOutput("interpretasi_normalitas"),
                         div(
                           style = "margin-top: 10px; text-align: right;",
                           downloadButton("unduh_normalitas", "Unduh Hasil (.docx)", class = "btn btn-success")
                         )
                       )
                       
                ),
                 box(
                   style = "padding-bottom: 18px;",
                   width = 8,
                   title = "Visualisasi Distribusi",
                   solidHeader = TRUE,
                   status = "primary",
                   plotOutput("plot_normalitas", height = "300px"),
                   plotOutput("qqplot_normalitas", height = "300px")
                 )
              )
      ),
      
      tabItem(tabName = "homogenitas",
              fluidRow(
                column(width = 6,
                       style = "padding-left:0;padding-right:0;",
                       box(
                         width = 12,
                         title = "Input Uji Homogenitas",
                         solidHeader = TRUE,
                         status = "success",
                         selectInput("var_homogenitas", "Pilih Variabel Target:",
                                     choices = c("LOWEDU", "ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED"),
                                     selected = "LOWEDU"),
                         uiOutput("pilihan_group_uji")
                       ),
                       box(
                         width = 12,
                         title = "Hasil Uji Levene",
                         solidHeader = TRUE,
                         status = "warning",
                         verbatimTextOutput("hasil_levene")
                       ),
                       box(
                         width = 12,
                         title = "Interpretasi",
                         solidHeader = TRUE,
                         status = "info",
                         HTML(
                           "<p><strong>Hipotesis:</strong><br/>
                            H₀: Varians antar kelompok sama (homogen)<br/>
                            H₁: Varians antar kelompok berbeda (tidak homogen)</p>"
                         ),
                         uiOutput("interpretasi_levene"),
                         div(
                           style = "margin-top: 10px; text-align: right;",
                           downloadButton("unduh_levene", "Unduh Hasil (.docx)", class = "btn btn-success")
                         )
                         
                       )
                ),
                column(width = 6,
                       style = "padding-left:0;padding-right:0; text-align: justify;",
                       box(
                         width = 12,
                         title = "Info",
                         solidHeader = TRUE,
                         status = "warning",
                         p("Uji Lavene tidak bisa dilakukan jika tidak ada variabel kategorik yang tersedia.
                           Silakan lakukan proses kategorisasi terlebih dahulu melalui tab 'Manajemen Data' sebelum menjalankan uji homogenitas.")
                       ),
                       box(
                         style = "padding-bottom: 17px",
                         width = 12,
                         title = "Visualisasi Boxplot",
                         solidHeader = TRUE,
                         status = "primary",
                         plotOutput("plot_levene", height = "450px")
                       )
                )
              )
      ),
      
 
      tabItem(tabName = "uji_rerata",
              fluidRow(
                box(width = 4, 
                    title = "Input Uji", 
                    status = "primary", solidHeader = TRUE,
                    radioButtons("jenis_uji_rerata", "Pilih Jenis Uji:",
                                 choices = c("Uji Satu Rata-rata" = "satu",
                                             "Uji Dua Rata-rata" = "dua",
                                             "Uji Rata-rata Berpasangan" = "pasangan")),
                    
                    conditionalPanel(
                      condition = "input.jenis_uji_rerata == 'satu'",
                      selectInput("var_rerata1", "Pilih Variabel:", choices = NULL),
                      numericInput("mu", "Nilai Hipotesis (μ₀):", value = 0),
                      radioButtons("arah_uji1", "Arah Uji:",
                                   choices = c("Dua Arah" = "two.sided",
                                               "Satu Arah Kiri (μ < μ₀)" = "less",
                                               "Satu Arah Kanan (μ > μ₀)" = "greater"))
                    ),
                    
                    conditionalPanel(
                      condition = "input.jenis_uji_rerata == 'dua'",
                      selectInput("var_rerata2", "Pilih Variabel:", choices = NULL),
                      
                      uiOutput("group_rerata_ui"),
                      
                      radioButtons("equal_var", "Asumsi Variansi:",
                                   choices = c("Sama" = TRUE,
                                               "Tidak Sama" = FALSE)),
                      radioButtons("arah_uji2", "Arah Uji:",
                                   choices = c("Dua Arah" = "two.sided",
                                               "Satu Arah Kiri (μ₁ < μ₂)" = "less",
                                               "Satu Arah Kanan (μ₁ > μ₂)" = "greater"))
                    ),
                    
                    conditionalPanel(
                      condition = "input.jenis_uji_rerata == 'pasangan'",
                      selectInput("paired_var1", "Variabel Sebelum Perlakuan", choices = NULL),
                      selectInput("paired_var2", "Variabel Sesudah Perlakuan", choices = NULL),
                      radioButtons("arah_uji3", "Arah Uji:",
                                   choices = c("Dua Arah" = "two.sided",
                                               "Satu Arah Kiri (μD < 0)" = "less",
                                               "Satu Arah Kanan (μD > 0)" = "greater"))
                    ),
                    
                    actionButton("proses_uji_rerata", 
                                 label = tagList(icon("flask"), "Lakukan Uji"),
                                 class = "btn-success")
                ),
                
                box(width = 8, title = "Hasil Uji Beda Rata-rata", status = "primary", solidHeader = TRUE,
                    verbatimTextOutput("hasil_uji_rerata"),
                    htmlOutput("interpretasi_uji_rerata"),
                    downloadButton("unduh_uji_rerata_docx", "Unduh Hasil (.docx)", class = "btn-success float-right")
                )
              )
      ),
      
      
      tabItem(tabName = "uji_proporsi",
              fluidRow(
                box(width = 4, title = "Input Uji Proporsi", status = "primary", solidHeader = TRUE,
                    
                    radioButtons("jenis_uji_proporsi", "Pilih Jenis Uji:",
                                 choices = c("Uji Proporsi Satu Kelompok" = "satu",
                                             "Uji Proporsi Dua Kelompok"   = "dua")),
                    
                    conditionalPanel(
                      condition = "input.jenis_uji_proporsi == 'satu'",
                      selectInput("var_kat_proporsi", "Variabel Kategorik (2 kategori):", choices = NULL),
                      uiOutput("nilai_kat_proporsi_ui"),
                      numericInput("nilai_p0_proporsi", "Proporsi Acuan (p₀):", value = 0.5, min = 0, max = 1, step = 0.01),
                      radioButtons("arah_uji_proporsi", "Arah Uji:",
                                   choices = c("Dua Arah (≠)" = "two.sided",
                                               "Kurang dari (<)" = "less",
                                               "Lebih dari (>)"  = "greater"))
                    ),
                    
                    conditionalPanel(
                      condition = "input.jenis_uji_proporsi == 'dua'",
                      selectInput("var_kat_proporsi2", "Variabel Kategorik (2 kategori):", choices = NULL),
                      uiOutput("nilai_kat_proporsi2_ui"),
                      uiOutput("group_proporsi_ui"),
                      radioButtons("arah_uji_proporsi2", "Arah Uji:",
                                   choices = c("Dua Arah (≠)"             = "two.sided",
                                               "Kelompok 1 < Kelompok 2" = "less",
                                               "Kelompok 1 > Kelompok 2" = "greater"))
                    ),
                    
                    numericInput("alpha_proporsi", "Tingkat Signifikansi (α):", value = 0.05, min = 0.001, max = 0.2, step = 0.001),
                    actionButton("proses_uji_proporsi", 
                                 label = tagList(icon("flask"), "Lakukan Uji"),
                                 class = "btn-success")
                    
                ),
                
                box(width = 8, title = "Hasil Uji Proporsi", status = "primary", solidHeader = TRUE,
                    verbatimTextOutput("hasil_uji_proporsi"),
                    htmlOutput("interpretasi_uji_proporsi"),
                    downloadButton("unduh_uji_proporsi_docx", "Unduh Hasil (.docx)", class = "btn-warning float-right")
                )
              )
      ),
      
      tabItem(tabName = "uji_varians",
              fluidRow(
                box(width = 4, title = "Input Uji Varians", status = "primary", solidHeader = TRUE,
                    radioButtons("jenis_uji_varians", "Pilih Jenis Uji:",
                                 choices = c("Uji Varians Satu Kelompok" = "satu",
                                             "Uji Varians Dua Kelompok" = "dua")),
                    
                    conditionalPanel(
                      condition = "input.jenis_uji_varians == 'satu'",
                      selectInput("var_varian_satu", "Variabel Numerik:", choices = NULL),
                      numericInput("nilai_varian0", "Varians Acuan (σ²₀):", value = 1, min = 0.0001),
                      radioButtons("arah_uji_varian_satu", "Arah Uji:",
                                   choices = c("Dua Arah (≠)" = "two.sided",
                                               "Kurang dari (<)" = "less",
                                               "Lebih dari (>)" = "greater"))
                    ),
                    
                    conditionalPanel(
                      condition = "input.jenis_uji_varians == 'dua'",
                      selectInput("var_varian_dua", "Variabel Numerik:", choices = NULL),
                      uiOutput("group_varian_dua_ui"),
                      radioButtons("arah_uji_varian_dua", "Arah Uji:",
                                   choices = c("Dua Arah (≠)" = "two.sided",
                                               "Kelompok 1 < Kelompok 2" = "less",
                                               "Kelompok 1 > Kelompok 2" = "greater"))
                    ),
                    
                    numericInput("alpha_varians", "Tingkat Signifikansi (α):", value = 0.05, min = 0.001, max = 0.2),
                    actionButton("proses_uji_varians", label = tagList(icon("flask"), "Lakukan Uji"), class = "btn-success")
                ),
                
                box(width = 8, title = "Hasil Uji Varians", status = "primary", solidHeader = TRUE,
                    verbatimTextOutput("hasil_uji_varians"),
                    htmlOutput("interpretasi_uji_varians"),
                    downloadButton("unduh_uji_varians_docx", "Unduh Hasil (.docx)", class = "btn-warning float-right")
                )
              )
      ),
      
      tabItem(tabName = "anova",
              fluidRow(
                box(width = 4, title = "Input Uji ANOVA", status = "primary", solidHeader = TRUE,
                    
                    radioButtons("jenis_uji_anova", "Pilih Jenis Uji:",
                                 choices = c("ANOVA Satu Arah" = "satu",
                                             "ANOVA Dua Arah" = "dua")),
                    
                    # Input ANOVA Satu Arah
                    conditionalPanel(
                      condition = "input.jenis_uji_anova == 'satu'",
                      selectInput("anova_var_numerik1", "Variabel Numerik (Respon):", choices = NULL),
                      uiOutput("anova_group1_ui")  # Gunakan UI dinamis untuk faktor kategorik
                    ),
                    
                    # Input ANOVA Dua Arah
                    conditionalPanel(
                      condition = "input.jenis_uji_anova == 'dua'",
                      selectInput("anova_var_numerik2", "Variabel Numerik (Respon):", choices = NULL),
                      uiOutput("anova_group2a_ui"),  # Faktor Kategorik A
                      uiOutput("anova_group2b_ui"),  # Faktor Kategorik B
                      checkboxInput("anova_interaksi", "Sertakan Interaksi Faktor", value = FALSE)
                    ),
                    
                    numericInput("alpha_anova", "Tingkat Signifikansi (α):", value = 0.05, min = 0.001, max = 0.2),
                    actionButton("proses_uji_anova", label = tagList(icon("flask"), "Lakukan Uji"), class = "btn-success")
                ),
                
                box(width = 8, title = "Hasil Uji ANOVA", status = "primary", solidHeader = TRUE,
                    verbatimTextOutput("hasil_uji_anova"),
                    htmlOutput("interpretasi_uji_anova"),
                    downloadButton("unduh_uji_anova_docx", "Unduh Hasil (.docx)", class = "btn-warning float-right")
                    
                )
              )
      ),
      
      
      tabItem(tabName = "rlb",
              fluidRow(
                 box(width = 4, title = "Input Model Regresi", status = "primary", solidHeader = TRUE,
                     selectInput("y_rlb", "Variabel Respon (Y):", 
                                 choices = "LOWEDU", selected = "LOWEDU"),
                     selectInput("x_rlb", "Pilih Variabel Prediktor (X):", 
                                 choices = c("ELDERLY", "FHEAD", "FAMILYSIZE", "POVERTY", "NOELECTRIC", "RENTED"), 
                                 multiple = TRUE),
                     actionButton("proses_regresi", label = tagList(icon("play"), "Lakukan Regresi"), class = "btn-success"),
                     downloadButton("unduh_regresi_docx", "Unduh Hasil (.docx)", class = "btn btn-primary")
                     
                 ),
                
                column(width = 8,
                       style = "padding-left:0;padding-right:0;",
                       tabBox(width = 12, id = "tab_regresi", side = "left",
                              tabPanel("Ringkasan",
                                       verbatimTextOutput("output_regresi"),
                                       htmlOutput("interpretasi_regresi")
                              ),
                              tabPanel("Normalitas",
                                       verbatimTextOutput("uji_normalitas"),
                                       htmlOutput("interpretasi_normalitas")
                              ),
                              tabPanel("Multikolinearitas",
                                       verbatimTextOutput("uji_vif"),
                                       htmlOutput("interpretasi_vif")
                              ),
                              tabPanel("Heteroskedastisitas",
                                       verbatimTextOutput("uji_breusch"),
                                       htmlOutput("interpretasi_breusch")
                              ),
                              tabPanel("Autokorelasi Spasial",
                                       verbatimTextOutput("uji_moran"),
                                       htmlOutput("interpretasi_moran")
                              )
                       )
                )
              )
      ),
      
      
      tabItem(tabName = "estimasi",
              fluidRow(
                # Kolom kiri: Input kalkulator
                box(width = 4, title = "Kalkulator Estimasi LOWEDU", status = "primary", solidHeader = TRUE,
                    numericInput("input_elderly", "Persentase Lansia (ELDERLY)", value = NULL, min = 0, step = 0.1),
                    numericInput("input_fhead", "Persentase Kepala Keluarga Perempuan (FHEAD)", value = NULL, min = 0, step = 0.1),
                    numericInput("input_poverty", "Persentase Kemiskinan (POVERTY)", value = NULL, min = 0, step = 0.1),
                    numericInput("input_rented", "Persentase Rumah Sewa (RENTED)", value = NULL, min = 0, step = 0.1),
                    br(),
                    actionButton("hitung_estimasi", label = tagList(icon("calculator"), "Hitung Estimasi"), class = "btn btn-success")
                ),
                
                # Kolom kanan: Output hasil dan unduh
                box(width = 8, title = "Hasil Estimasi dan Interpretasi", status = "success", solidHeader = TRUE,
                    htmlOutput("hasil_estimasi"),
                    br(),
                    downloadButton("unduh_estimasi", "Unduh Estimasi (.docx)", class = "btn btn-primary")
                ),
                box(
                  width = 8,
                  title = "Ringkasan Model dan Interpretasi",
                  status = "info",
                  solidHeader = TRUE,
                  htmlOutput("interpretasi_estimasi")
                )
              )
      ),
      
      tabItem(tabName = "catatan",
              fluidRow(
                box(title = "Ringkasan Model Regresi untuk Estimasi", width = 12, status = "primary", solidHeader = TRUE,
                    verbatimTextOutput("catatan_ringkasan_model"),
                    htmlOutput("catatan_interpretasi_model")
                )
              ),
              fluidRow(
                box(title = "Uji Normalitas Residual", width = 6, status = "success", solidHeader = TRUE,
                    verbatimTextOutput("catatan_uji_normalitas"),
                    htmlOutput("catatan_interpretasi_normalitas")
                ),
                box(title = "Uji Multikolinearitas (VIF)", width = 6, status = "success", solidHeader = TRUE,
                    verbatimTextOutput("catatan_uji_vif"),
                    htmlOutput("catatan_interpretasi_vif")
                )
              ),
              fluidRow(
                box(title = "Uji Autokorelasi Spasial (Moran’s I)", width = 6, status = "success", solidHeader = TRUE,
                    verbatimTextOutput("catatan_uji_moran"),
                    htmlOutput("catatan_interpretasi_moran")
                ),
                
                box(title = "Uji Heteroskedastisitas (Breusch-Pagan)", width = 6, status = "success", solidHeader = TRUE,
                    verbatimTextOutput("catatan_uji_breusch"),
                    htmlOutput("catatan_interpretasi_breusch")
                )
                
              ),
              fluidRow(
                box(title = "Rangkuman Uji Model Regresi", width = 12, status = "success", solidHeader = TRUE,
                    htmlOutput("catatan_rangkuman_model"),
                    br(),
                    downloadButton("unduh_catatan_model", "Unduh Ringkasan Model (.docx)", class = "btn btn-primary")
                )
              )
      ),
      
      
      tabItem(tabName = "info",
              fluidRow(
                 box(width = 3, title = "Profil", status = "primary", solidHeader = TRUE,
                     div(style = "text-align: center;",
                         img(src = "foto.jpg", height = "250px", style = "border-radius: 7px; margin-bottom: 10px; margin-top:10px;"),
                         h4(strong("Arif Budiman")),
                         p("NIM : 222312994", style = "font-size: 13px; padding-bottom:9px;")
                     )
                 ),
                 
                 box(
                   width = 9, 
                   title = "Tentang Saya", 
                   status = "success", 
                   solidHeader = TRUE,
                   div(
                     style = "text-align: justify; font-size: 14px; line-height: 1.7;",
                     p("Halo! Saya Arif Budiman, biasa dipanggil Abu, mahasiswa aktif Program Studi Komputasi Statistik Politeknik Statistika STIS."),
                     p("Dashboard ini saya kembangkan sebagai bagian dari Proyek Ujian Akhir Semester Mata Kuliah Komputasi Statistik. Dashboard ini adalah 'kloningan' dari dashboard sebelumnya yang saya bangun bersama kelompok untuk tugas akhir, 
                       yang kini dikembangkan ulang secara mandiri dengan fitur yang lebih lengkap dan dengan waktu yang lebih singkat."),
                     p("Saya percaya bahwa data bukan hanya angka-tetapi cerita yang bisa mengarahkan perubahan. Menguasai data menguasai dunia!"),
                     p("Dengan berbagai fitur statistik dan visualisasi interaktif, saya berharap dashboard ini dapat menjadi salah satu kontribusi kecil dalam memahami persoalan sosial berbasis data. Saya sangat mengharapkan kritik dan saran yang membangun demi pengembangan dasboard ini maupun diri saya ke arah yang lebih baik."),
                     p("Kontak:"),
                     tags$ul(
                       tags$li("Email: 222312994@stis.ac.id"),
                       tags$li(HTML('GitHub: <a href="https://github.com/engabuu35" target="_blank">github.com/engabuu35</a>'))
                     )
                   )
                 )
                 
              )
      )
      
    )
    
  )
)