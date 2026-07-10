library(dplyr)
library(readxl)
library(janitor)

# ==========================================================
# Leer archivo
# ==========================================================

base <- readxl::read_xlsx(
  "C:/Users/brittany.pereo/Downloads/metas_planes_justicia.xlsx"
) %>%
  clean_names()

# ==========================================================
# Catálogo de CLUES
# ==========================================================

clues_info <- base %>%
  transmute(
    id = clues_imb,
    clues = clues_imb,
    nombre = nombre_de_la_unidad,
    entidad,
    nivel_atencion,
    categoria_gerencial,
    estatus_de_operacion
  ) %>%
  distinct()

# ==========================================================
# Metas por CLUES
# ==========================================================

metas_clues <- base %>%
  transmute(
    clues = clues_imb,
    meta_general_anual,
    meta_especialidad_anual,
    meta_cirugia_anual,
    meta_egresos_anual
  ) %>%
  distinct()

# ==========================================================
# Guardar ambos objetos en un .rda
# ==========================================================

save(
  clues_info,
  metas_clues,
  file = "C:/Users/brittany.pereo/Downloads/sysdata.rda",
  compress = "xz"
)


cubos_planes <- readxl::read_xlsx(
  "C:/Users/brittany.pereo/GitHub/planes_de_justicia/inst/app/data/cubos_completos_plan_justicia.xlsx"
)

arrow::write_parquet(cubos_planes,
                    "C:/Users/brittany.pereo/GitHub/planes_de_justicia/inst/app/data/Cubos_completos_2020_2025.parquet"
                    
                    )


procedimientos_personas <- readxl::read_xlsx(
  "C:/Users/brittany.pereo/GitHub/planes_de_justicia/inst/app/data/procedimientos_personas_plan_justicia.xlsx"
)

arrow::write_parquet(procedimientos_personas,
                     "C:/Users/brittany.pereo/GitHub/planes_de_justicia/inst/app/data/procedimientos_personas.parquet"
                     
)

