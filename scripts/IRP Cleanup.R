library(tidyverse)
library(tidyxl)
library(janitor)

irp <- xlsx_cells("IRP DATA 10-1-25 SR saved copy.xlsx")

# Clean up nameplate capacity

capacity_cleaner <- function(sheet_name, utility){
  
  fuel_types <- irp |> 
    filter(sheet == sheet_name, 
           row %in% c(1:17) & col == 1) |> 
    select(character) |> 
    row_to_names(row_number = 1) |> 
    rename(fuel_type = 1)
  
  nameplate_values <- irp |> 
    filter(sheet == sheet_name, 
           row %in% c(2:17) & col %in% c(2:15)) |> 
    select(row, col, numeric) |> 
    pivot_wider(names_from = col, values_from = numeric)
  
  nameplate_colnames <- irp |> 
    filter(sheet == sheet_name, 
           row == 1, col %in% c(1:15)) |> 
    mutate(values_use = case_when(
      data_type == "character" ~ character, 
      data_type == "numeric" ~ content
    )) |> 
    pull(values_use)
  
  colnames(nameplate_values) <- nameplate_colnames
  
  nameplate_projections <- fuel_types |> 
    bind_cols(nameplate_values) |>
    select(-c(`Nameplate (MW)`)) |> 
    pivot_longer(!fuel_type, names_to = "year", values_to = "nameplate_mw") |> 
    filter(!(fuel_type %in% c("Total", "Contract:Purchase", "Contract:Sale", "Demand:Distributed Generation", 
                              "Demand:Energy Efficiency", "Other:Other", "Renewable:Biomass", "Renewable:Landfill", "Demand:Demand Response"))) |> 
    mutate(fuel_type = case_when(
      fuel_type == "Coal:Conventional" ~ "Coal",
      fuel_type %in% c("Gas/Oil:Combined Cycle", "Gas/Oil:Combustion Turbine") ~ "Natural Gas",
      fuel_type == "Hydro:Hydroelectric" ~ "Hydroelectric",
      fuel_type == "Nuclear:Nuclear" ~ "Nuclear",
      fuel_type == "Renewable:Solar PV" ~ "Solar",
      fuel_type == "Renewable:Wind" ~ "Wind",
      fuel_type == "Renewable:Biomass" ~ "Biomass",
      fuel_type == "Storage:Battery" ~ "Battery"
    ), 
    year = as.numeric(year)) |> 
    group_by(fuel_type, year) |> 
    summarize(nameplate_mw = sum(nameplate_mw)) |> 
    mutate(
      
    capacity_label = case_when(
      year == max(year) ~ round(nameplate_mw), 
      TRUE ~ NA
    ), 
    
    fuel_label = case_when(
      year == max(year) ~ fuel_type, 
      TRUE ~ ""
    ),
    
    nameplate_mw = replace_na(nameplate_mw, 0),
    
    utility = utility)
  
  return(nameplate_projections)
  
  
}

nameplate_projections <- capacity_cleaner(sheet_name = "Xcel", utility = "Xcel Energy") |> 
  bind_rows(capacity_cleaner(sheet_name = "GRE", utility = "Great River Energy")) |> 
  bind_rows(capacity_cleaner(sheet_name = "MP", utility = "Minnesota Power")) |> 
  bind_rows(capacity_cleaner(sheet_name = "OTP", utility = "Otter Tail Power"))

#write_csv(nameplate_projections, "C:/Users/dduffy/OneDrive - State of Minnesota - MN365/Quad 2.0 Data - Documents/data/irp_nameplate_projections.csv", na = "")

# Clean up Energy Mix

mix_cleaner <- function(sheet_name, utility){
  
  fuel_types <- irp |> 
    filter(sheet == sheet_name, 
           row %in% c(19:35) & col == 1) |> 
    select(character) |> 
    row_to_names(row_number = 1) |> 
    rename(fuel_type = 1)
  
  
  production_values <- irp |> 
    filter(sheet == sheet_name, 
           row %in% c(20:35) & col %in% c(2:15)) |> 
    select(row, col, numeric) |> 
    pivot_wider(names_from = col, values_from = numeric)
  
  production_colnames <- irp |> 
    filter(sheet == sheet_name, 
           row == 19, col %in% c(1:15)) |> 
    mutate(values_use = case_when(
      data_type == "character" ~ character, 
      data_type == "numeric" ~ content
    )) |> 
    pull(values_use)
  
  colnames(production_values) <- production_colnames
  
  energy_mix <- fuel_types |> 
    bind_cols(production_values) |>
    select(-c(`Energy Mix (GWh)`)) |> 
    pivot_longer(!fuel_type, names_to = "year", values_to = "production_gwh") |> 
    filter(!(fuel_type %in% c("Total", "Contract:Purchase", "Contract:Sale", "Demand:Distributed Generation", 
                              "Demand:Energy Efficiency", "Other:Other", "Renewable:Landfill", "Demand:Demand Response"))) |> 
    mutate(fuel_type = case_when(
      fuel_type == "Coal:Conventional" ~ "Coal",
      fuel_type %in% c("Gas/Oil:Combined Cycle", "Gas/Oil:Combustion Turbine") ~ "Natural Gas",
      fuel_type == "Hydro:Hydroelectric" ~ "Hydroelectric",
      fuel_type == "Nuclear:Nuclear" ~ "Nuclear",
      fuel_type == "Renewable:Solar PV" ~ "Solar",
      fuel_type == "Renewable:Wind" ~ "Wind",
      fuel_type == "Storage:Battery" ~ "Battery", 
      fuel_type == "Renewable:Biomass" ~ "Biomass"
    ), 
    year = as.numeric(year)) |>
    group_by(fuel_type, year) |> 
    summarize(production_gwh = sum(production_gwh)) |> 
    mutate(
      generation_label = case_when(
        year == max(year) ~ round(production_gwh), 
        TRUE ~ NA, 
      ),
      
      fuel_label = case_when(
        year == max(year) ~ fuel_type, 
        TRUE ~ ""
      ),
      
      production_gwh = replace_na(production_gwh, 0),
      
      utility = utility)
  
  return(energy_mix)
  
}

energy_mix <- mix_cleaner(sheet_name = "Xcel", utility = "Xcel Energy") |> 
  bind_rows(mix_cleaner(sheet_name = "GRE", utility = "Great River Energy")) |> 
  bind_rows(mix_cleaner(sheet_name = "MP", utility = "Minnesota Power")) |> 
  bind_rows(mix_cleaner(sheet_name = "OTP", utility = "Otter Tail Power"))


#write_csv(energy_mix, "C:/Users/dduffy/OneDrive - State of Minnesota - MN365/Quad 2.0 Data - Documents/data/irp_energy_mix_projections.csv", na = "")
write_csv(energy_mix, "I:/Enrgy_div/SEO/CleanEnegyTechUnit/CET Projects/Data Repository/Utility IRPs/irp_energy_mix_projections.csv", na = "")



