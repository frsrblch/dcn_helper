#![feature(option_result_contains, path_try_exists)]

use calamine::{open_workbook, Reader, Xlsx};
use oem_types::work_order::*;
use std::collections::{BTreeSet, HashSet};
use std::error::Error;
use std::str::FromStr;
use types::*;

pub mod types;

fn main() {
    let output = load_work_orders()
        .and_then(|work_orders| get_target_fgs().map(|fgs| Output { fgs, work_orders }))
        .map(|output| output.to_string())
        .unwrap_or_else(|e| e.to_string());

    std::fs::write("output.txt", output).unwrap();

    std::process::Command::new("notepad")
        .arg("output.txt")
        .spawn()
        .expect("error starting notepad");
}

fn load_work_orders() -> Result<WorkOrderData, Box<dyn Error>> {
    const PATH: &'static str = "WIP.xlsx";

    if !std::fs::try_exists(PATH)? {
        return Err(format!("File not found: {}", PATH).into());
    }

    let mut workbook: Xlsx<_> =
        open_workbook(PATH).map_err(|_| format!("Cannot read file as .xlsx"))?;

    let sheet = workbook
        .worksheet_range_at(0)
        .ok_or_else(|| format!("No sheets in workbook"))??;

    let data = WorkOrderRow::from_sheet(&sheet)?;

    if data.is_empty() {
        return Err("No work order data found".to_string().into());
    }

    if let Ok(modified) = std::fs::metadata(PATH)
        .and_then(|metadata| metadata.modified().or_else(|_| metadata.created()))
    {
        if let Ok(duration) = std::time::SystemTime::now().duration_since(modified) {
            let minutes = duration.as_secs() / 60;
            if minutes > 60 {
                return Err(format!(
                    "WOT data updated {} minutes ago, refresh and retry.",
                    minutes
                )
                .into());
            }
        }
    }

    Ok(data.into_iter().collect())
}

fn get_target_fgs() -> Result<HashSet<FinishedGood>, Box<dyn Error>> {
    const PATH: &'static str = "fgs.txt";

    let mut fgs = HashSet::new();

    for fg in std::fs::read_to_string(PATH)
        .map_err(|_| format!("File not found: {}", PATH))?
        .replace(",", " ")
        .trim()
        .to_string()
        .split_whitespace()
        .filter_map(|s| FinishedGood::from_str(s).ok())
    {
        fgs.insert(fg);
    }

    if fgs.is_empty() {
        return Err(format!("No finished goods found in {}", PATH).into());
    }

    Ok(fgs)
}

fn get_fg_chains(
    fgs: &HashSet<FinishedGood>,
    work_orders: &WorkOrderData,
) -> BTreeSet<Vec<FinishedGood>> {
    let mut fgs: BTreeSet<Vec<FinishedGood>> =
        fgs.into_iter().cloned().map(|fg| vec![fg]).collect();
    let mut modified = true;

    while modified {
        modified = false;

        fgs = fgs
            .into_iter()
            .flat_map(|chain| {
                let parents = work_orders.get_parent_fgs(chain.first().unwrap());

                if parents.is_empty() {
                    vec![chain]
                } else {
                    modified = true;

                    parents
                        .into_iter()
                        .map(|parent| {
                            let mut new = vec![parent];
                            new.extend(chain.clone());
                            new
                        })
                        .collect()
                }
            })
            .collect();
    }

    fgs
}

struct Output {
    fgs: HashSet<FinishedGood>,
    work_orders: WorkOrderData,
}

impl std::fmt::Display for Output {
    fn fmt(&self, f: &mut std::fmt::Formatter) -> std::fmt::Result {
        for chain in get_fg_chains(&self.fgs, &self.work_orders) {
            for fg in chain.iter() {
                write!(f, "{}\tRouting Step\t", fg)?;
            }

            writeln!(f, "Work Type\tSales Level")?;

            for child_wo in self.work_orders.get_sorted_by_routing_step(&chain) {
                for wo in get_wo_chain(child_wo, &self.work_orders) {
                    let routing_step = if wo.fg.is_step() && !wo.fg.is_step_number(1) {
                        // get Step 1 work order routing

                        let parent = wo
                            .parent
                            .unwrap_or_else(|| panic!("Step WO missing parent: {}", wo.work_order));

                        self.work_orders
                            .iter()
                            .filter(|wo| wo.parent.contains(&parent) && wo.fg.is_step_number(1))
                            .next()
                            .unwrap_or_else(|| {
                                panic!("Step 1 work order not found under WO {}", parent)
                            })
                            .routing_step
                    } else {
                        wo.routing_step
                    };

                    write!(f, "{}\t{}\t", wo.work_order, routing_step)?;
                }

                writeln!(f, "{}\t{}", child_wo.work_type, child_wo.sales_level)?;
            }

            writeln!(f)?;
        }

        Ok(())
    }
}
