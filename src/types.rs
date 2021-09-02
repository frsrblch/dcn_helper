use oem_types::work_order::*;

pub(crate) fn get_wo_chain<'a>(
    mut wo: &'a WorkOrderRow,
    work_orders: &'a WorkOrderData,
) -> Vec<&'a WorkOrderRow> {
    let mut wos = vec![wo];

    while let Some(parent) = wo.parent {
        wo = &work_orders[parent];
        wos.insert(0, wo);
    }

    wos
}
