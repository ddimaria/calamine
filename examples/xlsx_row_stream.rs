// SPDX-License-Identifier: MIT
//
// Copyright 2016-2026, Johann Tuffe.

//! Streams XLSX rows with value, formula, and style in one worksheet pass.
//!
//! ```text
//! $ cargo run -q --example xlsx_row_stream
//! ```

use calamine::{open_workbook, Xlsx};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook: Xlsx<_> = open_workbook("tests/shared_formula_simple.xlsx")?;
    let dims = workbook.worksheet_dimensions("Sheet1")?;
    println!(
        "Sheet1 {}×{} ({} cells)",
        dims.height(),
        dims.width(),
        dims.len()
    );

    let mut reader = workbook.worksheet_cells_reader("Sheet1")?;
    while let Some(row) = reader.next_row()? {
        for cell in row.cells {
            println!(
                "r{}c{} value={:?} formula={:?} styled={}",
                cell.pos.0 + 1,
                cell.pos.1 + 1,
                cell.value,
                cell.formula,
                cell.style.is_some()
            );
        }
    }

    Ok(())
}
