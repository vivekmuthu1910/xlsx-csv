use pyo3::prelude::*;
use pyo3::wrap_pyfunction;

use calamine::{open_workbook_auto, DataType, Range, Reader};
use std::fs::File;
use std::io::{BufWriter, Write};
use std::path::PathBuf;

fn write_range<W: Write>(dest: &mut W, range: &Range<DataType>) -> std::io::Result<()> {
    let n = range.get_size().1 - 1;
    for r in range.rows() {
        for (i, c) in r.iter().enumerate() {
            match *c {
                DataType::Empty => Ok(()),
                DataType::String(ref s) => {
                    if s.contains(',') {
                        write!(dest, "\"{}\"", s)
                    } else {
                        write!(dest, "{}", s)
                    }
                }
                DataType::Float(ref f) => write!(dest, "{}", f),
                DataType::DateTime(_) => {
                    let date = c.as_datetime().unwrap();
                    write!(dest, "{}", date.format("%Y-%m-%d %H:%M:%S"))
                }
                DataType::Int(ref i) => write!(dest, "{}", i),
                DataType::Error(ref e) => write!(dest, "{:?}", e),
                DataType::Bool(ref b) => write!(dest, "{}", b),
            }?;
            if i != n {
                write!(dest, ",")?;
            }
        }
        write!(dest, "\n")?;
    }
    Ok(())
}

/// Formats the sum of two numbers as string.
#[pyfunction]
fn xlsx2csv(file: &str, sheet: &str) -> PyResult<bool> {
    // converts first argument into a csv (same name, silently overrides
    // if the file already exists

    let src = PathBuf::from(file);
    match src.extension().and_then(|s| s.to_str()) {
        Some("xlsx") | Some("xlsm") | Some("xlsb") | Some("xls") => (),
        _ => panic!("Expecting an excel file"),
    }

    let dest = src.with_extension("csv");
    let mut dest = BufWriter::new(File::create(dest).unwrap());
    let mut xl = open_workbook_auto(&src).unwrap();
    let range = xl.worksheet_range(&sheet).unwrap().unwrap();

    write_range(&mut dest, &range).unwrap();
    Ok(true)
}

/// A Python module implemented in Rust.
#[pymodule]
fn xlsx_csv(_py: Python, m: &PyModule) -> PyResult<()> {
    m.add_function(wrap_pyfunction!(xlsx2csv, m)?)?;
    Ok(())
}
