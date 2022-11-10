use std::path::{Path, PathBuf};

use anyhow::{anyhow, Result};
use clap::*;

/// Simple program manage and generate excel file
#[derive(Parser, Debug)]
#[command(author, version, about, long_about = None)]
pub struct CliArgs {
    /// file excel a to process
    file_a: PathBuf,
    /// file excel b to process
    #[clap(default_value = "file_target.xlsx")]
    file_b: PathBuf,
    /// file excel b to process
    #[clap(short, long, default_value = "output.xlsx")]
    output: PathBuf,

    #[command(subcommand)]
    action: Action,
}

#[derive(clap::Subcommand, Debug)]
enum Action {
    /// Move Row from match on comparing row from file two to one
    Move {
        /// indicies to move from
        #[clap(short, long, required = true)]
        a: Vec<usize>,
        /// indicies to move into
        #[clap(short, long, required = true)]
        b: Vec<usize>,
        /// indicies to compare
        #[clap(short, long, required = true)]
        idx: Vec<usize>,
    },
    /// Modifie Row from maatch on compareing row from single file
    Mod,
    /// check if row in file a is exist in row file b
    Check {
        /// indicies to compare
        #[clap(short, long)]
        idx: Vec<usize>,
    },
}

fn main() -> Result<()> {
    let args = CliArgs::parse();
    match args.action {
        Action::Move { a, b, idx } => {
            let result = move_row_and_compare(&args.file_a, &args.file_b, a, b, |(a, b)| {
                idx.iter().all(|i| b.contains(&a[*i]))
            })?;
            convert_csv_to_excel(result, args.output, "Sheet1".to_owned());
        }
        Action::Mod => {
            let result = change_row(args.file_a, 3, 2)?;
            convert_csv_to_excel(result, args.output, "Sheet1".to_owned());
        }
        Action::Check { idx } => {
            let result = get_error(args.file_a, args.file_b, |(a, b)| {
                idx.iter().all(|i| b.contains(&a[*i]))
            })?;
            convert_csv_to_excel(result, args.output, "Sheet1".to_owned());
        }
    }
    Ok(())
}

fn move_row_and_compare<IT, CB, S>(
    file_a: S,
    file_b: S,
    it_a: IT,
    it_b: IT,
    compare: CB,
) -> Result<Vec<Vec<String>>>
where
    IT: IntoIterator<Item = usize> + Clone,
    S: AsRef<Path>,
    CB: Fn((&Vec<String>, &Vec<String>)) -> bool,
{
    let mut exc_a = open_excel(file_a.as_ref().to_str().unwrap(), None)?;
    let exc_b = open_excel(file_b.as_ref().to_str().unwrap(), None)?;

    for row_a in exc_a.iter_mut() {
        for row_b in exc_b.iter() {
            if !compare((&row_a, row_b)) {
                continue;
            }
            it_a.clone()
                .into_iter()
                .zip(it_b.clone())
                .for_each(|(a, b)| {
                    row_a[a] = row_b[b].clone();
                });
        }
    }
    Ok(exc_a)
}

#[allow(unused)]
fn change_row<P>(file: P, idx_a: usize, idx_b: usize) -> Result<Vec<Vec<String>>>
where
    P: AsRef<Path>,
{
    let mut file_excel = open_excel(file.as_ref().to_str().unwrap(), None)?;
    let changed: Vec<Vec<String>> = file_excel
        .iter_mut()
        .map(|item| {
            let mut removed = item.swap_remove(idx_a);
            let npsn = item.get(idx_b).expect("index 2 is empty").clone();
            let idx_name = removed.find("kabjember").unwrap_or_else(|| removed.len());
            for (idx, c) in npsn.chars().enumerate() {
                removed.insert(idx_name + idx, c);
            }
            removed.insert(idx_name + npsn.len(), '.');
            item.insert(idx_a, removed);
            item.to_vec()
        })
        .collect();

    Ok(changed)
}

#[allow(unused)]
fn get_error<P, CB>(file_1: P, file_2: P, callback: CB) -> Result<Vec<Vec<String>>>
where
    P: AsRef<Path>,
    CB: Fn((&Vec<String>, &Vec<String>)) -> bool,
{
    let mut excel_one = open_excel(file_1.as_ref().to_str().unwrap(), None)?;
    let excel_two = open_excel(file_2.as_ref().to_str().unwrap(), None)?;

    let mut to_delete = vec![true; excel_one.len()];
    for (i, row_a) in excel_one.iter().enumerate().rev() {
        for row_b in excel_two.iter() {
            if callback((row_a, row_b)) {
                to_delete[i] = false;
            }
        }
    }
    let mut iter_delete = to_delete.iter();
    excel_one.retain(|_| *iter_delete.next().unwrap_or_else(|| &true));
    eprintln!("size error: {}", excel_one.len());
    Ok(excel_one)
}

#[allow(unused)]
pub(crate) fn convert_csv_to_excel<P>(
    csv_data: Vec<Vec<String>>,
    excel_path: P,
    sheets_name: String,
) where
    P: AsRef<Path>,
{
    let mut wb = simple_excel_writer::Workbook::create(
        excel_path.as_ref().to_str().unwrap_or("output.xlsx"),
    );
    let mut sheet = wb.create_sheet(&sheets_name);

    wb.write_sheet(&mut sheet, |sw| {
        for csv in csv_data {
            let mut row = simple_excel_writer::Row::new();
            for field in csv {
                row.add_cell(field);
            }
            sw.append_row(row)?;
        }
        Ok(())
    })
    .expect("cannot write sheet");

    wb.close().expect("Cannot close Workbook");
}

#[allow(unused)]
pub fn deserialize_data_excel(range: &calamine::Range<calamine::DataType>) -> Vec<Vec<String>> {
    use calamine::*;
    // let mut dest = String::new();
    let mut out = Vec::new();
    out.reserve(range.get_size().0);
    for rows in range.rows() {
        let mut row = Vec::new();
        row.reserve(rows.len());
        for c in rows.iter() {
            match *c {
                DataType::Empty => row.push("-".to_owned()),
                DataType::String(ref s) => row.push(s.trim().to_owned().to_uppercase()),
                DataType::Float(ref f) | DataType::DateTime(ref f) => row.push(f.to_string()),
                DataType::Int(ref i) => row.push(i.to_string()),
                DataType::Error(ref e) => row.push(e.to_string().to_uppercase()),
                DataType::Bool(ref b) => row.push(b.to_string()),
            };
        }
        out.push(row);
    }
    out
}

fn open_excel(path: &str, sheets: Option<&str>) -> Result<Vec<Vec<String>>> {
    use calamine::*;
    let mut sheet =
        open_workbook_auto(&path).map_err(|d| anyhow!("Error on opening workbook excel! {}", d))?;
    if let Some(sh) = sheets {
        let data = sheet
            .worksheet_range(sh)
            .expect("error on opening sheets")?;
        Ok(deserialize_data_excel(&data))
    } else {
        let data = sheet
            .worksheet_range_at(0)
            .expect("Error on processing data")?;
        Ok(deserialize_data_excel(&data))
    }
}
