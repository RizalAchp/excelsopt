fn main() {
    let mut args = std::env::args();
    let mode = args
        .nth(1)
        .expect("need more arguments, [mode : 1, 2]")
        .parse()
        .unwrap_or(0);
    let file_one = args.next().expect("provide file in arguments 2");
    let file_two = args.next().expect("provide file in arguments 3");
    let file_output = args.next().unwrap_or("output.xlsx".to_owned());
    eprintln!("running mode: {}", mode);
    if mode == 1 {
        let result = get_error(file_one, file_two, None, None);
        convert_csv_to_excel(result, file_output, "Sheet1".to_owned());
    } else if mode == 2 {
        let result = change_row(file_one, 3, 2);
        convert_csv_to_excel(result, file_output, "Sheet1".to_owned());
    } else {
        let result = change_from_one_to_another(file_one, file_two, 4, 5);
        convert_csv_to_excel(result, file_output, "Sheet1".to_owned());
    }
}

fn change_from_one_to_another(
    one: String,      // file two
    two: String,      // file two
    ito_one_a: usize, // index to compare in file one
    ito_one_b: usize, // index to compare in file one
) -> Vec<Vec<String>> {
    let mut exc_one = open_excel(&one, None);
    let exc_two = open_excel(&two, None);

    exc_one.iter_mut().for_each(|row_one| {
        for row_two in exc_two.iter() {
            if row_two.contains(&row_one[ito_one_a]) && row_two.contains(&row_one[ito_one_b]) {
                row_one[3] = row_two[1].clone();
                row_one[6] = row_two[4].clone();
                row_one[7] = row_two[3].clone();
            }
        }
    });

    exc_one
}

#[allow(unused)]
fn change_row(file: String, idx_a: usize, idx_b: usize) -> Vec<Vec<String>> {
    let mut file_excel = open_excel(&file, None);
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

    return changed;
}

#[allow(unused)]
fn get_error(
    file_1: String,
    file_2: String,
    idx_one: Option<usize>,
    idx_two: Option<usize>,
) -> Vec<Vec<String>> {
    let idx_one = idx_one.unwrap_or(5);
    let idx_two = idx_two.unwrap_or(4);
    let mut excel_one = open_excel(&file_1, None);
    let excel_two = open_excel(&file_2, None);

    let mut to_delete = vec![true; excel_one.len()];
    for (i, set_a) in excel_one.iter().enumerate().rev() {
        for set_b in excel_two.iter() {
            if !set_b.contains(&set_a[idx_one])
                && !(set_b[idx_two].to_ascii_lowercase() == set_a[idx_one].to_ascii_lowercase())
            {
                continue;
            }
            to_delete[i] = false;
        }
    }
    let mut iter_delete = to_delete.iter();
    excel_one.retain(|_| *iter_delete.next().unwrap_or_else(|| &true));
    eprintln!("size error: {}", excel_one.len());
    return excel_one;
}

#[allow(unused)]
pub(crate) fn convert_csv_to_excel<P>(
    csv_data: Vec<Vec<String>>,
    excel_path: P,
    sheets_name: String,
) where
    P: AsRef<std::path::Path>,
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
    for r in range.rows() {
        let mut row = Vec::new();
        row.reserve(r.len());
        for c in r.iter() {
            match *c {
                DataType::Empty => row.push("-".to_owned()),
                DataType::String(ref s) => row.push(s.trim().to_owned().to_uppercase()),
                DataType::Float(ref f) | DataType::DateTime(ref f) => row.push(f.to_string()),
                DataType::Int(ref i) => row.push(i.to_string()),
                DataType::Error(ref e) => row.push(e.to_string().to_uppercase()),
                DataType::Bool(ref b) => row.push(b.to_string().to_uppercase()),
            };
        }
        out.push(row);
    }
    out
}

fn open_excel(path: &str, sheets: Option<&str>) -> Vec<Vec<String>> {
    use calamine::*;
    let mut sheet = open_workbook_auto(&path).expect("cant open Workbook");
    if let Some(sh) = sheets {
        let data = sheet
            .worksheet_range(sh)
            .expect("error on opening sheets")
            .expect("Error on processing data");
        deserialize_data_excel(&data)
    } else {
        let data = sheet
            .worksheet_range_at(0)
            .expect("error on opening sheets")
            .expect("Error on processing data");
        deserialize_data_excel(&data)
    }
}
