mod compare;
use std::path::PathBuf;

use anyhow::{anyhow, Ok, Result};
use clap::*;
use compare::*;
use nom::{bytes::complete::tag, multi::separated_list0};

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
    Test,
}

fn main() -> Result<()> {
    let args = CliArgs::parse();
    match args.action {
        Action::Move { a, b, idx } => {
            let result = move_row_and_compare(&args.file_a, &args.file_b, a, b, |(a, b)| {
                idx.iter().all(|i| b.contains(&a[*i]))
            })?;
            Ok(convert_csv_to_excel(
                result,
                args.output,
                "Sheet1".to_owned(),
            ))
        }
        Action::Mod => {
            let result = change_row(args.file_a, 3, 2)?;
            Ok(convert_csv_to_excel(
                result,
                args.output,
                "Sheet1".to_owned(),
            ))
        }
        Action::Check { idx } => {
            let result = get_error(args.file_a, args.file_b, |(a, b)| {
                idx.iter().all(|i| b.contains(&a[*i]))
            })?;
            Ok(convert_csv_to_excel(
                result,
                args.output,
                "Sheet1".to_owned(),
            ))
        }
        Action::Test => testing(args.file_a),
    }
}

pub fn testing(file: PathBuf) -> Result<()> {
    if !file.exists() {
        return Err(anyhow!("File {} is not exists!", file.display()));
    }
    let file = std::fs::read(file)?;
    let parsed = separated_list0(tag(b"\n\n"), separated_list0(tag(b"\t"), nom::bytes::complete::take(0..)))(&file);

    Ok(())
}
