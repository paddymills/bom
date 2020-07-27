use calamine::{open_workbook, DataType, Reader, Xlsx};
use std::path::Path;

pub fn test() {
    let path = Path::new(r"C:\Users\PMiller1\git\bom\src\Book1.xlsx");

    let mut wb: Xlsx<_> = open_workbook(path).unwrap();
    println!("{:#?}", wb.sheet_names());

    let rng = wb.worksheet_range("Sheet1").unwrap().unwrap();
    let mut row_sep = "==============================================\n";
    for r in rng.rows() {
        let mut sep = "";

        for c in r.iter() {
            print!(" {}", sep);
            match c {
                DataType::Empty => print!("        "),
                DataType::Bool(ref b) => print!("{:^8}", b),
                DataType::Int(ref i) => print!("{:^8}", i),
                DataType::Float(ref f) => print!("{:^8}", f),
                DataType::String(ref s) => print!("{:^8}", s),
                DataType::Error(ref e) => print!("{:^8?}", e),
            }

            sep = "| ";
        }

        print!("\n{}", row_sep);
        row_sep = "";
    }
}
