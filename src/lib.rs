#![allow(unused_must_use)]
use std::collections::HashMap;
use std::fmt::Write;

use calamine::{open_workbook_auto, Reader, DataType};

use proc_macro::{TokenStream};

#[proc_macro_attribute]
pub fn sheet(args: TokenStream, _input: TokenStream) -> TokenStream {
    println!("Parsing the provided Excel sheet");
    // i am aware of syn and i'm also aware that i didn't want to make
    // my code look worse in the name of making it "more proper"
    // (syn looks awful)
    let mut s: String = String::new();
    let mut structs: Vec<String> = Vec::new();
    for arg in args.into_iter() {
        match arg {
            // only continue if given a string as an argument
            proc_macro::TokenTree::Literal(a) => {
                let mut n = 0;
                let path = &a.to_string().replace("\"","");
                let workbook = open_workbook_auto(&path).unwrap();
                // this package puts every sheet type into different enums but the value given by the worksheets method is always the same.
                let workbooks = match workbook {
                    calamine::Sheets::Xls(mut a) => { 
                        a.worksheets()
                    },
                    calamine::Sheets::Xlsx(mut a) => {
                        a.worksheets()
                    },
                    calamine::Sheets::Xlsb(mut a) => {
                        a.worksheets()
                    },
                    calamine::Sheets::Ods(mut a) => {
                        a.worksheets()
                    },
                };
                
                for workbook in workbooks {
                    let mut field_names: HashMap<i32, String> = HashMap::new();
                    let mut first_row = true;
                    let a = workbook.1;
                    for row in a.rows().by_ref() {
                        println!("ROW: {:?}",row);
                        let len = row.len();
                        let row = [row, &[calamine::DataType::String("dummy".to_string())]].concat();
                        // the first row is alwayse used to designate the names of the fields.
                        if first_row {
                            let mut i = 0;
                            let mut first_row_first_cell = true;
                            for cell in row {
                                match cell {
                                    calamine::DataType::String(a) => {
                                        if first_row_first_cell {
                                            if a.to_lowercase() != "name" {
                                                panic!("The first cell of the first row is reserved for names, and must be titled 'name' (case insensitive)")
                                            }
                                            first_row_first_cell = false;
                                        }
                                        field_names.insert(i, a.to_lowercase().replace(" ","_"));
                                    },
                                    calamine::DataType::Empty => break,
                                    calamine::DataType::Error(a) => panic!("The provided sheet has an error: {:?}",a),
                                    _ => panic!("All of the cells in the first row must be strings.")
                                };
                                i += 1;
                            }
                            first_row = false;
                        } else {
                            let mut struct_name = String::from("");
                            let mut fields: HashMap<String, DataType> = HashMap::new();
                            // go through it once to create the structs.
                            for i in 0..row.len() {
                                let cell = row.get(i).unwrap();
                                // first cell is for the name of the struct
                                if struct_name == "" {
                                    match cell {
                                        calamine::DataType::Int(a) => {
                                            panic!("The first cell in row {} is an int ('{}') and not a string.",n,a);
                                        },
                                        calamine::DataType::Float(a) => {
                                            panic!("The first cell in row {} is a float ('{}') and not a string.",n,a);
                                        },
                                        calamine::DataType::String(a) => {
                                            struct_name = a.clone();
                                        },
                                        calamine::DataType::Bool(a) => {
                                            panic!("The first cell in row {} is a boolean ('{}') and not a string.",n,a);
                                        },
                                        calamine::DataType::Error(a) => {
                                            panic!("The first cell in row {} is an error ('{:?}') and not a string.",n,a);
                                        },
                                        calamine::DataType::DateTime(a) => {
                                            panic!("The first cell in row {} is a datetime ('{:?}') and not a string.",n,a);
                                        }
                                        calamine::DataType::Empty => {
                                            break;
                                        },
                                    };
                                // others are for the values
                                } else {
                                    let na = field_names.get_key_value(&(i as i32)).unwrap_or_else(|| {
                                        panic!("oh");
                                    });
                                    fields.insert(na.1.clone(), cell.clone());
                                    if i < (len-1) {
                                        continue;
                                    }
                                    

                                    // Create the struct first.
                                    s.write_str(format!(
                                        "pub struct {} {{",
                                        struct_name
                                    ).as_str());
                                    for (name, field) in fields.clone() {
                                        s.write_str(format!(
                                            "pub {}: ",
                                            name
                                        ).as_str());
                                        match field {
                                            calamine::DataType::Int(_) => s.write_str("i32"),
                                            calamine::DataType::Float(_) => s.write_str("f32"),
                                            calamine::DataType::String(_) => s.write_str("String"),
                                            calamine::DataType::Bool(_) => s.write_str("bool"),
                                            calamine::DataType::Error(a) => {
                                                panic!("Error at row {} cell {}: {:?}",n,i,a);
                                            }
                                            calamine::DataType::DateTime(_) => s.write_str("f64"),
                                            calamine::DataType::Empty => s.write_str("u128"),
                                        };
                                        s.write_str(",");
                                    }
                                    s.write_str("}");
                                    // Create the Default impl second.
                                    s.write_str(format!(
                                        "impl Default for {} {{
                                            fn default() -> Self {{
                                                Self {{
                                            ",
                                        struct_name
                                    ).as_str());
                                    for (name, field) in fields.clone() {
                                        match field {
                                            calamine::DataType::Int(a) => {s.write_str(format!("{}: {},",name,a).as_str());}
                                            calamine::DataType::Float(a) => {s.write_str(format!("{}: {},",name,a).as_str());}
                                            calamine::DataType::String(a) => {s.write_str(format!("{}: String::from(\"{}\"),",name,a).as_str());}
                                            calamine::DataType::Bool(a) => {s.write_str(format!("{}: {},",name,a).as_str());}
                                            calamine::DataType::Error(a) => {
                                                panic!("Error at row {} cell {}: {:?}",n,i,a);
                                            }
                                            calamine::DataType::DateTime(a) => {s.write_str(format!("{}: {},",name,a).as_str());}
                                            calamine::DataType::Empty => {s.write_str(format!("{}: 0,",name).as_str());},

                                        }
                                    }
                                    s.write_str("}}}");
                                    // finally, create the new constructor, which just uses the default one.
                                    s.write_str(format!(
                                        "impl {} {{
                                            pub fn new() -> Self {{
                                                return Default::default();
                                            }}
                                        }}
                                            ",
                                        struct_name
                                    ).as_str());
                                    structs.push(struct_name);
                                    break;
                                }
                            }
                        }
                    }
                } 
                n += 1;
            },
            _ => {
                panic!("must pass a string");
                //Diagnostic::new(Warning, "only literals are considered").emit();
            },
        }
    };
    s.write_str(format!("
        pub enum StructsFromExcel {{
    ").as_str());
    for struct_name in structs {
        s.write_str(format!("{s}({s}),",s=struct_name).as_str());
    }
    s.write_str(format!("
        }}
    ").as_str());
    println!("Finished parsing the excel sheet.");
    s.parse().expect("Generated invalid tokens")
}