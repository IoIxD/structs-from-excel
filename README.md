# structs-from-excel

This crate adds a procedural macro that will generate structs based on a given Excel spreadsheet. It was made for a project where I knew that one of the people helping was not going to want to work with JSON or XML or anything reasonable. I'm sorry.

Invocation of the macro is as follows:

```
use structs_from_excel;

#[sheet("resources/objects.xls")]
pub struct Object; // This struct will get ignored and replaced by whatever structs you define. Anything can go here, pretty much.
```

Each sheet is then read and parsed, keeping the following rules in mind:

- The first cell of the first row should be called name.
- The first cell of every row beyond the first should be used as the name of each struct.
- Subsequent cells in the first row are used for field names.

The following two sheets:

![image](https://user-images.githubusercontent.com/30945097/224210459-093eb187-5847-4a71-8d57-3b093bdff703.png)

![image](https://user-images.githubusercontent.com/30945097/224211840-0a173b35-63ae-4cea-bc63-bf021640e25e.png)

become:

```
    enum StructsFromExcel {
        Slime(Slime),
        OtherSlime(OtherSlime),
        PlayerOne(PlayerOne),
        PlayerTwo(PlayerTwo),
        PlayerThree(PlayerThree),
        PlayerFour(PlayerFour),
    }
    pub struct Slime {
        pub hp: i32,
    }
    impl Default for Slime {
        fn default() -> Self {
            Self { hp: 30 }
        }
    }
    impl Slime {
        pub fn new() -> Self {
            return Default::default();
        }
    }
    pub struct OtherSlime {
        pub hp: i32,
    }
    impl Default for OtherSlime {
        fn default() -> Self {
            Self { hp: 60 }
        }
    }
    impl OtherSlime {
        pub fn new() -> Self {
            return Default::default();
        }
    }
    pub struct PlayerOne {
        pub what: f32,
        pub wa: String,
        pub hp: i32,
    }
    impl Default for PlayerOne {
        fn default() -> Self {
            Self {
                what: 15.5,
                wa: String::from("lorem"),
                hp: 20,
            }
        }
    }
    impl PlayerOne {
        pub fn new() -> Self {
            return Default::default();
        }
    }
    pub struct PlayerTwo {
        pub hp: i32,
        pub wa: String,
        pub what: f32,
    }
    impl Default for PlayerTwo {
        fn default() -> Self {
            Self {
                hp: 20,
                wa: String::from("ipsum"),
                what: 15.5,
            }
        }
    }
    impl PlayerTwo {
        pub fn new() -> Self {
            return Default::default();
        }
    }
    pub struct PlayerThree {
        pub what: f32,
        pub wa: String,
        pub hp: i32,
    }
    impl Default for PlayerThree {
        fn default() -> Self {
            Self {
                what: 15.5,
                wa: String::from("dot"),
                hp: 20,
            }
        }
    }
    impl PlayerThree {
        pub fn new() -> Self {
            return Default::default();
        }
    }
    pub struct PlayerFour {
        pub what: f32,
        pub wa: String,
        pub hp: i32,
    }
    impl Default for PlayerFour {
        fn default() -> Self {
            Self {
                what: 15.5,
                wa: String::from("ament"),
                hp: 20,
            }
        }
    }
    impl PlayerFour {
        pub fn new() -> Self {
            return Default::default();
        }
    }
```
