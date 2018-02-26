extern crate libc;
extern crate calamine;

use std::ffi::{CString, CStr};
use libc::{c_int, c_char, c_double, uintptr_t};

use calamine::{open_workbook_auto, Reader, DataType};

//
// Prepare Ruby bindings
//

// VALUE (pointer to a ruby object)
type Value = uintptr_t;

// Some ruby constant values
const NIL: usize   = 0x08;
const TRUE: usize  = 0x14;
const FALSE: usize = 0x00;

// Load some Ruby API functions
extern "C" {
    // Set instance variables
    fn rb_iv_set(object: Value, name: *const c_char, value: Value) -> Value;

    // Array
    fn rb_ary_new() -> Value;
    fn rb_ary_push(array: Value, elem: Value) -> Value;

    // C data to Ruby
    fn rb_int2big(num: c_int) -> Value;
    fn rb_float_new(num: c_double) -> Value;
    fn rb_utf8_str_new_cstr(str: *const c_char) -> Value;
}

//
// Utils
//

// C string from Rust string
pub fn cstr(string: &str) -> CString {
    CString::new(string).unwrap()
}

// Rust string from Ruby string
pub fn rstr(string: *const c_char) -> String {
    unsafe {
        CStr::from_ptr(string).to_string_lossy().into_owned()
    }
}

//
// Functions to use in Ruby
//

// Read the sheet
#[no_mangle]
pub unsafe extern fn read(this: Value, rb_file_name: *const c_char, rb_sheet_name: *const c_char) -> Value {
    let mut workbook = open_workbook_auto(rstr(rb_file_name)).expect("Cannot open file");

    // TODO: allow use different worksheets
    let sheet = workbook.worksheet_range(&rstr(rb_sheet_name)).unwrap().unwrap();

    let rows = rb_ary_new();

    for row in sheet.rows() {
        let new_row = rb_ary_new();

        for (_, c) in row.iter().enumerate() {
            rb_ary_push(
                new_row,
                match *c {
                    // vba error
                    DataType::Error(_) => NIL,
                    DataType::Empty => NIL,
                    DataType::Float(ref f) => rb_float_new(*f as c_double),
                    DataType::Int(ref i) => rb_int2big(*i as c_int),
                    DataType::Bool(ref b) => if *b { TRUE } else { FALSE },
                    DataType::String(ref s) => {
                        let st = s.trim();
                        if st.is_empty() {
                            NIL
                        } else {
                            rb_utf8_str_new_cstr(cstr(st).as_ptr())
                        }
                    }
                }
            );
        }

        rb_ary_push(rows, new_row);
    }


    // Set instance variables
    rb_iv_set(
        this,
        cstr("@width").as_ptr(),
        rb_int2big(sheet.width() as i32)
    );

    rb_iv_set(
        this,
        cstr("@height").as_ptr(),
        rb_int2big(sheet.height() as i32)
    );

    rb_iv_set(
        this,
        cstr("@rows").as_ptr(),
        rows
    );

    this
}
