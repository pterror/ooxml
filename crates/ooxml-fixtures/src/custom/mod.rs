// Custom handwritten fixture functions.
//
// Fixtures that cannot be expressed by the YAML templates are written here
// by hand and referenced from `spec/ooxml-fixture-spec.yaml` via
// `template: custom` + `custom_fn: <function_name>`.

pub mod pml;
pub mod sml;
pub mod wml;
