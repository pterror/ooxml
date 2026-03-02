//! Extension traits for OMML (Office Math Markup Language) types.
//!
//! Provides convenient accessor and inspection methods for the math types
//! defined in [`crate::math`].
//!
//! # Design
//!
//! Extension traits are defined here rather than as inherent methods to keep
//! the generated/parsed types simple and allow consumers to opt in selectively.
//!
//! ECMA-376 Part 4, Section 22 defines the OMML schema.

use crate::math::{
    Fraction, FractionType, LimitLocation, MathElement, MathZone, Matrix, Nary, Radical,
};

// =============================================================================
// MathZoneExt
// =============================================================================

/// Extension methods for [`MathZone`] (m:oMath).
pub trait MathZoneExt {
    /// Get the number of top-level elements in this math zone.
    fn element_count(&self) -> usize;

    /// Check if the math zone contains no elements.
    fn is_empty(&self) -> bool;

    /// Check if any top-level element is a fraction.
    fn has_fractions(&self) -> bool;

    /// Check if any top-level element is a radical (square root or nth root).
    fn has_radicals(&self) -> bool;

    /// Check if any top-level element is a matrix.
    fn has_matrices(&self) -> bool;

    /// Check if any top-level element is a script (subscript or superscript).
    fn has_scripts(&self) -> bool;

    /// Check if any top-level element is an n-ary operator (sum, integral, etc.).
    fn has_nary(&self) -> bool;
}

impl MathZoneExt for MathZone {
    fn element_count(&self) -> usize {
        self.elements.len()
    }

    fn is_empty(&self) -> bool {
        self.elements.is_empty()
    }

    fn has_fractions(&self) -> bool {
        self.elements.iter().any(|e| e.is_fraction())
    }

    fn has_radicals(&self) -> bool {
        self.elements.iter().any(|e| e.is_radical())
    }

    fn has_matrices(&self) -> bool {
        self.elements.iter().any(|e| e.is_matrix())
    }

    fn has_scripts(&self) -> bool {
        self.elements.iter().any(|e| e.is_script())
    }

    fn has_nary(&self) -> bool {
        self.elements.iter().any(|e| e.is_nary())
    }
}

// =============================================================================
// MathElementExt
// =============================================================================

/// Extension methods for [`MathElement`].
pub trait MathElementExt {
    /// Check if this is a fraction element.
    fn is_fraction(&self) -> bool;
    /// Check if this is a radical (root) element.
    fn is_radical(&self) -> bool;
    /// Check if this is an n-ary operator (sum, integral, product, etc.).
    fn is_nary(&self) -> bool;
    /// Check if this is any kind of script (sub, sup, sub+sup, or pre-script).
    fn is_script(&self) -> bool;
    /// Check if this is a matrix.
    fn is_matrix(&self) -> bool;
    /// Check if this is a delimiter (parentheses, brackets, etc.).
    fn is_delimiter(&self) -> bool;
    /// Check if this is a text run.
    fn is_run(&self) -> bool;

    /// Try to downcast to a [`Fraction`].
    fn as_fraction(&self) -> Option<&Fraction>;
    /// Try to downcast to a [`Radical`].
    fn as_radical(&self) -> Option<&Radical>;
    /// Try to downcast to an [`Nary`].
    fn as_nary(&self) -> Option<&Nary>;
    /// Try to downcast to a [`Matrix`].
    fn as_matrix(&self) -> Option<&Matrix>;

    /// Count the number of operator elements (Nary, Fraction, Radical) in this
    /// element recursively.
    fn operator_count(&self) -> usize;
}

impl MathElementExt for MathElement {
    fn is_fraction(&self) -> bool {
        matches!(self, MathElement::Fraction(_))
    }

    fn is_radical(&self) -> bool {
        matches!(self, MathElement::Radical(_))
    }

    fn is_nary(&self) -> bool {
        matches!(self, MathElement::Nary(_))
    }

    fn is_script(&self) -> bool {
        matches!(
            self,
            MathElement::Subscript(_)
                | MathElement::Superscript(_)
                | MathElement::SubSuperscript(_)
                | MathElement::PreScript(_)
        )
    }

    fn is_matrix(&self) -> bool {
        matches!(self, MathElement::Matrix(_))
    }

    fn is_delimiter(&self) -> bool {
        matches!(self, MathElement::Delimiter(_))
    }

    fn is_run(&self) -> bool {
        matches!(self, MathElement::Run(_))
    }

    fn as_fraction(&self) -> Option<&Fraction> {
        if let MathElement::Fraction(f) = self {
            Some(f)
        } else {
            None
        }
    }

    fn as_radical(&self) -> Option<&Radical> {
        if let MathElement::Radical(r) = self {
            Some(r)
        } else {
            None
        }
    }

    fn as_nary(&self) -> Option<&Nary> {
        if let MathElement::Nary(n) = self {
            Some(n)
        } else {
            None
        }
    }

    fn as_matrix(&self) -> Option<&Matrix> {
        if let MathElement::Matrix(m) = self {
            Some(m)
        } else {
            None
        }
    }

    fn operator_count(&self) -> usize {
        count_operators_in_zone_list(std::slice::from_ref(self))
    }
}

/// Count operator elements (Nary, Fraction, Radical) recursively across a
/// slice of `MathElement`.
fn count_operators_in_zone_list(elements: &[MathElement]) -> usize {
    elements.iter().map(count_operators).sum()
}

fn count_operators(e: &MathElement) -> usize {
    match e {
        MathElement::Fraction(f) => {
            1 + count_operators_in_zone(&f.numerator) + count_operators_in_zone(&f.denominator)
        }
        MathElement::Radical(r) => {
            1 + count_operators_in_zone(&r.base) + count_operators_in_zone(&r.degree)
        }
        MathElement::Nary(n) => {
            1 + count_operators_in_zone(&n.subscript)
                + count_operators_in_zone(&n.superscript)
                + count_operators_in_zone(&n.base)
        }
        MathElement::Subscript(s) | MathElement::Superscript(s) => {
            count_operators_in_zone(&s.base) + count_operators_in_zone(&s.script)
        }
        MathElement::SubSuperscript(s) => {
            count_operators_in_zone(&s.base)
                + count_operators_in_zone(&s.subscript)
                + count_operators_in_zone(&s.superscript)
        }
        MathElement::PreScript(p) => {
            count_operators_in_zone(&p.base)
                + count_operators_in_zone(&p.subscript)
                + count_operators_in_zone(&p.superscript)
        }
        MathElement::Delimiter(d) => d.elements.iter().map(count_operators_in_zone).sum(),
        MathElement::Matrix(m) => m
            .rows
            .iter()
            .flat_map(|row| row.iter())
            .map(count_operators_in_zone)
            .sum(),
        MathElement::Function(f) => {
            count_operators_in_zone(&f.name) + count_operators_in_zone(&f.argument)
        }
        MathElement::Accent(a) => count_operators_in_zone(&a.base),
        MathElement::Bar(b) => count_operators_in_zone(&b.base),
        MathElement::Box(b) => count_operators_in_zone(&b.content),
        MathElement::BorderBox(b) => count_operators_in_zone(&b.content),
        MathElement::EquationArray(e) => e.equations.iter().map(count_operators_in_zone).sum(),
        MathElement::LowerLimit(l) | MathElement::UpperLimit(l) => {
            count_operators_in_zone(&l.base) + count_operators_in_zone(&l.limit)
        }
        MathElement::GroupChar(g) => count_operators_in_zone(&g.base),
        MathElement::Phantom(p) => count_operators_in_zone(&p.content),
        MathElement::Run(_) => 0,
    }
}

fn count_operators_in_zone(zone: &MathZone) -> usize {
    count_operators_in_zone_list(&zone.elements)
}

// =============================================================================
// FractionExt
// =============================================================================

/// Extension methods for [`Fraction`] (m:f).
pub trait FractionExt {
    /// Get the fraction type (bar, skewed, linear, no-bar).
    fn fraction_type(&self) -> Option<FractionType>;
    /// Get the numerator zone.
    fn numerator(&self) -> &MathZone;
    /// Get the denominator zone.
    fn denominator(&self) -> &MathZone;
    /// Check if this is a skewed (diagonal) fraction.
    fn is_skewed(&self) -> bool;
}

impl FractionExt for Fraction {
    fn fraction_type(&self) -> Option<FractionType> {
        self.fraction_type
    }

    fn numerator(&self) -> &MathZone {
        &self.numerator
    }

    fn denominator(&self) -> &MathZone {
        &self.denominator
    }

    fn is_skewed(&self) -> bool {
        self.fraction_type == Some(FractionType::Skewed)
    }
}

// =============================================================================
// NaryExt
// =============================================================================

/// Extension methods for [`Nary`] (m:nary — summation, integral, product, etc.).
pub trait NaryExt {
    /// Get the lower limit (subscript) zone.
    fn lower_limit(&self) -> &MathZone;
    /// Get the upper limit (superscript) zone.
    fn upper_limit(&self) -> &MathZone;
    /// Get the limit location (under/over the operator vs. sub/superscript).
    fn limit_location(&self) -> Option<LimitLocation>;
}

impl NaryExt for Nary {
    fn lower_limit(&self) -> &MathZone {
        &self.subscript
    }

    fn upper_limit(&self) -> &MathZone {
        &self.superscript
    }

    fn limit_location(&self) -> Option<LimitLocation> {
        self.limit_location
    }
}

// =============================================================================
// RadicalExt
// =============================================================================

/// Extension methods for [`Radical`] (m:rad).
pub trait RadicalExt {
    /// Get the radicand (the expression under the radical sign).
    fn radicand(&self) -> &MathZone;
    /// Get the degree zone (for nth roots; empty for square roots).
    fn degree(&self) -> &MathZone;
    /// Check if this is a square root (degree is empty / hidden).
    fn is_square_root(&self) -> bool;
}

impl RadicalExt for Radical {
    fn radicand(&self) -> &MathZone {
        &self.base
    }

    fn degree(&self) -> &MathZone {
        &self.degree
    }

    fn is_square_root(&self) -> bool {
        self.degree.elements.is_empty() || self.hide_degree
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::math::{Delimiter, Fraction, FractionType, MathRun, Nary, Radical, Script};

    fn run(text: &str) -> MathElement {
        MathElement::Run(MathRun {
            text: text.to_string(),
            properties: None,
        })
    }

    fn zone(elements: Vec<MathElement>) -> MathZone {
        MathZone { elements }
    }

    // -------------------------------------------------------------------------
    // MathZoneExt tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_math_zone_empty() {
        let z = MathZone::default();
        assert!(z.is_empty());
        assert_eq!(z.element_count(), 0);
        assert!(!z.has_fractions());
        assert!(!z.has_radicals());
        assert!(!z.has_nary());
    }

    #[test]
    fn test_math_zone_has_fraction() {
        let z = zone(vec![MathElement::Fraction(Fraction::default())]);
        assert!(!z.is_empty());
        assert_eq!(z.element_count(), 1);
        assert!(z.has_fractions());
        assert!(!z.has_radicals());
    }

    #[test]
    fn test_math_zone_has_radical() {
        let z = zone(vec![MathElement::Radical(Radical::default())]);
        assert!(z.has_radicals());
        assert!(!z.has_fractions());
    }

    #[test]
    fn test_math_zone_has_nary() {
        let z = zone(vec![MathElement::Nary(Nary::default())]);
        assert!(z.has_nary());
        assert!(!z.has_matrices());
    }

    // -------------------------------------------------------------------------
    // MathElementExt tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_element_type_checks() {
        let frac = MathElement::Fraction(Fraction::default());
        assert!(frac.is_fraction());
        assert!(!frac.is_radical());
        assert!(!frac.is_run());
        assert!(frac.as_fraction().is_some());
        assert!(frac.as_radical().is_none());

        let r = run("x");
        assert!(r.is_run());
        assert!(!r.is_fraction());
        assert!(r.as_fraction().is_none());
    }

    #[test]
    fn test_script_check() {
        let sub = MathElement::Subscript(Script::default());
        let sup = MathElement::Superscript(Script::default());
        assert!(sub.is_script());
        assert!(sup.is_script());
        assert!(!sub.is_fraction());
    }

    #[test]
    fn test_delimiter_check() {
        let d = MathElement::Delimiter(Delimiter::default());
        assert!(d.is_delimiter());
        assert!(!d.is_run());
    }

    #[test]
    fn test_operator_count_nested() {
        // fraction containing a radical in the numerator
        let frac = MathElement::Fraction(Fraction {
            numerator: zone(vec![MathElement::Radical(Radical::default())]),
            denominator: zone(vec![run("2")]),
            fraction_type: None,
        });
        // 1 (fraction) + 1 (radical) = 2
        assert_eq!(frac.operator_count(), 2);
    }

    #[test]
    fn test_operator_count_run() {
        assert_eq!(run("x").operator_count(), 0);
    }

    // -------------------------------------------------------------------------
    // FractionExt tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_fraction_ext() {
        let f = Fraction {
            numerator: zone(vec![run("1")]),
            denominator: zone(vec![run("2")]),
            fraction_type: Some(FractionType::Skewed),
        };
        assert!(f.is_skewed());
        assert_eq!(f.fraction_type(), Some(FractionType::Skewed));
        assert_eq!(f.numerator().elements.len(), 1);
        assert_eq!(f.denominator().elements.len(), 1);
    }

    #[test]
    fn test_fraction_not_skewed() {
        let f = Fraction {
            fraction_type: Some(FractionType::Bar),
            ..Default::default()
        };
        assert!(!f.is_skewed());
    }

    // -------------------------------------------------------------------------
    // RadicalExt tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_radical_square_root() {
        let r = Radical::default(); // degree is empty
        assert!(r.is_square_root());
        assert_eq!(r.degree().elements.len(), 0);
    }

    #[test]
    fn test_radical_nth_root() {
        let r = Radical {
            degree: zone(vec![run("3")]),
            base: zone(vec![run("x")]),
            hide_degree: false,
        };
        assert!(!r.is_square_root());
        assert_eq!(r.degree().elements.len(), 1);
        assert_eq!(r.radicand().elements.len(), 1);
    }

    #[test]
    fn test_radical_hide_degree_treated_as_square() {
        let r = Radical {
            degree: zone(vec![run("2")]),
            base: zone(vec![run("x")]),
            hide_degree: true,
        };
        // hide_degree=true → treated as square root regardless of degree content
        assert!(r.is_square_root());
    }

    // -------------------------------------------------------------------------
    // NaryExt tests
    // -------------------------------------------------------------------------

    #[test]
    fn test_nary_ext() {
        let n = Nary {
            operator: Some("∑".to_string()),
            subscript: zone(vec![run("i=0")]),
            superscript: zone(vec![run("n")]),
            base: zone(vec![run("x")]),
            limit_location: Some(LimitLocation::UnderOver),
            grow: false,
        };
        assert_eq!(n.lower_limit().elements.len(), 1);
        assert_eq!(n.upper_limit().elements.len(), 1);
        assert_eq!(n.limit_location(), Some(LimitLocation::UnderOver));
    }
}
