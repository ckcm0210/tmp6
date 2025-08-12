
from .link_analyzer import is_external_link_regex_match

def classify_formula_type(formula):
    formula_str = str(formula)
    
    # Use the more comprehensive external link detection
    if is_external_link_regex_match(formula_str):
        return 'external link'
    
    if '!' in formula_str:
        return 'local link'
    
    if formula_str.startswith('='):
        return 'formula'
    
    return 'formula'