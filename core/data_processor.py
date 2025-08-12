import re
from utils.range_optimizer import smart_range_display

def _get_summary_data(controller):
    formulas_to_summarize = [controller.view.result_tree.item(item, "values") for item in controller.view.result_tree.get_children()]
    is_filtered = len(formulas_to_summarize) != len(controller.all_formulas) if controller.all_formulas else True
    return formulas_to_summarize, is_filtered

def get_unique_external_links(formulas_to_summarize, tree_columns):
    external_path_pattern = re.compile(r"'([^']+\\[^\]]+\.(?:xlsx|xls|xlsm|xlsb)\][^']*?)'", re.IGNORECASE)
    unique_full_paths = set()
    formula_idx = tree_columns.index("formula")

    for formula_data in formulas_to_summarize:
        if len(formula_data) > formula_idx:
            formula_content = formula_data[formula_idx]
            matches = external_path_pattern.findall(str(formula_content))
            if matches:
                unique_full_paths.update(matches)

    sorted_full_paths = sorted(list(unique_full_paths))
    return sorted_full_paths
