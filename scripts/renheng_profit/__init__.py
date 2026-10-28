"""仁恒合并利润表确定性计算内核的稳定公开接口。"""

from .common import (
    COMPANY_FILE_KEYWORDS,
    GROUP_COMPANIES,
    GROUP_COMPANY_NAMES,
    PRE_SUMMARY_FILENAME,
    QUARTER_END_MONTHS,
    REQUIRED_INVOICE_SOURCES,
    RULE_VERSION,
    US_COMPANY,
    ZERO,
    ParseMetadata,
    SourceFile,
    active_group_company_names,
    compact_text,
    elimination_to_money,
    json_default,
    parse_profit_period,
    to_number,
)
from .consolidation import (
    build_m6_output_rows,
    build_prior_year_rows,
    read_pre_elimination_summary,
    read_result_rows,
    recalculate_operating_cost_after_elimination,
    run_m5,
    write_final_template,
    write_m6_outputs,
    write_result_rows,
)
from .income_statements import (
    aggregate_six_companies,
    parse_income_statement,
    register_sources,
    write_profit_outputs,
)
from .invoices import (
    calculate_invoice_revenue_overrides,
    calculate_yangyuan_self_sales,
    find_invoice_file,
    run_m4,
    write_invoice_revenue_override_output,
    write_yangyuan_self_sales_output,
)
from .ledger_validation import (
    build_current_subject_m7_breakdown,
    build_m7_validation_rows,
    diff_snapshot_hash,
    enrich_m7_rows_with_unbilled,
    ledger_summary,
    nonzero_m7_keys,
    read_final_revenue_values,
    read_ledger_dataset,
    read_prior_ledger_us_internal,
    read_prior_sales_ledger,
    read_us_buyer_internal,
    read_us_xinjiang_revenue,
    write_m7_validation_outputs,
    write_subject_m7_breakdown,
)
from .runtime import (
    input_fingerprint,
    validate_profit_sources,
    validate_template,
    write_run_status,
)

__all__ = [name for name in globals() if not name.startswith("_")]
