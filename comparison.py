"""Deal comparison analysis utilities for VARO rebilling dashboard."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timezone
from io import BytesIO
from typing import Any, Dict, Iterable, List, Optional, Tuple
from uuid import uuid4

import numpy as np
import pandas as pd
from fpdf import FPDF
from openpyxl.utils import column_index_from_string


def _normalize_column_name(column: Any) -> str:
    """Normalize a column header into snake case for easier matching."""
    text = str(column or "").strip().lower()
    normalized: List[str] = []
    for char in text:
        if char.isalnum():
            normalized.append(char)
        else:
            normalized.append("_")
    value = "".join(normalized).strip("_")
    return value or "column"


def _pretty_label(column: str) -> str:
    """Return a nicely formatted column label."""
    if not column:
        return "Column"
    cleaned = column.replace("_", " ").strip()
    return cleaned.title()


def _format_currency(value: float) -> str:
    return f"${value:,.2f}"


def _format_percentage(value: Optional[float]) -> str:
    if value is None or np.isnan(value):
        return "-"
    return f"{value:.2f}%"


@dataclass
class CostColumn:
    key: str
    label: str
    formatted_column: Optional[str] = None
    comparison_column: Optional[str] = None


class DealComparisonAnalyzer:
    """Perform advanced comparison between formatted and reference workbooks."""

    COST_KEYWORDS = (
        "cost",
        "insurance",
        "inspection",
        "superintendent",
        "charge",
        "fee",
        "logistics",
    )

    def __init__(self) -> None:
        self._cache: Dict[str, Dict[str, Any]] = {}

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def analyze(
        self,
        formatted_bytes: bytes,
        comparison_bytes: bytes,
        *,
        formatted_sheet: Optional[str] = None,
        comparison_sheet: Optional[str] = None,
        formatted_quantity_letter: str = "L",
        comparison_quantity_column: Optional[str] = None,
    ) -> Dict[str, Any]:
        """Run the complete comparison workflow and return JSON payload."""

        df_formatted_raw = self._load_dataframe(formatted_bytes, formatted_sheet)
        df_comparison_raw = self._load_dataframe(comparison_bytes, comparison_sheet)

        df_formatted, formatted_meta = self._standardize_dataframe(df_formatted_raw)
        df_comparison, comparison_meta = self._standardize_dataframe(df_comparison_raw)

        deal_col_formatted = self._identify_deal_column(df_formatted)
        deal_col_comparison = self._identify_deal_column(df_comparison)

        formatted_quantity_column = self._column_by_letter(
            df_formatted, formatted_quantity_letter
        )

        comparison_quantity_column = self._identify_quantity_column(
            df_comparison,
            preferred=comparison_quantity_column,
            fallback_letter=formatted_quantity_letter,
        )

        formatted_costs = self._extract_cost_columns(
            df_formatted, formatted_quantity_column
        )
        comparison_costs = self._extract_cost_columns(
            df_comparison, comparison_quantity_column
        )

        cost_info = self._build_cost_registry(
            formatted_costs,
            comparison_costs,
            formatted_meta,
            comparison_meta,
        )

        formatted_numeric = self._prepare_numeric_dataset(
            df_formatted,
            deal_col_formatted,
            formatted_quantity_column,
            formatted_costs,
        )
        comparison_numeric = self._prepare_numeric_dataset(
            df_comparison,
            deal_col_comparison,
            comparison_quantity_column,
            comparison_costs,
        )

        merged = self._merge_datasets(formatted_numeric, comparison_numeric)

        results = self._build_analysis_payload(merged, cost_info)

        token = uuid4().hex
        self._cache[token] = {
            "timestamp": datetime.now(timezone.utc),
            "dataframe": merged,
            "cost_info": cost_info,
            "payload": results,
        }
        results["token"] = token
        return results

    def get_cached_payload(self, token: str) -> Dict[str, Any]:
        if token not in self._cache:
            raise KeyError("Analysis token not found")
        return self._cache[token]

    def generate_excel(self, token: str) -> bytes:
        cached = self.get_cached_payload(token)
        merged: pd.DataFrame = cached["dataframe"]
        cost_info: Dict[str, CostColumn] = cached["cost_info"]
        payload: Dict[str, Any] = cached["payload"]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            overview_df = pd.DataFrame(
                [
                    {
                        "Metric": "Deals with variance",
                        "Value": payload["overview"]["total_deals"],
                    },
                    {
                        "Metric": "Total USD discrepancy",
                        "Value": payload["overview"]["total_difference"],
                    },
                    {
                        "Metric": "Average variance %",
                        "Value": payload["overview"]["average_variance"],
                    },
                    {
                        "Metric": "Unregistered cost types",
                        "Value": payload["overview"]["unregistered_cost_types"],
                    },
                ]
            )
            overview_df.to_excel(writer, sheet_name="Overview", index=False)

            deals_df = pd.DataFrame(payload["deals"])
            deals_df.to_excel(writer, sheet_name="Deal Differences", index=False)

            breakdown_df = pd.DataFrame(payload["cost_breakdown"])
            breakdown_df.to_excel(writer, sheet_name="Cost Breakdown", index=False)

            unregistered_df = pd.DataFrame(payload["unregistered_costs"])
            unregistered_df.to_excel(
                writer, sheet_name="Unregistered Costs", index=False
            )

            heatmap = payload["heatmap"]
            matrix_df = pd.DataFrame(
                heatmap["status_matrix"],
                columns=heatmap["cost_types"],
                index=heatmap["deal_ids"],
            )
            matrix_df.index.name = "Deal"
            matrix_df.reset_index().to_excel(
                writer, sheet_name="Heatmap", index=False
            )

            merged.to_excel(writer, sheet_name="Raw Data", index=False)

        return output.getvalue()

    def generate_csv(self, token: str) -> bytes:
        cached = self.get_cached_payload(token)
        payload: Dict[str, Any] = cached["payload"]
        deals_df = pd.DataFrame(payload["deals"])
        return deals_df.to_csv(index=False).encode("utf-8")

    def generate_pdf(self, token: str) -> bytes:
        cached = self.get_cached_payload(token)
        payload: Dict[str, Any] = cached["payload"]
        summary = payload["summary_report"]

        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Helvetica", "B", 16)
        pdf.cell(0, 10, "Deal Comparison Summary", ln=True)

        pdf.set_font("Helvetica", size=12)
        pdf.multi_cell(0, 8, summary["headline"])
        pdf.ln(2)
        pdf.multi_cell(0, 8, summary["top_contributors"])
        pdf.ln(2)
        pdf.multi_cell(0, 8, summary["unregistered_costs"])
        pdf.ln(4)

        pdf.set_font("Helvetica", "B", 12)
        pdf.cell(0, 8, "Recommended Actions", ln=True)
        pdf.set_font("Helvetica", size=12)
        for item in summary["recommended_actions"]:
            pdf.multi_cell(0, 6, f"• {item}")

        return pdf.output(dest="S").encode("latin-1")

    # ------------------------------------------------------------------
    # Helpers
    # ------------------------------------------------------------------
    def _load_dataframe(
        self, file_bytes: bytes, sheet_name: Optional[str]
    ) -> pd.DataFrame:
        if not file_bytes:
            raise ValueError("Uploaded file is empty")
        excel_stream = BytesIO(file_bytes)
        df = pd.read_excel(
            excel_stream,
            sheet_name=sheet_name if sheet_name else 0,
            dtype=object,
        )
        df = df.dropna(how="all")
        return df

    def _standardize_dataframe(
        self, df: pd.DataFrame
    ) -> Tuple[pd.DataFrame, Dict[str, Dict[str, str]]]:
        columns: List[str] = []
        original: Dict[str, str] = {}
        display: Dict[str, str] = {}
        used: Dict[str, int] = {}

        for column in df.columns:
            base = _normalize_column_name(column)
            if base in used:
                used[base] += 1
                base = f"{base}_{used[base]}"
            else:
                used[base] = 1

            columns.append(base)
            original[base] = str(column)
            display[base] = _pretty_label(str(column))

        standardized = df.copy()
        standardized.columns = columns
        return standardized, {"original": original, "display": display}

    def _identify_deal_column(self, df: pd.DataFrame) -> str:
        priority = (
            "deal_id",
            "varo_deal",
            "deal",
            "deal_number",
            "deal_no",
            "dealname",
            "vsa_deal",
        )
        for name in priority:
            if name in df.columns:
                return name

        for column in df.columns:
            if "deal" in column:
                return column

        raise ValueError("Unable to identify deal identifier column")

    def _column_by_letter(self, df: pd.DataFrame, letter: str) -> str:
        index = column_index_from_string(letter.upper()) - 1
        if index >= len(df.columns):
            raise ValueError(
                f"Column {letter.upper()} not found in formatted worksheet"
            )
        return df.columns[index]

    def _identify_quantity_column(
        self,
        df: pd.DataFrame,
        *,
        preferred: Optional[str],
        fallback_letter: str,
    ) -> str:
        if preferred and preferred in df.columns:
            return preferred

        hints = (
            "total_usd",
            "total",
            "usd_total",
            "qty_usd",
            "amount",
            "usd",
        )
        for hint in hints:
            for column in df.columns:
                if hint in column:
                    return column

        try:
            return self._column_by_letter(df, fallback_letter)
        except ValueError:
            pass

        numeric_columns = [
            column
            for column in df.columns
            if pd.to_numeric(df[column], errors="coerce").notna().any()
        ]
        if numeric_columns:
            return numeric_columns[0]

        raise ValueError("Unable to locate quantity column in comparison sheet")

    def _extract_cost_columns(
        self, df: pd.DataFrame, quantity_column: str
    ) -> List[str]:
        cost_columns: List[str] = []
        for column in df.columns:
            if column == quantity_column:
                continue
            if any(keyword in column for keyword in self.COST_KEYWORDS):
                cost_columns.append(column)
        return cost_columns

    def _build_cost_registry(
        self,
        formatted_costs: Iterable[str],
        comparison_costs: Iterable[str],
        formatted_meta: Dict[str, Dict[str, str]],
        comparison_meta: Dict[str, Dict[str, str]],
    ) -> Dict[str, CostColumn]:
        registry: Dict[str, CostColumn] = {}

        def ensure_entry(column: str, meta: Dict[str, Dict[str, str]]) -> CostColumn:
            key = _normalize_column_name(column)
            label = meta["display"].get(column, _pretty_label(column))
            entry = registry.get(key)
            if not entry:
                entry = CostColumn(key=key, label=label)
                registry[key] = entry
            elif not entry.label or entry.label.lower() == entry.key:
                entry.label = label
            return entry

        for column in formatted_costs:
            entry = ensure_entry(column, formatted_meta)
            entry.formatted_column = column

        for column in comparison_costs:
            entry = ensure_entry(column, comparison_meta)
            entry.comparison_column = column

        return registry

    def _prepare_numeric_dataset(
        self,
        df: pd.DataFrame,
        deal_column: str,
        quantity_column: str,
        cost_columns: Iterable[str],
    ) -> pd.DataFrame:
        numeric_columns = [quantity_column, *cost_columns]
        subset = df[[deal_column, *numeric_columns]].copy()

        for column in numeric_columns:
            subset[column] = pd.to_numeric(subset[column], errors="coerce").fillna(0)

        grouped = subset.groupby(deal_column, dropna=False).sum().reset_index()
        grouped = grouped.rename(columns={deal_column: "deal_id"})
        grouped = grouped.rename(columns={quantity_column: "total_quantity"})

        renamed = {}
        for column in cost_columns:
            key = _normalize_column_name(column)
            renamed[column] = f"cost_{key}"

        return grouped.rename(columns=renamed)

    def _merge_datasets(
        self, formatted: pd.DataFrame, comparison: pd.DataFrame
    ) -> pd.DataFrame:
        merged = formatted.merge(
            comparison,
            on="deal_id",
            how="outer",
            suffixes=("_formatted", "_comparison"),
        )
        merged = merged.fillna(0)

        merged["quantity_difference"] = (
            merged["total_quantity_formatted"] - merged["total_quantity_comparison"]
        )
        merged["percentage_variance"] = np.where(
            merged["total_quantity_comparison"] == 0,
            np.nan,
            (merged["quantity_difference"]
             / merged["total_quantity_comparison"]) * 100,
        )
        merged["abs_difference"] = merged["quantity_difference"].abs()
        merged.sort_values("abs_difference", ascending=False, inplace=True)

        if len(merged):
            merged["rank"] = (
                merged["abs_difference"].rank(method="first", ascending=False).astype(int)
            )
        else:
            merged["rank"] = []

        return merged

    def _build_analysis_payload(
        self, merged: pd.DataFrame, cost_info: Dict[str, CostColumn]
    ) -> Dict[str, Any]:
        filtered = merged[merged["quantity_difference"] > 0].copy()

        deals_payload: List[Dict[str, Any]] = []
        unregistered_cost_tracker: Dict[str, Dict[str, Any]] = {}


        cost_columns = sorted(cost_info.values(), key=lambda c: c.label.lower())
        for _, row in filtered.iterrows():
            deal_id = row["deal_id"]
            cost_details: List[Dict[str, Any]] = []
            unregistered_for_deal: List[str] = []
            partial_for_deal: List[str] = []


            for cost in cost_columns:
                formatted_value = (
                    row.get(f"cost_{cost.key}_formatted", 0.0) if cost.formatted_column else 0.0
                )
                comparison_value = (
                    row.get(f"cost_{cost.key}_comparison", 0.0)
                    if cost.comparison_column
                    else 0.0
                )

                difference = formatted_value - comparison_value
                percentage = (
                    (difference / comparison_value * 100)
                    if comparison_value
                    else (np.nan if formatted_value == 0 else np.nan)
                )

                status = "Missing"
                if formatted_value and not comparison_value:
                    status = "Unregistered"
                    unregistered_for_deal.append(cost.label)
                    tracker = unregistered_cost_tracker.setdefault(
                        cost.label,
                        {"total_difference": 0.0, "deals": set()}
                    )
                    tracker["total_difference"] += float(difference)
                    tracker["deals"].add(deal_id)

                elif formatted_value and comparison_value:
                    status = "Registered"
                    variance = abs(difference) / comparison_value * 100 if comparison_value else 0
                    if variance >= 5:
                        status = "Partial"
                        partial_for_deal.append(cost.label)
                elif not formatted_value and comparison_value:
                    status = "Partial"
                    partial_for_deal.append(cost.label)

                cost_details.append(
                    {
                        "cost_type": cost.label,
                        "formatted": round(float(formatted_value), 2),
                        "comparison": round(float(comparison_value), 2),
                        "difference": round(float(difference), 2),
                        "percentage": None if np.isnan(percentage) else round(float(percentage), 2),
                        "status": status,
                    }
                )

            overall_status = "Registered"
            if unregistered_for_deal:
                overall_status = "Unregistered"

            elif partial_for_deal:
                overall_status = "Partial"

            deals_payload.append(
                {
                    "deal_id": deal_id,
                    "formatted_quantity": round(float(row["total_quantity_formatted"]), 2),
                    "comparison_quantity": round(float(row["total_quantity_comparison"]), 2),
                    "difference": round(float(row["quantity_difference"]), 2),
                    "percentage_variance": None
                    if np.isnan(row["percentage_variance"])
                    else round(float(row["percentage_variance"]), 2),
                    "rank": int(row.get("rank", 0)) if len(filtered) else 0,
                    "cost_registry_status": overall_status,
                    "costs": cost_details,

                }
            )

        total_difference = float(filtered["quantity_difference"].sum())
        average_variance = float(
            filtered["percentage_variance"].mean(skipna=True)
            if len(filtered)
            else 0.0
        )

        top_deals = sorted(deals_payload, key=lambda d: d["difference"], reverse=True)[:20]

        cost_breakdown: List[Dict[str, Any]] = []
        for cost in cost_columns:
            formatted_series = filtered.get(f"cost_{cost.key}_formatted", pd.Series(dtype=float))
            comparison_series = filtered.get(f"cost_{cost.key}_comparison", pd.Series(dtype=float))

            formatted_total = float(formatted_series.sum()) if len(formatted_series) else 0.0
            comparison_total = float(comparison_series.sum()) if len(comparison_series) else 0.0
            difference = formatted_total - comparison_total
            percentage = (
                (difference / comparison_total * 100)
                if comparison_total
                else (np.nan if formatted_total == 0 else 100.0)
            )

            status = "Registered"
            if formatted_total and not comparison_total:
                status = "Unregistered"
            elif abs(difference) > 0 and comparison_total:
                variance = abs(difference) / (comparison_total or 1) * 100
                if variance >= 5:
                    status = "Partial"

            cost_breakdown.append(
                {
                    "cost_type": cost.label,
                    "formatted_total": round(formatted_total, 2),
                    "comparison_total": round(comparison_total, 2),
                    "difference": round(difference, 2),
                    "percentage": None if np.isnan(percentage) else round(float(percentage), 2),
                    "status": status,
                }
            )

        unregistered_costs: List[Dict[str, Any]] = []
        for cost_label, data in unregistered_cost_tracker.items():
            unregistered_costs.append(
                {
                    "cost_type": cost_label,
                    "impact": round(float(data["total_difference"]), 2),

                }
            )

        heatmap = self._build_heatmap(filtered, cost_columns)
        anomalies, cost_anomalies = self._detect_anomalies(filtered, cost_columns)
        patterns = self._detect_patterns(deals_payload)

        summary_report = self._build_summary_report(
            deals_payload,
            total_difference,
            unregistered_costs,
            patterns,
        )

        return {
            "overview": {
                "total_deals": len(deals_payload),
                "total_difference": round(total_difference, 2),
                "average_variance": round(average_variance, 2),
                "unregistered_cost_types": len(unregistered_costs),

                "anomaly_count": len(anomalies),
            },
            "deals": deals_payload,
            "top_deals": top_deals,
            "cost_breakdown": cost_breakdown,
            "unregistered_costs": unregistered_costs,
            "heatmap": heatmap,
            "anomalies": anomalies,
            "cost_anomalies": cost_anomalies,
            "patterns": patterns,
            "summary_report": summary_report,
        }

    def _build_heatmap(
        self, df: pd.DataFrame, cost_columns: List[CostColumn]
    ) -> Dict[str, Any]:
        deal_ids = df["deal_id"].tolist()
        cost_labels = [cost.label for cost in cost_columns]
        status_matrix: List[List[str]] = []
        z_values: List[List[float]] = []
        hover: List[List[str]] = []

        for _, row in df.iterrows():
            status_row: List[str] = []
            value_row: List[float] = []
            hover_row: List[str] = []
            for cost in cost_columns:
                formatted_value = (
                    row.get(f"cost_{cost.key}_formatted", 0.0) if cost.formatted_column else 0.0
                )
                comparison_value = (
                    row.get(f"cost_{cost.key}_comparison", 0.0)
                    if cost.comparison_column
                    else 0.0
                )

                if not cost.formatted_column and not cost.comparison_column:
                    status_row.append("Missing")
                    value_row.append(np.nan)
                    hover_row.append("No data")
                    continue

                if formatted_value == 0 and comparison_value == 0:
                    status = "Within"
                    value = 0
                elif formatted_value and not comparison_value:
                    status = "Unregistered"
                    value = -10
                elif not formatted_value and comparison_value:
                    status = "5-20% Lower"
                    value = -1
                else:
                    difference = formatted_value - comparison_value
                    percentage = (
                        difference / comparison_value * 100 if comparison_value else 0
                    )
                    if percentage >= 20:
                        status = ">20% Higher"
                        value = 2
                    elif percentage >= 5:
                        status = "5-20% Higher"
                        value = 1
                    elif percentage <= -20:
                        status = ">20% Lower"
                        value = -2
                    elif percentage <= -5:
                        status = "5-20% Lower"
                        value = -1
                    else:
                        status = "Within"
                        value = 0

                status_row.append(status)
                value_row.append(value)
                diff_value = float(formatted_value) - float(comparison_value)
                direction = "↑" if diff_value > 0 else ("↓" if diff_value < 0 else "–")
                hover_row.append(
                    f"Formatted: {_format_currency(float(formatted_value))}<br>"
                    f"Comparison: {_format_currency(float(comparison_value))}<br>"
                    f"Δ {direction} {_format_currency(diff_value)}"
                )

            status_matrix.append(status_row)
            z_values.append(value_row)
            hover.append(hover_row)

        return {
            "deal_ids": deal_ids,
            "cost_types": cost_labels,
            "matrix": z_values,
            "status_matrix": status_matrix,
            "hover": hover,
        }

    def _detect_anomalies(
        self, df: pd.DataFrame, cost_columns: List[CostColumn]
    ) -> Tuple[List[Dict[str, Any]], List[Dict[str, Any]]]:
        anomalies: List[Dict[str, Any]] = []
        cost_anomalies: List[Dict[str, Any]] = []

        if len(df) >= 2:
            diffs = df["quantity_difference"]
            mean = diffs.mean()
            std = diffs.std(ddof=0)
            if std > 0:
                threshold = mean + 2 * std
                for _, row in df.iterrows():
                    if row["quantity_difference"] > threshold:
                        anomalies.append(
                            {
                                "deal_id": row["deal_id"],
                                "difference": round(float(row["quantity_difference"]), 2),
                                "formatted_quantity": round(
                                    float(row["total_quantity_formatted"]), 2
                                ),
                                "comparison_quantity": round(
                                    float(row["total_quantity_comparison"]), 2
                                ),
                            }
                        )

        for cost in cost_columns:
            formatted_col = f"cost_{cost.key}_formatted"
            comparison_col = f"cost_{cost.key}_comparison"
            if formatted_col not in df.columns and comparison_col not in df.columns:
                continue
            differences = (
                df.get(formatted_col, pd.Series(dtype=float))
                - df.get(comparison_col, pd.Series(dtype=float))
            )
            if len(differences) < 2:
                continue
            mean = differences.mean()
            std = differences.std(ddof=0)
            if std <= 0:
                continue
            threshold = mean + 2 * std
            for _, row in df.iterrows():
                diff_value = (
                    row.get(formatted_col, 0.0) - row.get(comparison_col, 0.0)
                )
                if diff_value > threshold:
                    cost_anomalies.append(
                        {
                            "deal_id": row["deal_id"],
                            "cost_type": cost.label,
                            "difference": round(float(diff_value), 2),
                        }
                    )

        return anomalies, cost_anomalies

    def _detect_patterns(self, deals: List[Dict[str, Any]]) -> Dict[str, Any]:
        pattern_map: Dict[Tuple[str, ...], List[str]] = {}
        status_counts: Dict[str, int] = {"Registered": 0, "Partial": 0, "Unregistered": 0}

        for deal in deals:
            status = deal["cost_registry_status"]
            status_counts[status] = status_counts.get(status, 0) + 1
            unregistered = sorted(
                cost["cost_type"]
                for cost in deal["costs"]
                if cost["status"] == "Unregistered"
            )
            if unregistered:
                key = tuple(unregistered)
                pattern_map.setdefault(key, []).append(deal["deal_id"])

        repeating_patterns = []
        for cost_tuple, deals_for_pattern in pattern_map.items():
            if len(deals_for_pattern) >= 2:
                repeating_patterns.append(
                    {
                        "cost_types": list(cost_tuple),
                        "deals": deals_for_pattern,
                    }
                )

        return {
            "status_counts": status_counts,
            "repeating_patterns": repeating_patterns,
        }

    def _build_summary_report(
        self,
        deals: List[Dict[str, Any]],
        total_difference: float,
        unregistered_costs: List[Dict[str, Any]],
        patterns: Dict[str, Any],
    ) -> Dict[str, Any]:
        if deals:
            headline = (
                f"{len(deals)} deals show higher quantities in processed sheet, "
                f"totaling {_format_currency(total_difference)} difference."
            )
        else:
            headline = "No qualifying deals were found."

        top_three = sorted(deals, key=lambda d: d["difference"], reverse=True)[:3]
        if top_three:
            parts = [
                f"{deal['deal_id']}: {_format_currency(deal['difference'])}"
                for deal in top_three
            ]
            top_contributors = (
                "Top 3 deals contributing to variance: " + ", ".join(parts)
            )
        else:
            top_contributors = "No significant deal variances detected."

        if unregistered_costs:
            top_unregistered = max(
                unregistered_costs, key=lambda item: item["impact"], default=None
            )
            if top_unregistered:
                unregistered_summary = (
                    "Largest unregistered impact: "
                    f"{top_unregistered['cost_type']} "
                    f"({_format_currency(top_unregistered['impact'])})."
                )
            else:
                unregistered_summary = "Unregistered cost details unavailable."
        else:
            unregistered_summary = "No unregistered cost types detected."

        recommendations: List[str] = []
        if top_three:
            recommendations.append(
                "Review the top contributing deals for manual confirmation of quantities."
            )
        if unregistered_costs:
            impacted = sorted(
                unregistered_costs, key=lambda item: item["impact"], reverse=True
            )[:2]
            names = ", ".join(item["cost_type"] for item in impacted)
            recommendations.append(
                f"Ensure cost registration for high impact types: {names}."
            )
        if patterns["repeating_patterns"]:
            names: List[str] = []
            for pattern in patterns["repeating_patterns"]:
                names.extend(pattern["cost_types"])
            unique = ", ".join(sorted(set(names)))
            recommendations.append(
                f"Investigate systematic issues causing repeated gaps in {unique}."
            )
        if not recommendations:
            recommendations.append("No immediate actions detected; continue monitoring.")

        return {
            "headline": headline,
            "top_contributors": top_contributors,
            "unregistered_costs": unregistered_summary,
            "recommended_actions": recommendations,
        }


__all__ = ["DealComparisonAnalyzer"]
