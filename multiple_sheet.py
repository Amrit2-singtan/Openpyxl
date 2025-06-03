from openpyxl import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# Your nested export structure
export_fields = {
    "Summary": {
        "Appraisal Year": {
            "Username": "appraisee.username",
            "Name": "appraisee.full_name",
            "Branch": "appraisee.detail.branch.name",
            "Job Title": "appraisee.detail.job_title.title",
            "Level": "appraisee.detail.employment_level.title",
            "Appraisal Period From": "annual_appraisal.start_date",
            "Appraisal Period To": "annual_appraisal.end_date",
            "Appraiser Name": "appraiser_full_name",
            "Appraiser Username": "appraiser_username",
            "Zone List": "zone_name",
            "Ethics and Integrity Declaration": "ethics_and_integrity_declaration",
            "Ethics and Integrity Declaration Comment": "ethics_and_integrity_comment",
            "Reviewer Name": "reviewer_full_name",
            "Reviewer Username": "reviewer_username",
            "Reviewers Endorsement (Comments)": "reviewer_comment",
            "Appraisee's Acknowledgment Comment": "appraisee_acknowledgment_comment",
            "Revised Zone List": "revised_zone_list",
            "Additional Notes": "additional_notes",
            "Final Zone List": "final_zone_list"
        }
    },
    "Detailed BCD Indicators": {
        "Appraisal Year": {
            "Appraiser Name": "appraiser_full_name",
            "Appraiser Username": "appraiser_username",
        },
        "Behavioral Indicators": {
            "ownership & Initiative": "",
            "Collaboration & Communication": "",
            "Customer & Stakeholder Experience": "",
        },
        "Compliance": {
            "Regulatory & Policy Adherence": "",
            "Job Knowledge & Application": "",
            "Risk Awareness & Accountability": "",
        },
        "Delivery": {
            "Achievement of Assigned Goals & Role-Specific Deliverables": "",
            "Quality of Work & Problem-Solving": "",
            "Process & Operational Efficiency": "",
        },
        "Final": {
            "Final Rating": "",
            "Zone list": "",
            "Final Zone list after revied from HR Committee": ""
        }
    },
    "Performance Gap Indicators": {
        "Appraisal Year": {
            "Username": "appraiser_username",
            "Name": "appraiser_full_name",
            "Branch": "appraisee.detail.branch.name",
            "Job Title": "appraisee.detail.job_title.title",
            "Level": "appraisee.detail.employment_level.title",
        },
        "Performance Gap Indicators": {
            "Performance Area": "",
            "Examples": "",
            "Performance Gap Observed": "",
        }
    },
    "Action Plan": {
        "Appraisal Year": {
            "Username": "appraiser_username",
            "Name": "appraiser_full_name",
            "Branch": "appraisee.detail.branch.name",
            "Job Title": "appraisee.detail.job_title.title",
            "Level": "appraisee.detail.employment_level.title",
        },
        "Recommended Action Plan for Improvement & Support Required": {
            "Action Plan for Improvement": "",
            "Support/Resources Required": ""
        }
    },
}

# Tab color presets
TAB_COLORS = ['FF9999', '99CCFF', 'CCFFCC', 'FFFF99', 'FFCCFF']

# Generate sample data for given headers
def generate_sample_data(headers, row_index=1):
    return [f"Sample {header} {row_index}" for header in headers]

# Extract two-level structured headers from nested dict
def extract_structured_headers(field_groups):
    top_headers = []
    sub_headers = []
    for group_title, fields in field_groups.items():
        field_names = list(fields.keys())
        top_headers.extend([group_title] * len(field_names))
        sub_headers.extend(field_names)
    return top_headers, sub_headers

# Main function to create workbook
def create_excel_from_export_fields(export_fields, filename):
    wb = Workbook()
    wb.remove(wb.active)

    for idx, (sheet_name, field_groups) in enumerate(export_fields.items()):
        ws = wb.create_sheet(title=sheet_name)
        ws.sheet_properties.tabColor = TAB_COLORS[idx % len(TAB_COLORS)]

        # Extract headers
        top_headers, sub_headers = extract_structured_headers(field_groups)
        num_columns = len(sub_headers)

        # First row: group titles
        ws.append(top_headers)

        # Merge same group headers
        current_group = None
        start_idx = 0
        for i, group in enumerate(top_headers + ["END"]):  # Add sentinel
            if group != current_group:
                if current_group is not None:
                    end_idx = i
                    if end_idx - start_idx > 1:
                        ws.merge_cells(
                            start_row=1, start_column=start_idx + 1,
                            end_row=1, end_column=end_idx
                        )
                current_group = group
                start_idx = i

        # Second row: sub-headers
        ws.append(sub_headers)

        # Align headers
        for row in ws.iter_rows(min_row=1, max_row=2, min_col=1, max_col=num_columns):
            for cell in row:
                cell.alignment = Alignment(horizontal="center", vertical="center")

        # Add sample data rows
        for i in range(1, 4):
            ws.append(generate_sample_data(sub_headers, i))

    wb.save(filename)

# Call to generate the file
create_excel_from_export_fields(export_fields, "appraisal_export_grouped.xlsx")
