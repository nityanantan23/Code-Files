import streamlit as st
import io
import pandas as pd
from backend import (
    analyze_documents,
    highlight_mistakes,
    apply_corrections,
    insert_missing_sections,
)
import time

# --- Streamlit Page Config ---
st.set_page_config(page_title="JIWE Document Formatter", layout="wide")
st.title("📑 JIWE Document Formatter")
st.write(
    "Upload your **Template** and **Manuscript** DOCX files, analyze formatting, and download the mistakes as Excel."
)

# Accuracy disclaimer on main page
st.warning(
    """
⚠️ **Accuracy Note**: This tool has approximately 75-81% detection accuracy. 
Always manually double-check the results for complete formatting verification.
"""
)

# --- Initialize Session State ---
if "analysis_done" not in st.session_state:
    st.session_state.analysis_done = False
if "excel_bytes" not in st.session_state:
    st.session_state.excel_bytes = None
if "mistakes_df" not in st.session_state:
    st.session_state.mistakes_df = None
if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0
if "processed_doc_bytes" not in st.session_state:
    st.session_state.processed_doc_bytes = None
if "processed_doc_name" not in st.session_state:
    st.session_state.processed_doc_name = None
if "processing_done" not in st.session_state:
    st.session_state.processing_done = False
if "force_clear" not in st.session_state:
    st.session_state.force_clear = False
if "highlight_debug_info" not in st.session_state:
    st.session_state.highlight_debug_info = {"summary": {}, "paragraphs": []}

# --- Tabs for clean UI ---
tabs = st.tabs(
    [
        "📂 Upload Files & Analyze",
        "📊 Results & Download (Excel)",
        "🛠️ Auto Process & Download Processed Journal",
        "📖 User Manual",
    ]
)

# ------------------------------
# TAB 1: Upload Files & Analyze
# ------------------------------
with tabs[0]:
    st.header("📂 Upload Files & Analyze")

    # Accuracy disclaimer
    st.info(
        """
    💡 **Accuracy Notice**: This program detects approximately 75-81% of formatting issues. 
    Manual verification is still required for complete accuracy.
    """
    )

    # File Upload Section
    st.subheader("📄 Upload Files")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("**Template File (.docx)**")
        # Use a unique key that changes when reset_counter changes
        template_key = f"template_uploader_{st.session_state.reset_counter}"
        template_file = st.file_uploader(
            "Choose Template file",
            type="docx",
            key=template_key,
            label_visibility="collapsed",
        )
        st.caption(
            "💡 The template should contain the standard formatting rules and styles"
        )

    with col2:
        st.markdown("**Manuscript File (.docx)**")
        # Use a unique key that changes when reset_counter changes
        manuscript_key = f"manuscript_uploader_{st.session_state.reset_counter}"
        manuscript_file = st.file_uploader(
            "Choose Manuscript file",
            type="docx",
            key=manuscript_key,
            label_visibility="collapsed",
        )
        st.caption(
            "💡 This document will be checked against the template for formatting compliance"
        )

    st.divider()

    # Analysis Section
    st.subheader("🔍 Analyze Document")

    # Additional guidance
    if not template_file and not manuscript_file:
        st.info("👆 Please upload both files above first")
    elif template_file and manuscript_file:
        st.success("✅ Both files uploaded successfully! Ready to analyze.")

    analyze_btn = st.button(
        "🔍 Analyze Document",
        type="primary",
        help="Compare manuscript against the template",
    )

    # Analysis logic
    if analyze_btn:
        if not template_file or not manuscript_file:
            st.error(
                "⚠️ Please upload both template and manuscript files before analyzing."
            )
        else:
            st.info("⏳ Converting documents to XML and analyzing... Please wait.")

            try:
                findings, missing, xml_previews = analyze_documents(
                    template_file, manuscript_file
                )

                # Convert findings to DataFrame
                df_full = pd.DataFrame(findings)

                if not df_full.empty:
                    preferred_order = [
                        "type",
                        "section",
                        "found",
                        "expected",
                        "snippet",
                        "suggested_fix",
                        "paragraph_indices",
                        "pages",
                        "suggested_action",
                    ]
                    ordered_columns = [
                        col for col in preferred_order if col in df_full.columns
                    ]
                    ordered_columns += [
                        col for col in df_full.columns if col not in ordered_columns
                    ]
                    df_full = df_full[ordered_columns]
                st.session_state.mistakes_df = df_full

                df_display = st.session_state.mistakes_df

                # Save Excel report (using display version)
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_display.to_excel(writer, index=False, sheet_name="Mistakes")
                    if missing:
                        pd.DataFrame({"Missing Sections": missing}).to_excel(
                            writer, index=False, sheet_name="Missing Sections"
                        )
                st.session_state.excel_bytes = output.getvalue()

                st.session_state.analysis_done = True
                st.session_state.missing_sections = missing or []
                st.success(
                    f"✅ Analysis complete! Found {len(df_display)} formatting issues."
                )

                # Show mistakes on screen (using display version)
                if not df_display.empty:
                    st.subheader("📋 Detected Formatting Issues")
                    st.dataframe(df_display, use_container_width=True)
                else:
                    st.success("🎉 No formatting mistakes found!")

            except Exception as e:
                st.error(f"❌ Analysis failed: {e}")
                import traceback

                st.error(f"Detailed error: {traceback.format_exc()}")
                st.session_state.analysis_done = False

    # Reset button - ALWAYS VISIBLE
    st.divider()

    # Compact left-aligned reset button
    if st.button(
        "🔄 Reset and Clear All Data",
        type="secondary",
        use_container_width=False,  # Don't use full container width
        help="Completely clear all uploaded files and analysis results",
    ):

        # Clear all session state variables
        st.session_state.analysis_done = False
        st.session_state.excel_bytes = None
        st.session_state.mistakes_df = None
        st.session_state.processed_doc_bytes = None
        st.session_state.processed_doc_name = None
        st.session_state.processing_done = False

        # Force clear by changing the reset counter
        st.session_state.reset_counter += 1

        # Show success message and force rerun
        st.success(
            "✅ **System Reset Complete!** All files and data have been cleared."
        )
        st.info("📝 **You can now upload new files**")

        # Add a small delay and rerun
        time.sleep(1)
        st.rerun()

# ------------------------------
# TAB 2: Results & Download (Excel)
# ------------------------------
with tabs[1]:
    st.header("📊 Results & Download (Excel)")

    # Accuracy disclaimer
    st.warning(
        """
    📊 **Results Accuracy**: The detected issues represent approximately 75-81% of actual formatting problems. 
    Manual review is essential for complete accuracy.
    """
    )

    if not st.session_state.analysis_done:
        st.warning(
            "⚠️ Please analyze the document in the 'Upload Files & Analyze' tab first."
        )
    else:
        st.subheader("📋 Detected Formatting Issues")

        if st.session_state.mistakes_df is not None:
            df_display = st.session_state.mistakes_df
            st.dataframe(df_display, use_container_width=True)

            # Show summary stats
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Issues", len(df_display))
            with col2:
                issue_types = (
                    df_display["type"].nunique() if "type" in df_display.columns else 0
                )
                st.metric("Issue Types", issue_types)
            with col3:
                st.metric("Detection Accuracy", "75-81%")

        st.divider()
        st.subheader("📥 Download Excel Report")

        if st.session_state.excel_bytes:
            st.download_button(
                "📥 Download Excel Report",
                st.session_state.excel_bytes,
                file_name="formatting_issues.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        # Reminder for manual check
        st.info(
            """
        🔍 **Important**: After downloading, please manually verify the document as this tool 
        may not catch all formatting issues (75-81% detection rate).
        """
        )

# ------------------------------
# TAB 3: Auto Process & Download Processed Journal
# ------------------------------
with tabs[2]:
    st.header("🛠️ Auto Process & Download Processed Journal")

    # Accuracy disclaimer
    st.warning(
        """
    ⚠️ **Processing Accuracy**: Auto-processing works on the detected issues (75-81% accuracy). 
    Manual verification is required to catch all formatting problems.
    """
    )

    # Check if analysis is done
    if not st.session_state.analysis_done:
        st.warning(
            """
        **⚠️ Please analyze the document first**
        
        Go to the 'Upload Files & Analyze' tab to:
        1. Upload your template and manuscript files
        2. Run the analysis
        3. Then come back here to process the document
        """
        )
    elif st.session_state.mistakes_df is None or st.session_state.mistakes_df.empty:
        st.success("🎉 No formatting issues detected! No processing needed.")
    else:
        # Processing Section
        st.subheader("🛠️ Auto Processing Options")

        st.info(
            """
        **Choose how you want to process the formatting issues:**
        - **🟨 Auto Highlight**: Highlight issues in yellow (safe - no changes made)
        - **🔧 Auto Correct**: Automatically fix formatting issues  
        - **⚡ Auto Correct & Highlight**: Fix issues AND highlight what was changed
        """
        )

        # Show detected issues
        st.subheader("📋 Issues to Process")
        df_display = st.session_state.mistakes_df
        st.dataframe(df_display, use_container_width=True)

        # Processing mode selection
        col1, col2, col3 = st.columns(3)

        with col1:
            highlight_btn = st.button(
                "🟨 Auto Highlight Only",
                use_container_width=True,
                help="Highlight issues in yellow without making changes",
            )

        with col2:
            correct_btn = st.button(
                "🔧 Auto Correct Only",
                use_container_width=True,
                help="Automatically fix formatting issues",
            )

        with col3:
            both_btn = st.button(
                "⚡ Auto Correct & Highlight",
                use_container_width=True,
                type="primary",
                help="Fix issues AND highlight what was changed",
            )

        # Process based on selection
        if highlight_btn or correct_btn or both_btn:
            with st.spinner("🔄 Processing document..."):
                try:
                    # Get current file objects using the current keys
                    template_key = f"template_uploader_{st.session_state.reset_counter}"
                    manuscript_key = (
                        f"manuscript_uploader_{st.session_state.reset_counter}"
                    )

                    template_file = st.session_state.get(template_key)
                    manuscript_file = st.session_state.get(manuscript_key)

                    if not template_file or not manuscript_file:
                        st.error("❌ Files not found. Please re-upload files in Tab 1.")
                    else:
                        # Reset file pointers
                        template_file.seek(0)
                        manuscript_file.seek(0)

                        processed_bytes = None
                        highlight_debug = {"summary": {}, "paragraphs": []}
                        process_type = ""
                        process_description = ""
                        original_name = manuscript_file.name

                        # Prepare current manuscript bytes for processing
                        base_bytes = None
                        try:
                            manuscript_file.seek(0)
                            base_bytes = manuscript_file.read()
                            manuscript_file.seek(0)
                        except Exception:
                            base_bytes = None

                        missing_sections = (
                            st.session_state.get("missing_sections") or []
                        )

                        if highlight_btn:
                            # Highlight only
                            # Insert missing sections first (with yellow highlight), then highlight issues
                            if missing_sections:
                                inserted = insert_missing_sections(
                                    template_file,
                                    io.BytesIO(base_bytes or manuscript_file.read()),
                                    missing_sections,
                                )
                            else:
                                inserted = base_bytes or manuscript_file.read()
                            from io import BytesIO

                            inserted_stream = BytesIO(inserted)
                            template_file.seek(0)
                            processed_bytes, highlight_debug = highlight_mistakes(
                                template_file,
                                inserted_stream,
                                st.session_state.mistakes_df,
                            )
                            process_type = "HIGHLIGHTED"
                            process_description = "🟨 Highlighted Document"
                            success_message = "✅ Issues highlighted successfully!"

                        elif correct_btn:
                            # Correct only
                            corrected_bytes = apply_corrections(
                                template_file,
                                manuscript_file,
                                st.session_state.mistakes_df,
                            )
                            # Then insert missing sections (no additional highlight beyond inserted yellow)
                            if corrected_bytes is not None:
                                processed_bytes = insert_missing_sections(
                                    template_file,
                                    io.BytesIO(corrected_bytes),
                                    missing_sections,
                                )
                            else:
                                processed_bytes = None
                            process_type = "CORRECTED"
                            process_description = "🔧 Corrected Document"
                            success_message = "✅ Corrections applied successfully!"

                        elif both_btn:
                            # Correct first, then highlight what was corrected
                            corrected_bytes = apply_corrections(
                                template_file,
                                manuscript_file,
                                st.session_state.mistakes_df,
                            )
                            if corrected_bytes:
                                # Insert missing sections after correction
                                inserted_bytes = insert_missing_sections(
                                    template_file,
                                    io.BytesIO(corrected_bytes),
                                    missing_sections,
                                )
                                # Then highlight the issues
                                from io import BytesIO

                                corrected_stream = BytesIO(inserted_bytes)
                                template_file.seek(0)  # Reset template
                                processed_bytes, highlight_debug = highlight_mistakes(
                                    template_file,
                                    corrected_stream,
                                    st.session_state.mistakes_df,
                                )
                                process_type = "CORRECTED_AND_HIGHLIGHTED"
                                process_description = (
                                    "⚡ Corrected & Highlighted Document"
                                )
                                success_message = (
                                    "✅ Corrections applied and issues highlighted!"
                                )
                            else:
                                st.error("❌ Correction failed, cannot highlight")
                                processed_bytes = None

                        if processed_bytes:
                            st.session_state.processed_doc_bytes = processed_bytes
                            st.session_state.processed_doc_name = (
                                f"{process_type}_{original_name}"
                            )
                            st.session_state.processing_done = True
                            st.session_state.process_description = process_description
                            st.session_state.highlight_debug_info = highlight_debug

                            st.success(success_message)

                            # Accuracy reminder
                            st.info(
                                """
                            📝 **Accuracy Reminder**: Processing is based on detected issues (75-81% accuracy). 
                            Please manually review the document to ensure all formatting is correct.
                            """
                            )

                        else:
                            st.error("❌ Processing failed. Please try again.")
                            st.session_state.highlight_debug_info = highlight_debug or {
                                "summary": {},
                                "paragraphs": [],
                            }

                except Exception as e:
                    st.error(f"❌ Processing failed: {str(e)}")
                    import traceback

                    st.error(f"Detailed error: {traceback.format_exc()}")

        debug_info = st.session_state.get("highlight_debug_info")
        if debug_info:
            if isinstance(debug_info, list):
                debug_info = {"summary": {}, "paragraphs": debug_info}
            summary = debug_info.get("summary") or {}
            paragraph_previews = debug_info.get("paragraphs") or []
            has_summary = bool(
                summary.get("row_count")
                or summary.get("note")
                or summary.get("sample_rows")
            )
            has_paragraphs = bool(paragraph_previews)
            if has_summary or has_paragraphs:
                with st.expander("🔍 Highlight debug details", expanded=False):
                    if has_summary:
                        st.markdown("**mistakes_df summary**")
                        st.write(f"Rows: {summary.get('row_count', 0)}")
                        columns = summary.get("columns") or []
                        st.write(
                            f"Columns: {', '.join(columns) if columns else 'None'}"
                        )
                        note = summary.get("note")
                        if note:
                            st.info(note)
                        sample_rows = summary.get("sample_rows") or []
                        if sample_rows:
                            st.dataframe(pd.DataFrame(sample_rows))
                    if has_paragraphs:
                        st.markdown("**Paragraph highlights**")
                        for preview in paragraph_previews:
                            issue_types = (
                                ", ".join(preview.get("issue_types", [])) or "N/A"
                            )
                            issue_count = preview.get("issue_count", 0)
                            st.write(
                                f"Paragraph {preview.get('paragraph_index')} • "
                                f"{issue_count} issue(s) • Types: {issue_types}"
                            )
                            paragraph_text = preview.get("paragraph_text") or "—"
                            st.caption(paragraph_text)

        st.divider()

        # Download Section
        st.subheader("📥 Download Processed Document")

        if not st.session_state.get("processing_done", False):
            st.info(
                "👆 Process the document using the options above to enable download"
            )
        else:
            st.success("✅ Document has been processed and is ready for download!")

            if (
                st.session_state.processed_doc_bytes
                and st.session_state.processed_doc_name
            ):
                # Show what was done
                st.info(
                    f"**{st.session_state.get('process_description', 'Processed Document')}**"
                )

                # Final accuracy reminder
                st.info(
                    """
                🔍 **Final Verification Needed**: 
                - This processed document is based on 75-81% detection accuracy
                - Please manually review the entire document
                - Check for any formatting issues the tool might have missed
                """
                )

                st.download_button(
                    "📥 Download Processed Document",
                    st.session_state.processed_doc_bytes,
                    file_name=st.session_state.processed_doc_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )

                # Show processing summary
                st.subheader("🔍 Processing Details")
                st.write(f"**Document:** {st.session_state.processed_doc_name}")
                st.write(
                    f"**Total issues processed:** {len(st.session_state.mistakes_df)}"
                )

                # Show what was processed
                issue_types = st.session_state.mistakes_df["type"].value_counts()
                st.write("**Issues processed by type:**")
                for issue_type, count in issue_types.items():
                    st.write(f"- {issue_type}: {count}")

# ------------------------------
# TAB 4: User Manual Guide
# ------------------------------
with tabs[3]:
    st.header("📖 User Manual Guide")

    # Accuracy disclaimer at the top
    st.warning(
        """
    ⚠️ **Important Accuracy Notice**: 
    This tool detects approximately 75-81% of formatting issues. 
    **Always perform manual verification** for complete formatting accuracy.
    """
    )

    st.markdown(
        """
    ### 🚀 Quick Start Guide
    
    Follow these steps to use the JIWE Document Formatter:
    """
    )

    # Step-by-step instructions with proper numbering
    st.markdown(
        """
    **Step 1: 📂 Upload Files & Analyze**  
    - Upload Template and Manuscript DOCX files
    - Click 'Analyze Document' to check for formatting issues
    
    **Step 2: 📊 Review Results & Download Excel**  
    - Check the analysis results and detected issues
    - Download Excel report if needed
    
    **Step 3: 🛠️ Auto Process & Download**  
    - Choose processing method: Highlight, Correct, or Both
    - Download the processed document
    
    **Step 4: 🔍 Manual Verification**  
    - **Important**: Manually review the document (75-81% accuracy)
    - Check for any missed formatting issues
    """
    )

    st.divider()

    # Processing Options Explained
    st.subheader("🛠️ Processing Options")

    col1, col2, col3 = st.columns(3)

    with col1:
        st.markdown(
            """
        ### 🟨 Auto Highlight
        - **What it does**: Marks formatting issues in yellow
        - **Use when**: You want to see problems but fix them manually
        - **Result**: Document with yellow highlights on issues
        """
        )

    with col2:
        st.markdown(
            """
        ### 🔧 Auto Correct  
        - **What it does**: Automatically fixes formatting issues
        - **Use when**: You trust the automatic corrections
        - **Result**: Clean, corrected document
        """
        )

    with col3:
        st.markdown(
            """
        ### ⚡ Auto Correct & Highlight
        - **What it does**: Fixes issues AND highlights what was changed
        - **Use when**: You want to review what was automatically fixed
        - **Result**: Corrected document with highlights on fixed areas
        """
        )

    st.divider()

    # Accuracy Section
    st.subheader("🎯 Accuracy Information")

    st.markdown(
        """
    ### Understanding Detection Rates
    
    **Current Detection Accuracy**: 75-81%
    
    **What this means**:
    - The tool will catch **about half to two-thirds** of formatting issues
    - **Some issues may be missed** - manual review is essential
    - **False positives are possible** - some detected "issues" might be correct
    
    **Always double-check** your document manually before final submission.
    """
    )

    st.divider()

    # File Requirements
    st.subheader("📋 File Requirements")

    st.markdown(
        """
    - **Template File**: DOCX format containing correct formatting rules
    - **Manuscript File**: DOCX format that needs formatting check
    - **Output**: Processed DOCX file and/or Excel report
    """
    )

    st.divider()

    # Troubleshooting
    st.subheader("❓ Troubleshooting")

    troubleshooting = [
        "**Files won't upload?** Make sure they are DOCX format and always reset if you want upload/test new journals",
        "**Analysis failed?** Check that both files are uploaded and valid DOCX",
        "**No issues found?** Your document might already be properly formatted, but manually verify due to 75-81% accuracy",
        "**Processing stuck?** Try resetting and uploading files again",
        "**Unexpected results?** Remember the 75-81% accuracy - manual check is required",
    ]

    for issue in troubleshooting:
        st.write(f"• {issue}")

    st.divider()
