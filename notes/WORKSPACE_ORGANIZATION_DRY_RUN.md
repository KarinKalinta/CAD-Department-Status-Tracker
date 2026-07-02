# Workspace Organization Dry Run

Date prepared: 2026-07-02  
Phase: 1 scan and dry-run plan only  
Master folder path: `C:\Users\User\Desktop\Codex Dumpster`

## Safety Status

- No files were moved.
- No files were renamed.
- No backup was created.
- No project folders were created.
- This file is the only new root documentation file created during Phase 1.

## Scan Limits Used

The scan intentionally avoided deep inspection and skipped project-heavy or generated areas where possible:

- `.git`
- `node_modules`
- `.venv`
- `venv`
- `__pycache__`
- `.pytest_cache`
- `dist`
- `build`
- cache folders
- backup folders unless clearly relevant
- generated/render/image folders unless they were top-level project clues

Only folder names, immediate subfolders, README/markdown files, obvious scripts, entry/config files, and top-level file names were used for this dry-run.

## Existing File-Organization Logic Search

Requested names searched:

- `file_workflow_app`
- `naming_engine.py`
- `validator.py`
- `workflow_engine.py`
- `manifest_service.py`
- `template_service.py`
- `dependency_tracker.py`
- `script_registry.py`
- `script_runner.py`
- `log_service.py`
- `yaml_service.py`
- `generic_project.yaml`
- `manifest.yaml`

Result: no matching files or folders were found inside `C:\Users\User\Desktop\Codex Dumpster` outside skipped cache/backup paths.

Related existing organization evidence:

- `REORGANIZATION_REPORT.md` exists at workspace root.
- `Jewelry Research\Backup_Before_Reorganization` exists.
- Multiple `Project_###_*` folders already follow the target project structure.

Manual review note: the old file-organizer logic may exist outside this master folder. Do not assume it is absent from the computer, only absent from this workspace scan.

## Detected Projects

### Already Structured Projects

These already match the target folder structure and should generally be left in place unless the user approves a larger cleanup:

1. `Project_001_Organic_Jewelry`
2. `Project_002_Research_Slide_And_Mood_Boards`
3. `Project_003_Research_V1_App`
4. `Project_004_Silver_Production_Selection`
5. `Project_005_Vendor_Selection`
6. `Project_006_CAD_Skill_Matrix`
7. `Project_007_Uhrenschmuck24_Ladies_Jewellery_PD_Report`
8. `Project_008_Karin_Kalinta_CV`
9. `Project_009_Plan_Of_Action_Tracker`
10. `Project_010_Workspace_Unclear_Archive`

Proposed action: no movement during Phase 2 unless specifically approved. These are already organized.

### Project Candidates Requiring Normalization

#### Project Candidate: CAD Department Excel Tracker / CAD Status Tracker

Current location: loose files in workspace root  
Likely active: yes  
Reason: many recent `.xlsx`, `.py`, `.ps1`, and `.ts` files from 2026-06-29 to 2026-07-01.

Proposed project folder:

`Project_011_CAD_Status_Tracker`

Proposed structure:

```text
Project_011_CAD_Status_Tracker/
├── source_files/
├── scripts/
├── assets/
├── references/
├── exports/
├── notes/
├── archive/
└── README.md
```

Files proposed for `source_files/`:

- `cad_department_excel_tracker_office_script.ts`
- `CAD_Department_Plan_of_Action.xlsx`
- `CAD_Status_Tracker_Dashboard_working.xlsx`
- `CAD_Status_Tracker_Dashboard_FIXED.xlsx`
- `CAD_Status_Tracker_Batch_Aware.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_LIVE_FORMULAS.xlsx`

Files proposed for `scripts/`:

- `create_cad_department_excel_tracker.py`
- `create_batch_aware_cad_tracker.py`
- `make_manager_dashboard.py`
- `make_dashboard_live_table_refs.py`
- `add_priority_queue_dashboard.py`
- `fix_owner_workload_active.py`
- `fix_owner_workload_structured.py`
- `check_manager_dashboard_excel.ps1`

Files proposed for `assets/`:

- none detected at root for this project.

Files proposed for `references/`:

- `CAD_Department_Plan_of_Action_reference_layout.xlsx`
- `CAD_Department_Plan_of_Action_simple_timeline.xlsx`
- `CAD_Department_Plan_of_Action_visible_timeline.xlsx`
- `CAD_Department_Plan_of_Action_visible_timeline_english_dates.xlsx`
- `CAD_Department_Plan_of_Action_with_timeline.xlsx`

Files proposed for `exports/`:

- final usable workbook variants may belong here after manual selection.
- Candidate: `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_FINAL.xlsx`
- Candidate: `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_OWNER_WORKLOAD_FIXED.xlsx`

Files proposed for `notes/`:

- none detected at root for this project.

Files proposed for `archive/`:

- `CAD_Department_Plan_of_Action_auto_timeline.xlsx`
- `CAD_Department_Plan_of_Action_auto_timeline_8_weeks.xlsx`
- `CAD_Department_Plan_of_Action_auto_timeline_colored_inputs.xlsx`
- `CAD_Department_Plan_of_Action_auto_timeline_no_hash_dates.xlsx`
- `CAD_Department_Plan_of_Action_bottom_timeline.xlsx`
- `CAD_Department_Plan_of_Action_bottom_timeline_compact.xlsx`
- `CAD_Department_Plan_of_Action_bottom_timeline_editable_lists.xlsx`
- `CAD_Department_Plan_of_Action_bottom_timeline_project_first.xlsx`
- `CAD_Department_Plan_of_Action_bottom_timeline_readable.xlsx`
- `CAD_Department_Plan_of_Action_bottom_timeline_wide.xlsx`
- `CAD_Department_Plan_of_Action_deliverable_after_request.xlsx`
- `CAD_Department_Plan_of_Action_ready.xlsx`
- `CAD_Department_Plan_of_Action_ready_colored_boxes.xlsx`
- `CAD_Department_Plan_of_Action_request_type.xlsx`
- `CAD_Department_Plan_of_Action_timeline_check.xlsx`
- `CAD_Department_Plan_of_Action_with_samples.xlsx`
- `CAD_Status_Tracker_Batch_Aware_BATCH_SUMMARY_FIXED.xlsx`
- `CAD_Status_Tracker_Batch_Aware_colored_boxes.xlsx`
- `CAD_Status_Tracker_Batch_Aware_FIXED_FORMULAS.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_FINAL_OWNER_FIXED_source.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_FINAL_source.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_FIXED_source.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_OWNER_WORKLOAD_FIXED_source.xlsx`
- `CAD_Status_Tracker_Batch_Aware_OWNERS_UPDATED.xlsx`
- `CAD_Status_Tracker_Batch_Aware_OWNERS_UPDATED_source.xlsx`
- `CAD_Status_Tracker_Batch_Aware_REPAIRED.xlsx`
- `CAD_Status_Tracker_Batch_Aware_WITH_CHARTS.xlsx`
- `CAD_Status_Tracker_Batch_Aware_WITH_SAMPLES.xlsx`

Unclear files:

- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_LIVE_FORMULAS.xlsx` may be superseded by a live `.xlsm` file outside this workspace at `C:\Users\User\Desktop\All\CAD Status\...`. Do not move without confirming which is authoritative.

Risks:

- Many workbook versions are sequential iterations with similar names.
- Some files may be duplicate exports, sources, or repaired copies.
- User has been actively working outside this master folder in `C:\Users\User\Desktop\All\CAD Status`; do not assume the root copy is current.

#### Project Candidate: Job Search Thailand Singapore UAE

Current locations:

- `Job_Search_Thailand_Singapore_UAE`
- nested duplicate under `Job_Search_GitHub_Upload\Job_Search_Thailand_Singapore_UAE`

Likely active: older/possibly completed  
Reason: README at root says this was active on 2026-06-24, but current work has shifted to CAD tracker.

Proposed project folder:

`Project_012_Job_Search_Thailand_Singapore_UAE`

Files proposed for `source_files/`:

- `Job_Search_Thailand_Singapore_UAE\*.csv`
- `Job_Search_Thailand_Singapore_UAE\*.gs`

Files proposed for `scripts/`:

- `Job_Search_Thailand_Singapore_UAE\google_sheets_job_search_automation.gs`
- `Job_Search_Thailand_Singapore_UAE\strategic_job_research_system.gs`

Files proposed for `references/`:

- `Job_Search_Thailand_Singapore_UAE\sources_2026-06-24.md`
- `Job_Search_Thailand_Singapore_UAE\background_info_needed.md`

Files proposed for `notes/`:

- `Job_Search_Thailand_Singapore_UAE\summary_and_strategy_2026-06-24.md`
- `Job_Search_Thailand_Singapore_UAE\strategic_job_research_system_setup.md`
- `Job_Search_Thailand_Singapore_UAE\google_sheets_automation_setup.md`
- `Job_Search_Thailand_Singapore_UAE\opening_status_and_requirements_check_2026-06-24.md`

Files proposed for `archive/`:

- duplicate nested copy under `Job_Search_GitHub_Upload\Job_Search_Thailand_Singapore_UAE`, pending manual approval.

Possible shared/overlapping files:

- `.github\workflows\validate-job-tracker.yml`
- `GITHUB_UPLOAD_CHECKLIST.md`
- root `README.md`

Manual review:

- Decide whether `Job_Search_GitHub_Upload` is a deployment wrapper or duplicate archive. Do not merge automatically.

#### Project Candidate: Google Sheets Job Status Tracker

Current location:

`job_status_tracker_google_sheets`

Likely active/complete: complete template project  
Reason: already has README, docs, samples, Apps Script structure.

Proposed project folder:

Already acceptable as `job_status_tracker_google_sheets`, or normalize to:

`Project_013_Job_Status_Tracker_Google_Sheets`

Files proposed for `source_files/`:

- none; Apps Script source belongs in `scripts/` or preserved in current `apps_script/`.

Files proposed for `scripts/`:

- `job_status_tracker_google_sheets\apps_script\Code.gs`

Files proposed for `references/`:

- `job_status_tracker_google_sheets\docs\tracker_structure.md`

Files proposed for `notes/`:

- `job_status_tracker_google_sheets\docs\setup_instructions.md`
- `job_status_tracker_google_sheets\docs\manager_usage_guide.md`
- `job_status_tracker_google_sheets\changelog.md`

Files proposed for `exports/`:

- `job_status_tracker_google_sheets\samples\sample_data.csv`

Manual review:

- This project already has a meaningful custom structure. Renaming/restructuring may reduce clarity. Recommended: leave intact or add missing standard folders without moving files.

#### Project Candidate: Approved Designs Extraction

Current location:

`Approved_Designs_Extraction`

Likely status: old but coherent  
Reason: contains source PDFs, rendered pages, crops, contact sheets, scripts, and extracted workbooks.

Proposed project folder:

`Project_014_Approved_Designs_Extraction`

Files proposed for `source_files/`:

- `Approved_Designs_Extraction\source_pdfs\*`
- source Excel workbooks if they are original extracted data.

Files proposed for `scripts/`:

- `Approved_Designs_Extraction\add_bb3_1_bb4_to_approved_tracker.py`
- `Approved_Designs_Extraction\add_bb5_to_approved_tracker.py`
- `Approved_Designs_Extraction\add_photos_to_approved_workbook.py`

Files proposed for `assets/`:

- `Approved_Designs_Extraction\approved_design_crops\*`
- `Approved_Designs_Extraction\rendered_pages\*`
- image folders, if preserved as project assets.

Files proposed for `exports/`:

- `Approved_Designs_Extraction\Approved_Designs_Extracted*.xlsx`
- `Approved_Designs_Extraction\contact_sheets\*`
- `Approved_Designs_Extraction\text_exports\*`

Manual review:

- Image/render folders were not deeply inspected. Treat as project-owned but large/generated.

#### Project Candidate: Folder Location Status Tracking For Design

Current location:

`folder location status tracking for design`

Likely status: old but coherent  
Reason: contains tracking workbooks and generated thumbnail folders.

Proposed project folder:

`Project_015_Folder_Location_Status_Tracking_For_Design`

Files proposed for `source_files/`:

- original/source workbook if identifiable after approval.

Files proposed for `assets/`:

- `extraction_contact_sheets`
- `no_outlier_thumbnails_2025`
- `presentation_thumbnails_2025`
- `shared_tab_thumbnails_2025`
- `shared_tab_thumbnails_2025_copy`

Files proposed for `exports/`:

- `folder_location_status_tracking_for_design*.xlsx`

Manual review:

- Multiple similarly named workbooks require user choice for current/final file.

#### Project Candidate: Jewelry Master Workbook

Current location:

`Jewelry_Master_Workbook`

Likely status: old/coherent  
Reason: already has numbered internal folders and README.

Proposed project folder:

`Project_016_Jewelry_Master_Workbook`

Files proposed for `source_files/`:

- `01_Source_Files`

Files proposed for `scripts/`:

- `05_Scripts`

Files proposed for `assets/`:

- `03_Rendered_Pages`
- `04_Cropped_Images`
- `07_Contact_Sheets`

Files proposed for `exports/`:

- `00_READY_TO_USE`
- `02_Workbook_Versions`
- `06_Text_Exports`

Files proposed for `notes/`:

- `README.txt`

Manual review:

- Current folder already has a logical structure. Recommended: leave intact or map folders into standard names only after approval.

#### Project Candidate: Working / JewelCAD to 3DM

Current location:

`Working`

Likely status: unclear/old  
Reason: contains `jewelcad_to_3dm`, HTML/JS, PowerShell, JPG/SVG exports.

Proposed project folder:

`Project_017_JewelCAD_To_3DM_Working`

Files proposed for `source_files/`:

- `Working\index.html`
- `Working\app.js`
- `Working\jewelcad_to_3dm\*` excluding generated outputs.

Files proposed for `scripts/`:

- `Working\build_engraving_curves.ps1`
- `Working\serve.ps1`

Files proposed for `assets/`:

- `Working\C06047Y_2000x2000`
- `.jpg`
- `.svg`

Files proposed for `notes/`:

- `Working\README.md`
- `Working\jewelcad_to_3dm\README.md`

Manual review:

- `Working` is generic and risky to move without confirming current use.

#### Project Candidate: UXUI

Current location:

`UXUI`

Likely status: unclear/old  
Reason: folder exists but immediate scan showed no children.

Proposed project folder:

`Project_018_UXUI`

Manual review:

- Empty or inaccessible folder. Do not move until user confirms whether it should be archived.

#### Project Candidate: GitHub Upload / Repository Wrapper

Current location:

`Job_Search_GitHub_Upload`

Likely status: duplicate/deployment wrapper  
Reason: contains `.github`, README, checklist, and nested Job Search project.

Proposed project folder:

`Project_019_Job_Search_GitHub_Upload`

Files proposed for `source_files/`:

- none until duplicate relationship is clarified.

Files proposed for `notes/`:

- `Job_Search_GitHub_Upload\README.md`
- `Job_Search_GitHub_Upload\GITHUB_UPLOAD_CHECKLIST.md`

Files proposed for `references/`:

- `Job_Search_GitHub_Upload\.github\workflows\validate-job-tracker.yml`

Manual review:

- Overlaps with `Job_Search_Thailand_Singapore_UAE`.
- Do not merge automatically.

## Root Files Proposed For Movement

### Likely CAD Status Tracker Project Files

Move into `Project_011_CAD_Status_Tracker` after approval:

#### `source_files/`

- `cad_department_excel_tracker_office_script.ts`
- `CAD_Department_Plan_of_Action.xlsx`
- `CAD_Status_Tracker_Dashboard_working.xlsx`
- `CAD_Status_Tracker_Dashboard_FIXED.xlsx`
- `CAD_Status_Tracker_Batch_Aware.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_LIVE_FORMULAS.xlsx`

#### `scripts/`

- `add_priority_queue_dashboard.py`
- `create_batch_aware_cad_tracker.py`
- `create_cad_department_excel_tracker.py`
- `fix_owner_workload_active.py`
- `fix_owner_workload_structured.py`
- `make_dashboard_live_table_refs.py`
- `make_manager_dashboard.py`
- `check_manager_dashboard_excel.ps1`

#### `references/`

- `CAD_Department_Plan_of_Action_reference_layout.xlsx`
- `CAD_Department_Plan_of_Action_simple_timeline.xlsx`
- `CAD_Department_Plan_of_Action_visible_timeline.xlsx`
- `CAD_Department_Plan_of_Action_visible_timeline_english_dates.xlsx`
- `CAD_Department_Plan_of_Action_with_timeline.xlsx`

#### `exports/`

- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_FINAL.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_OWNER_WORKLOAD_FIXED.xlsx`

#### `archive/`

- all older/intermediate CAD workbook versions listed in the CAD project section above.

### Likely Job Search Project Files

Potentially move into `Project_012_Job_Search_Thailand_Singapore_UAE` after approval:

- `GITHUB_UPLOAD_CHECKLIST.md`
- root `README.md`

Manual review:

- Root README currently describes the job-search project, but the workspace is now used for many projects. Recommend replacing root README later with a master workspace README and moving current README content into the job-search project.

### General Workspace Documentation

Keep at root:

- `WORKSPACE_ORGANIZATION_DRY_RUN.md`
- `REORGANIZATION_REPORT.md`

Manual review:

- `REORGANIZATION_REPORT.md` documents an earlier reorganization and should probably remain in root `notes/` or master documentation, not inside a single project.

## Unclear Files

- `.agents`
- `.github`
- `README.md`
- `GITHUB_UPLOAD_CHECKLIST.md`
- `REORGANIZATION_REPORT.md`

Reason:

- These may be master/repository-level files rather than project files.
- Do not move without approval.

## Duplicate Or Overlapping Files

### Job Search Duplicate

- `Job_Search_Thailand_Singapore_UAE`
- `Job_Search_GitHub_Upload\Job_Search_Thailand_Singapore_UAE`

Risk:

- Same project appears in two places.
- One may be a deploy/upload copy.
- Do not merge automatically.

### Jewelry Research / Existing Project_### Folders

- `Jewelry Research\Backup_Before_Reorganization`
- root `Project_001_*` through `Project_010_*`
- `REORGANIZATION_REPORT.md`

Risk:

- Existing `Project_###` folders may already be reorganized output from `Jewelry Research\Master_Projects`.
- Do not move backup files into current project folders without approval.

### CAD Tracker Workbook Versions

Many CAD tracker workbooks are sequential variants. Examples:

- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_FINAL.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_OWNER_WORKLOAD_FIXED.xlsx`
- `CAD_Status_Tracker_Batch_Aware_MANAGER_VIEW_LIVE_FORMULAS.xlsx`
- source/fixed/repaired/chart/sample variants

Risk:

- Some are source snapshots.
- Some are generated versions.
- The currently active file is outside this workspace under `C:\Users\User\Desktop\All\CAD Status`.
- Do not delete or overwrite any of them.

## Possible Shared Files

- `.github\workflows\validate-job-tracker.yml`
- `GITHUB_UPLOAD_CHECKLIST.md`
- root `README.md`
- `REORGANIZATION_REPORT.md`

Recommended handling:

- Keep at root during Phase 2 unless the user explicitly approves moving them.
- If moved, copy references into relevant project `references/` folders and document overlap in README files.

## Manual Review Notes

Before Phase 2, user approval is needed for:

1. Whether to normalize already-structured `Project_001` through `Project_010` or leave them untouched.
2. Whether to create `Project_011_CAD_Status_Tracker` for all root CAD files.
3. Which CAD workbook is the current/final source of truth.
4. Whether root CAD versions should go to `archive/` or remain in `exports/`.
5. Whether `Job_Search_GitHub_Upload` should remain separate from `Job_Search_Thailand_Singapore_UAE`.
6. Whether root `README.md` should become a master README or be moved into the job-search project.
7. Whether `Working` is active or should be archived.
8. Whether `UXUI` is empty/old and safe to archive.
9. Whether `Jewelry_Master_Workbook` should keep its current internal folder names.
10. Whether old backup folders should be excluded from all movement.

## Risks Before Execution

- The workspace already contains an earlier reorganization, so a second reorganization could create duplicate `Project_###` numbering if not handled carefully.
- CAD workbook versions are highly similar; moving the wrong file to archive could confuse active work.
- Some active CAD work is outside the master folder, so the master folder does not contain the full current state.
- `Job_Search_GitHub_Upload` and `Job_Search_Thailand_Singapore_UAE` overlap.
- Existing `.github` workflow may depend on current paths.
- Moving root README/checklist files could affect GitHub upload context.
- Generic folders such as `Working` and `UXUI` require user confirmation.

## Phase 2 Draft Execution Plan

Only after explicit approval:

1. Create backup folder:

   `C:\Users\User\Desktop\Codex Dumpster\_backup_before_reorganization`

2. Copy the current workspace structure into the backup.
3. Create approved project folders only.
4. Create standard subfolders:

   ```text
   source_files/
   scripts/
   assets/
   references/
   exports/
   notes/
   archive/
   README.md
   ```

5. Move files only according to the approved sections above.
6. Put unclear, old, duplicate, temporary, or uncertain files in `archive/` only if approved.
7. Preserve original files and document overlaps in README files.

## Summary

Detected project/project-candidate count: 19

Likely active projects:

- `Project_011_CAD_Status_Tracker` candidate from root CAD tracker files.
- External active related location noted: `C:\Users\User\Desktop\All\CAD Status`.

Likely complete/older but coherent projects:

- `Project_001` through `Project_010`
- `job_status_tracker_google_sheets`
- `Approved_Designs_Extraction`
- `folder location status tracking for design`
- `Jewelry_Master_Workbook`
- `Job_Search_Thailand_Singapore_UAE`

Old/unclear projects:

- `Working`
- `UXUI`
- `Job_Search_GitHub_Upload` as possible wrapper/duplicate

Risky files to move:

- root CAD workbook versions
- root README/checklist files
- `.github` workflow
- duplicate job-search folders
- generic `Working` folder
- existing jewelry research backup folders

Approval needed before Phase 2:

- Confirm which proposed project folders to create.
- Confirm whether existing organized projects should remain untouched.
- Confirm current CAD tracker source of truth.
- Confirm how to handle root README/GitHub files.
- Confirm whether `Working` and `UXUI` should be archived or preserved.
