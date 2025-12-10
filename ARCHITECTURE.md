# System Architecture Diagram

## High-Level Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                         USER BROWSER                             │
│                     http://localhost:5000                        │
└────────────────┬────────────────────────────────────────────────┘
                 │
                 │ HTTP/HTTPS
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                      FLASK WEB SERVER                            │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │                    Frontend Layer                         │  │
│  │  • templates/index.html (Dashboard UI)                    │  │
│  │  • static/js/app.js (JavaScript Logic)                    │  │
│  │  • Tailwind CSS + Font Awesome                            │  │
│  └──────────────────────────────────────────────────────────┘  │
│                            │                                     │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │                     API Layer                             │  │
│  │  • app/routes.py (REST Endpoints)                         │  │
│  │  • GET/POST/DELETE operations                             │  │
│  │  • File upload/download handlers                          │  │
│  └──────────────────────────────────────────────────────────┘  │
│                            │                                     │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │                   Business Logic Layer                    │  │
│  │                                                            │  │
│  │  ┌─────────────────┐  ┌─────────────────┐               │  │
│  │  │  JobManager     │  │  FileTracker    │               │  │
│  │  │  - Create jobs  │  │  - SHA tracking │               │  │
│  │  │  - Process jobs │  │  - File copies  │               │  │
│  │  │  - Track status │  │  - Change detect│               │  │
│  │  └─────────────────┘  └─────────────────┘               │  │
│  │                                                            │  │
│  │  ┌─────────────────┐  ┌─────────────────┐               │  │
│  │  │ DocumentParser  │  │ TemplateProc    │               │  │
│  │  │ - Parse Excel   │  │ - Process DOCX  │               │  │
│  │  │ - Extract vars  │  │ - Process XLSX  │               │  │
│  │  │ - Validate data │  │ - Process MSG   │               │  │
│  │  └─────────────────┘  └─────────────────┘               │  │
│  │                                                            │  │
│  │  ┌─────────────────┐                                      │  │
│  │  │ FormatConverter │                                      │  │
│  │  │ - To PDF        │                                      │  │
│  │  │ - To Word       │                                      │  │
│  │  │ - To Excel      │                                      │  │
│  │  │ - To MSG        │                                      │  │
│  │  └─────────────────┘                                      │  │
│  └──────────────────────────────────────────────────────────┘  │
│                            │                                     │
│  ┌──────────────────────────────────────────────────────────┐  │
│  │                     Data Layer                            │  │
│  │  • models/job.py (Job Model)                              │  │
│  │  • JSON file persistence                                  │  │
│  │  • File system storage                                    │  │
│  └──────────────────────────────────────────────────────────┘  │
└────────────────┬────────────────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────────────────┐
│                       FILE SYSTEM                                │
│  ┌──────────────┐  ┌──────────────┐  ┌──────────────┐         │
│  │   jobs/      │  │  storage/    │  │  uploads/    │         │
│  │  Job metadata│  │  Tracked     │  │  User        │         │
│  │  Output files│  │  file copies │  │  uploads     │         │
│  │  ZIP archives│  │  SHA hashes  │  │              │         │
│  └──────────────┘  └──────────────┘  └──────────────┘         │
└─────────────────────────────────────────────────────────────────┘
```

## Data Flow Diagram

```
┌──────────┐
│  User    │
└────┬─────┘
     │ 1. Upload files / Provide paths
     ▼
┌──────────────────┐
│  API Endpoint    │ POST /api/jobs
│  (routes.py)     │
└────┬─────────────┘
     │ 2. Create job request
     ▼
┌──────────────────┐
│  JobManager      │ create_job()
└────┬─────────────┘
     │ 3. Track files
     ▼
┌──────────────────┐
│  FileTracker     │ track_file()
└────┬─────────────┘
     │ 4. Calculate SHA-256, Copy files
     ▼
┌──────────────────┐
│  storage/        │ Local file copies
└──────────────────┘
     │ 5. Job created
     ▼
┌──────────────────┐
│  Job Object      │ status: PENDING
└────┬─────────────┘
     │ 6. Background processing starts
     ▼
┌──────────────────┐
│  JobManager      │ process_job()
└────┬─────────────┘
     │ 7. Parse data file
     ▼
┌──────────────────┐
│  DocumentParser  │ parse_excel_data()
└────┬─────────────┘
     │ 8. Extract variables & rows
     ▼
     │ For each data row:
     │ ┌─────────────────────────────┐
     │ │ 9. Process template         │
     ▼ │                             │
┌──────────────────┐                │
│ TemplateProc     │ process_template()
└────┬─────────────┘                │
     │ 10. Substitute ##variables## │
     ▼                               │
┌──────────────────┐                │
│  Processed Doc   │                │
└────┬─────────────┘                │
     │ 11. Convert to formats       │
     ▼                               │
┌──────────────────┐                │
│ FormatConverter  │ convert()      │
└────┬─────────────┘                │
     │ 12. Generate PDF/Word/etc    │
     ▼                               │
┌──────────────────┐                │
│  Output Files    │                │
└────┬─────────────┘                │
     │ ◄──────────────────────────┘
     │ 13. All rows processed
     ▼
┌──────────────────┐
│  JobManager      │ create_zip()
└────┬─────────────┘
     │ 14. Create ZIP archive
     ▼
┌──────────────────┐
│  job_xxx.zip     │
└────┬─────────────┘
     │ 15. Update job status: COMPLETED
     ▼
┌──────────────────┐
│  Job Object      │ status: COMPLETED
└────┬─────────────┘
     │ 16. User downloads
     ▼
┌──────────┐
│  User    │ Gets ZIP file
└──────────┘
```

## Job Processing Flow

```
         ┌──────────┐
         │  PENDING │
         └─────┬────┘
               │ JobManager.process_job()
               ▼
         ┌──────────┐
         │PROCESSING│
         └──┬────┬──┘
            │    │
     Success│    │Error
            │    │
            ▼    ▼
    ┌──────────┐ ┌──────────┐
    │COMPLETED │ │  FAILED  │
    └──────────┘ └──────────┘
```

## File Tracking Flow

```
┌─────────────────┐
│ Original File   │
│ /path/to/file   │
└────────┬────────┘
         │
         ▼
┌─────────────────────┐
│ Calculate SHA-256   │
└────────┬────────────┘
         │
         ▼
┌─────────────────────┐     ┌─────────────────┐
│ Check metadata.json │────►│ Hash matches?   │
└─────────────────────┘     └────┬───────┬────┘
                                 │ No    │ Yes
                                 ▼       │
                      ┌──────────────┐  │
                      │ Copy to      │  │
                      │ storage/     │  │
                      └──────┬───────┘  │
                             │          │
                             ▼          │
                      ┌──────────────┐  │
                      │ Update       │  │
                      │ metadata     │  │
                      └──────┬───────┘  │
                             │          │
                             ▼          ▼
                      ┌──────────────────┐
                      │ Return local path│
                      └──────────────────┘
```

## Template Processing Flow

```
┌──────────────────┐
│ Template File    │
│ (DOCX/XLSX/MSG)  │
└────────┬─────────┘
         │
         ▼
┌────────────────────┐
│ Extract            │
│ ##variables##      │
└────────┬───────────┘
         │
         ▼
┌────────────────────┐     ┌──────────────┐
│ Data Row           │────►│ ##name##     │
│ {name: "John",     │     │ ##email##    │
│  email: "j@x.com"} │     │ ##amount##   │
└────────┬───────────┘     └──────────────┘
         │                          │
         └──────────┬───────────────┘
                    │ Replace placeholders
                    ▼
         ┌──────────────────────┐
         │ Processed Document   │
         │ "Dear John"          │
         │ "Email: j@x.com"     │
         └──────────┬───────────┘
                    │
                    ▼
         ┌──────────────────────┐
         │ Save to job folder   │
         └──────────────────────┘
```

## Directory Structure at Runtime

```
autoarendt/
├── jobs/
│   ├── job-abc123/
│   │   ├── metadata.json          # Job info
│   │   ├── template.docx          # Local copy
│   │   ├── data.xlsx              # Local copy
│   │   ├── outputs/               # Generated files
│   │   │   ├── pdf/
│   │   │   │   ├── processed_1.pdf
│   │   │   │   └── processed_2.pdf
│   │   │   ├── word/
│   │   │   │   ├── processed_1.docx
│   │   │   │   └── processed_2.docx
│   │   │   └── processed_1.docx   # Intermediate
│   │   └── job_abc123_output.zip  # Final archive
│   └── job-def456/
│       └── ...
│
├── storage/
│   ├── 9a7b3c2d1e4f.docx         # Tracked template
│   ├── 5e6f7g8h9i0j.xlsx         # Tracked data
│   └── file_metadata.json         # Tracking info
│
└── uploads/
    ├── template_20251210_103045.docx
    └── data_20251210_103045.xlsx
```

## API Request/Response Flow

```
┌──────────┐
│  Client  │
└────┬─────┘
     │ POST /api/jobs
     │ FormData {
     │   template_file: File
     │   data_file: File
     │   output_formats: "pdf,word"
     │ }
     ▼
┌──────────────────┐
│  Flask Server    │
└────┬─────────────┘
     │ Validate request
     │ Save uploads
     │ Create job
     ▼
┌──────────────────┐
│  Response        │
│  {               │
│    success: true,│
│    job: {        │
│      id: "abc",  │
│      status: "pending"
│    }             │
│  }               │
└────┬─────────────┘
     │
     ▼
┌──────────┐
│  Client  │ Updates UI
└──────────┘
```

## Component Dependencies

```
run.py
  └─► app/__init__.py
       ├─► app/routes.py
       │    └─► services/job_manager.py
       │         ├─► services/file_tracker.py
       │         ├─► services/document_parser.py
       │         ├─► services/template_processor.py
       │         ├─► services/format_converter.py
       │         └─► models/job.py
       │
       └─► config/config.py

templates/index.html
  └─► static/js/app.js
       └─► API Endpoints (app/routes.py)
```

## Security Flow

```
┌──────────────┐
│ User Upload  │
└──────┬───────┘
       │
       ▼
┌──────────────────┐
│ Validate         │
│ - File extension │
│ - File size      │
│ - Mime type      │
└──────┬───────────┘
       │ Valid?
       ▼
┌──────────────────┐
│ Secure filename  │
│ - Remove special │
│ - Prevent path   │
│   traversal      │
└──────┬───────────┘
       │
       ▼
┌──────────────────┐
│ Save to uploads/ │
│ with UUID name   │
└──────────────────┘
```

## Legend

```
┌────────┐
│  Box   │  = Component/Process
└────────┘

    │
    ▼      = Data flow direction

   ┌──┐
───┤  ├─  = Decision point
   └──┘

  ═════   = Layer boundary
```
