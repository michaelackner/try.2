# VARO REBILLING - Excel Processor Web Application

A modern web application that processes multi-sheet Excel files to create formatted reports with enriched data. Features a beautiful oil-themed dashboard with Python backend for reliable Excel processing.

## Architecture

- **Frontend**: JavaScript + HTML/CSS with oil-themed dashboard design
- **Backend**: Python FastAPI service with openpyxl for Excel processing
- **Processing**: All 8 business rules and formatting handled server-side

## Features

- **File Upload**: Drag & drop or select Excel files (.xlsx/.xls)
- **Advanced Settings**: Configurable sheet names and column mappings
- **Two-Step Processing**:
  - **Step 1**: Creates formatted report with headers A-V and applies styling
  - **Step 2**: Enriches data using lookup tables and business rules
- **Dashboard Metrics**: Displays total deals, volume, and product distribution
- **Perfect Formatting**: Bold headers, yellow month cells, column borders, and proper alignment
- **Download**: Exports final Excel workbook with professional formatting

## Getting Started

### Prerequisites
- Node.js (v14 or higher)
- Python 3.8 or higher
- npm

### Installation
```bash
# Install frontend dependencies
npm install

# Install Python backend dependencies
pip install -r requirements.txt
```

### Running the Application

#### Option 1: Use the start script (recommended)
```bash
./start.sh
```

#### Option 2: Start servers manually
```bash
# Terminal 1: Start Python backend
python3 server.py

# Terminal 2: Start frontend
npm start
```

The application will be available at:
- **Dashboard**: `http://127.0.0.1:8080`
- **API Backend**: `http://127.0.0.1:8000`
- **API Documentation**: `http://127.0.0.1:8000/docs`

### Testing
A sample test file can be created using:
```bash
node create_test_excel.js
```

## Excel File Requirements

### Sheet Structure
The Excel file must contain 3 sheets:

1. **Sheet 1** (WHB/CIF & base data):
   - Column B: VSA deal
   - Column AA: VESSEL
   - Column M: Product
   - Column L: Hedge
   - Column Q: Qty BBL
   - Column AB: Inco
   - Column AD: Contractual Location
   - Column AL: Risk
   - Column X: Date
   - Column BZ: Must contain "WHB" for WHB+CIF deals
   - Column AB: Must contain "CIF" for WHB+CIF deals

2. **Sheet 2** (Costs):
   - Column N: Deal number
   - Column AQ: Cost type (BOT, BLC, CIN, CLI)
   - Column AV: Amount

3. **Sheet 3** (Hedge):
   - Column M: Hedge number
   - Column BR: VSA comments
   - Column CN: Additional information

### Advanced Settings

- **Output Sheet Name**: Name for the generated report sheet (default: "Q1-Q2-Q3-Q4-2024")
- **Raw Sheet Names**: Override default sheet names if needed
- **Deal Number Column**: Configure the column containing deal numbers (default: "N")

## Processing Rules

### Step 1: Format & Structure
- Creates headers A-V with proper formatting
- Maps raw data columns to new structure
- Sorts by date with month grouping
- Applies Excel formatting (borders, fonts, column widths)

### Step 2: Business Rules
1. **MIDLANDS Product**: Sets columns E-K to 0 and locks them
2. **WHB+CIF Deals**: Sets insurance columns I,J to 0 and locks them
3. **LC Costs**: Populates column E with BLC cost totals
4. **CIN Insurance**: Populates column I with CIN costs
5. **CLI Insurance**: Populates column J with CLI costs
6. **TOTAL Formula**: Column L = SUM(E:K)
7. **VSA Comments**: Populates column U from hedge lookup
8. **Additional Information**: Populates column V from hedge lookup

## Output Columns

| Col | Name | Description |
|-----|------|-------------|
| A | Varo deal | |
| B | VSA deal | Mapped from raw B |
| C | VESSEL | Mapped from raw AA |
| D | VMAG % | |
| E | L/C costs | BLC totals |
| F | Load insp | |
| G | Discharge inspection | |
| H | Superintendent | |
| I | CIN insurance | CIN costs |
| J | CLI insurance | CLI costs |
| K | Provisional charge | |
| L | TOTAL USD | =SUM(E:K) |
| M | VARO comments | |
| N | Product | Mapped from raw M |
| O | Hedge | Mapped from raw L |
| P | Qty BBL | Mapped from raw Q |
| Q | Inco | Mapped from raw AB |
| R | Contractual Location | Mapped from raw AD |
| S | Risk | Mapped from raw AL |
| T | Date | Mapped from raw X |
| U | VSA comments | From hedge lookup |
| V | Additional information | From hedge lookup |

## Performance Notes

- Uses efficient lookup tables for O(1) data matching
- Processes large files using background operations
- Memory-optimized for Excel files up to 50MB

## Error Handling

The application validates:
- File format (.xlsx/.xls)
- Required sheet presence
- Essential column availability
- Minimum data requirements

Clear error messages guide users when issues are encountered.

## Browser Support

- Chrome/Edge (recommended)
- Firefox
- Safari

Requires modern browser with support for:
- ES6+ JavaScript
- File API
- ArrayBuffer processing

## Deployment

### Railway (recommended)
- Uses `railway.json` with Nixpacks and `python server.py` start command.
- Steps:
  1) Install CLI: `npm i -g @railway/cli`
  2) Login: `railway login`
  3) Create project: `railway init` (choose empty project or existing)
  4) Deploy: `railway up`

Your service will expose the FastAPI app on the assigned domain. Health path: `/health`. API docs: `/docs`.

### Heroku
- Uses `Procfile` and `runtime.txt` (Python 3.11).
- Steps:
  1) `heroku login`
  2) `heroku create varo-rebilling` (or omit name for random)
  3) `git push heroku HEAD:main` (or the current branch)

Heroku sets `PORT` automatically; `server.py` reads it for Uvicorn.

### Docker
- Build: `docker build -t varo-rebilling .`
- Run: `docker run -p 8000:8000 varo-rebilling`

Open `http://localhost:8000` for the UI, `/docs` for API docs, `/health` for health.

## License

MIT License
