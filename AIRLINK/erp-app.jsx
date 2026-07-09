      const { useMemo, useState, useEffect, useRef } = React;

      function airlinkAsset(file) {
        const isLocal = location.hostname === "localhost" || location.hostname === "127.0.0.1" || location.protocol === "file:";
        if (!isLocal) return "/AIRLINK/" + file;
        const p = (location.pathname || "").replace(/\\/g, "/");
        if (p.includes("/AIRLINK/")) return file;
        return "AIRLINK/" + file;
      }

      const AIRLINK_BRAND = {
        logo: airlinkAsset("logo.png"),
        loginLogo: airlinkAsset("login-logo.png"),
        name: "AIRLINK"
      };

      const JOB_TYPES = ["Project", "Service", "Maintenance", "Trading(Buy)", "Trading(Sell)"];
      const ERP_MODULES = ["dashboard", "management_report", "overview", "ongoing", "monthly_report", "clients", "quotation", "job", "ar", "ap", "vendors", "settings"];
      const QUOTATION_STATUSES = ["Draft", "Sent", "Accepted", "Rejected", "Expired"];
      const JOB_STATUSES = ["Open", "In Progress", "Completed", "On Hold"];
      const PAYMENT_STAGES = ["", "Deposit", "Progress", "Balance", "Retention", "Final"];
      const ERP_REGIONS = [
        { id: "HK", labelEn: "HK", labelZh: "香港" },
        { id: "TW", labelEn: "TW", labelZh: "台灣" },
        { id: "SG", labelEn: "SG", labelZh: "新加坡" },
        { id: "MY", labelEn: "MY", labelZh: "馬來西亞" },
        { id: "ID", labelEn: "ID", labelZh: "印尼" },
        { id: "ALL", labelEn: "ALL", labelZh: "全部" }
      ];
      const DATA_REGION_IDS = ["HK", "TW", "SG", "MY", "ID"];

      const REGION_MAP_PINS = {
        HK: { lon: 114.17, lat: 22.32, labelEn: "Hong Kong", labelZh: "香港" },
        TW: { lon: 121.0, lat: 25.0, labelEn: "Taiwan", labelZh: "台灣" },
        SG: { lon: 103.82, lat: 1.35, labelEn: "Singapore", labelZh: "新加坡" },
        MY: { lon: 101.7, lat: 3.1, labelEn: "Malaysia", labelZh: "馬來西亞" },
        ID: { lon: 106.85, lat: -6.2, labelEn: "Indonesia", labelZh: "印尼" }
      };

      const REGION_COUNTRY_IDS = { HK: "156", TW: "158", MY: "458", ID: "360" };
      const REGION_FLAGS = { HK: "🇭🇰", TW: "🇹🇼", SG: "🇸🇬", MY: "🇲🇾", ID: "🇮🇩" };

      function worldMapView() {
        return (window.ERP_WORLD_MAP && window.ERP_WORLD_MAP.view) || { lonMin: -170, lonMax: 170, latMin: -50, latMax: 72, width: 960, height: 420 };
      }

      function mapPinXY(pin) {
        const v = worldMapView();
        return {
          x: ((pin.lon - v.lonMin) / (v.lonMax - v.lonMin)) * v.width,
          y: ((v.latMax - pin.lat) / (v.latMax - v.latMin)) * v.height
        };
      }

      const REGION_CURRENCY = { HK: "HKD", TW: "TWD", SG: "SGD", MY: "MYR", ID: "IDR" };

      function currencyForRegion(activeRegion, headerCurrency) {
        if (activeRegion && activeRegion !== "ALL" && REGION_CURRENCY[activeRegion]) {
          return REGION_CURRENCY[activeRegion];
        }
        return headerCurrency;
      }

      function amtInCurrencyLabel(ccy, t) {
        return (t("amtInCurrency") || "Amt in {ccy}").replace("{ccy}", ccy);
      }

      function regionLabel(id, lang) {
        const r = ERP_REGIONS.find((x) => x.id === id);
        if (!r) return id;
        return i18nLang(lang) === "zh_TW" ? r.labelZh : r.labelEn;
      }

      function ensureRegionOnRecords(arr, fallback = "HK") {
        return (arr || []).map((x) => ({ ...x, region: x.region || fallback }));
      }
      const WORLD_TIMEZONES = [
        { id: "UTC", labelEn: "UTC (World Standard)", labelZh: "UTC（世界標準時間）" },
        { id: "Asia/Hong_Kong", labelEn: "Hong Kong", labelZh: "香港" },
        { id: "Asia/Taipei", labelEn: "Taipei", labelZh: "台北" },
        { id: "Asia/Singapore", labelEn: "Singapore", labelZh: "新加坡" },
        { id: "Asia/Tokyo", labelEn: "Tokyo", labelZh: "東京" },
        { id: "Europe/London", labelEn: "London", labelZh: "倫敦" },
        { id: "America/New_York", labelEn: "New York", labelZh: "紐約" }
      ];

      function timezoneLabel(id, lang) {
        const tz = WORLD_TIMEZONES.find((x) => x.id === id);
        if (!tz) return id.replace(/^Asia\//, "").replace(/_/g, " ");
        return i18nLang(lang) === "zh_TW" ? tz.labelZh : tz.labelEn;
      }

      const PERM_KEYS = ["view", "add", "edit", "delete", "export", "import"];

      function fullPermissions() {
        return Object.fromEntries(ERP_MODULES.map((m) => [m, { view: true, add: true, edit: true, delete: true, export: true, import: true }]));
      }

      function emptyPermissions() {
        return Object.fromEntries(ERP_MODULES.map((m) => [m, { view: false, add: false, edit: false, delete: false, export: false, import: false }]));
      }

      function normalizePermissions(perms) {
        const out = emptyPermissions();
        ERP_MODULES.forEach((m) => {
          if (perms && perms[m]) PERM_KEYS.forEach((k) => { out[m][k] = !!perms[m][k]; });
        });
        if (!out.overview.view && perms && (perms.dashboard?.view || perms.ar?.view || perms.ap?.view)) {
          out.overview.view = true;
        }
        return out;
      }

      const EXCEL_SCHEMAS = {
        clients: {
          sheetName: "Clients",
          filePrefix: "clients",
          keyField: "customer_no",
          columns: [
            { header: "Customer No", field: "customer_no", required: true, hint: "Required, unique key" },
            { header: "COMPANY", field: "company", required: true, hint: "Required" },
            { header: "GST#", field: "gst_no", hint: "" },
            { header: "Primary Contact", field: "primary_contact", hint: "" },
            { header: "Company Phone #", field: "company_phone", hint: "" },
            { header: "Mobile Phone #", field: "mobile_phone", hint: "" },
            { header: "E-mail", field: "email", hint: "" },
            { header: "Address", field: "address", hint: "" },
            { header: "Postal Code", field: "postal_code", hint: "" },
            { header: "Account Dept Contact", field: "account_dept_contact", hint: "" },
            { header: "Payment Terms", field: "payment_terms", hint: "e.g. 30 Days" },
            { header: "Invoice Title", field: "invoice_title", hint: "Optional — shown on invoices/AR" },
            { header: "BU", field: "is_bu", hint: "Y if same company, different contact/job" },
            { header: "BU No", field: "bu_no", hint: "Business unit no. when BU is Y" }
          ]
        },
        job: {
          sheetName: "Jobs",
          filePrefix: "jobs",
          keyField: "job_no",
          columns: [
            { header: "Job No", field: "job_no", required: true, hint: "Required, unique key" },
            { header: "Client Company", field: "company", required: true, hint: "Must match existing client" },
            { header: "Type", field: "job_type", required: true, hint: "Project / Service / Maintenance / Trading(Buy) / Trading(Sell)" },
            { header: "Quotation No", field: "quotation_no", hint: "" },
            { header: "Amount", field: "amount", hint: "Numeric" },
            { header: "Currency", field: "currency", hint: "USD / HKD / SGD / TWD / RMB / AUD" },
            { header: "Customer PO", field: "customer_po", hint: "" },
            { header: "Customer PO Amount", field: "customer_po_amount", hint: "Numeric" },
            { header: "Status", field: "status", hint: "Open / In Progress / Completed / On Hold" },
            { header: "Start Date", field: "start_date", hint: "YYYY-MM-DD" },
            { header: "Description", field: "description", hint: "Project name" },
            { header: "Remarks", field: "remarks", hint: "" }
          ]
        },
        ar: {
          sheetName: "AR",
          filePrefix: "ar",
          keyField: "invoice_no",
          columns: [
            { header: "Invoice No", field: "invoice_no", required: true, hint: "Required, unique key" },
            { header: "Job No", field: "job_no", required: true, hint: "Must match existing job" },
            { header: "Customer", field: "customer", hint: "Optional if job exists" },
            { header: "Customer PO", field: "customer_po", hint: "" },
            { header: "Invoice Currency", field: "invoice_currency", hint: "USD / HKD / SGD / AUD" },
            { header: "Payment Stage", field: "payment_stage", hint: "Deposit / Progress / Balance / Retention / Final" },
            { header: "Invoice Amt", field: "invoice_amt", required: true, hint: "Numeric" },
            { header: "Invoice Date", field: "invoice_date", hint: "YYYY-MM-DD" },
            { header: "Due Date", field: "due_date", hint: "YYYY-MM-DD" },
            { header: "Payment Received Date", field: "payment_received_date", hint: "YYYY-MM-DD" },
            { header: "Payment Status", field: "payment_status", hint: "Awaiting Payment / Overdue / Paid / Cancel" }
          ]
        },
        ap: {
          sheetName: "AP",
          filePrefix: "ap",
          keyField: "invoice_no",
          columns: [
            { header: "Invoice No", field: "invoice_no", required: true, hint: "Required, unique key" },
            { header: "Job No", field: "job_no", required: true, hint: "Must match existing job" },
            { header: "Payee Type", field: "payee_type", hint: "Vendor" },
            { header: "Company Name", field: "company_name", required: true, hint: "Vendor or SI name" },
            { header: "Invoice Date", field: "invoice_date", hint: "YYYY-MM-DD" },
            { header: "Due Date", field: "due_date", hint: "YYYY-MM-DD" },
            { header: "Currency", field: "currency", hint: "USD / HKD / SGD / AUD" },
            { header: "Amount", field: "amount", required: true, hint: "Numeric" },
            { header: "Payment Stage", field: "payment_stage", hint: "Deposit / Progress / Balance / Retention / Final" },
            { header: "Pay Date", field: "pay_date", hint: "YYYY-MM-DD" },
            { header: "Remarks", field: "remarks", hint: "" }
          ]
        },
        vendors: {
          sheetName: "Vendors",
          filePrefix: "vendors",
          keyField: "vendor_no",
          columns: [
            { header: "Vendor No", field: "vendor_no", required: true, hint: "Required, unique key" },
            { header: "Name", field: "name", required: true, hint: "Required" },
            { header: "Email", field: "email", hint: "" },
            { header: "Phone", field: "phone", hint: "" },
            { header: "Address", field: "address", hint: "" },
            { header: "Bank", field: "bank", hint: "" },
            { header: "SWIFT", field: "swift_code", hint: "" },
            { header: "Charge", field: "charge", hint: "e.g. OUR / SHA / BEN" },
            { header: "Contact", field: "contact", hint: "" },
            { header: "Finance Contact", field: "finance_contact", hint: "" },
            { header: "Finance Email", field: "finance_email", hint: "" }
          ]
        },
        quotation: {
          sheetName: "Quotations",
          filePrefix: "quotations",
          keyField: "quotation_no",
          columns: [
            { header: "Quotation No", field: "quotation_no", required: true, hint: "Required, unique key" },
            { header: "Client Company", field: "company", required: true, hint: "Must match existing client" },
            { header: "Quotation Date", field: "quotation_date", hint: "YYYY-MM-DD" },
            { header: "Valid Until", field: "valid_until", hint: "YYYY-MM-DD" },
            { header: "Amount", field: "amount", required: true, hint: "Numeric" },
            { header: "Currency", field: "currency", hint: "USD / HKD / SGD / TWD / AUD" },
            { header: "Status", field: "status", hint: "Draft / Sent / Accepted / Rejected / Expired" },
            { header: "Job No", field: "job_no", hint: "Optional linked job" },
            { header: "Description", field: "description", hint: "" }
          ]
        }
      };

      function rowFromSchema(schema, record) {
        const row = {};
        schema.columns.forEach((c) => {
          const v = record[c.field];
          row[c.field] = v != null && v !== "" ? v : "";
        });
        return row;
      }

      function downloadWorkbook(wb, filename) {
        XLSX.writeFile(wb, filename);
      }

      function buildWorkbook(schema, dataRows, includeHints) {
        const headerRow = schema.columns.map((c) => c.header);
        const hintRow = schema.columns.map((c) => c.hint || "");
        const body = dataRows.map((r) => schema.columns.map((c) => (r[c.field] != null ? r[c.field] : "")));
        const aoa = includeHints ? [headerRow, hintRow, ...body] : [headerRow, ...body];
        const ws = XLSX.utils.aoa_to_sheet(aoa);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, schema.sheetName);
        return wb;
      }

      function excelCellToString(v) {
        if (v == null || v === "") return "";
        if (typeof v === "number") {
          if (v > 25000 && v < 65000 && typeof XLSX !== "undefined" && XLSX.SSF && XLSX.SSF.parse_date_code) {
            const d = XLSX.SSF.parse_date_code(v);
            if (d && d.y >= 1990 && d.y <= 2100) {
              return d.y + "-" + String(d.m).padStart(2, "0") + "-" + String(d.d).padStart(2, "0");
            }
          }
          return String(v);
        }
        return String(v).trim();
      }

      function normalizeImportHeader(h) {
        return String(h).trim().toLowerCase().replace(/[\s_#./\-]+/g, "");
      }

      function importHeaderMatches(header, col) {
        const a = normalizeImportHeader(header);
        if (!a) return false;
        const names = [col.header, col.field, ...(col.aliases || [])].map((x) => normalizeImportHeader(x));
        if (names.includes(a)) return true;
        if (col.field === "company" && (a === "companyname" || a === "clientcompany" || a === "公司名稱" || a === "公司" || a === "客戶名稱")) return true;
        if (col.field === "company_name" && (a === "companyname" || a === "vendorname" || a === "payee")) return true;
        if (col.field === "company_phone" && (a.includes("companyphone") || a === "officephone")) return true;
        if (col.field === "mobile_phone" && (a.includes("mobilephone") || a === "mobile")) return true;
        if (col.field === "customer_no" && (a === "clientno" || a === "customerno")) return true;
        if (col.field === "vendor_no" && a === "vendorno") return true;
        if (col.field === "si_no" && a === "sino") return true;
        if (col.field === "job_no" && a === "jobno") return true;
        if (col.field === "invoice_no" && a === "invoiceno") return true;
        if (col.field === "quotation_no" && a === "quotationno") return true;
        return false;
      }

      function parseExcelRows(arrayBuffer, schema) {
        const wb = XLSX.read(arrayBuffer, { type: "array" });
        const sheet = wb.Sheets[schema.sheetName] || wb.Sheets[wb.SheetNames[0]];
        if (!sheet) return { rows: [], errors: ["No worksheet found"] };
        const raw = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: "" });
        if (!raw.length) return { rows: [], errors: ["Empty file"] };

        function mapHeaders(headers) {
          const fieldByCol = {};
          const usedCols = new Set();
          schema.columns.forEach((col) => {
            const idx = headers.findIndex((h, i) => !usedCols.has(i) && importHeaderMatches(h, col));
            if (idx >= 0) {
              fieldByCol[idx] = col.field;
              usedCols.add(idx);
            }
          });
          return fieldByCol;
        }

        let headerRowIndex = 0;
        let fieldByCol = mapHeaders(raw[0].map((h) => String(h).trim()));
        let missing = schema.columns.filter((c) => c.required && !Object.values(fieldByCol).includes(c.field));
        if (missing.length) {
          for (let tryRow = 1; tryRow < Math.min(6, raw.length); tryRow++) {
            const candidate = mapHeaders((raw[tryRow] || []).map((h) => String(h).trim()));
            const miss = schema.columns.filter((c) => c.required && !Object.values(candidate).includes(c.field));
            if (miss.length < missing.length) {
              headerRowIndex = tryRow;
              fieldByCol = candidate;
              missing = miss;
              if (!missing.length) break;
            }
          }
        }
        if (missing.length) {
          return {
            rows: [],
            errors: [
              "Missing columns: " + missing.map((c) => c.header).join(", "),
              "Tip: use Export Template (.xlsx) — sheet \"" + schema.sheetName + "\". Keep row 1 headers unchanged."
            ]
          };
        }
        const rows = [];
        for (let i = headerRowIndex + 1; i < raw.length; i++) {
          const line = raw[i];
          if (!line || line.every((cell) => excelCellToString(cell) === "")) continue;
          const first = excelCellToString(line[0]);
          if (first.startsWith("(") || first.toLowerCase().includes("required") || first.toLowerCase().includes("hint")) continue;
          const rec = {};
          Object.entries(fieldByCol).forEach(([idx, field]) => {
            rec[field] = excelCellToString(line[Number(idx)]);
          });
          rows.push(rec);
        }
        return { rows, errors: [] };
      }

      function formatWorldTime(date, timeZone, lang) {
        const locale = i18nLang(lang) === "zh_TW" ? "zh-TW" : "en-GB";
        return new Intl.DateTimeFormat(locale, {
          timeZone,
          year: "numeric",
          month: "2-digit",
          day: "2-digit",
          hour: "2-digit",
          minute: "2-digit",
          second: "2-digit",
          hour12: false
        }).format(date);
      }

      function moduleLabel(module, t) {
        const map = {
          dashboard: t("nav_dashboard"),
          management_report: t("nav_management_report"),
          overview: t("nav_overview"),
          clients: t("nav_clients"),
          quotation: t("nav_quotation"),
          job: t("nav_job"),
          ongoing: t("nav_ongoing"),
          ar: t("nav_ar"),
          ap: t("nav_ap"),
          monthly_report: t("nav_monthly_report"),
          vendors: t("nav_vendors"),
          settings: t("nav_settings")
        };
        return map[module] || module;
      }

      function i18nLang(lang) {
        if (!lang) return "en";
        if (lang === "zh_TW" || lang === "zh" || lang.startsWith("zh")) return "zh_TW";
        return "en";
      }

      const I18N = {
        en: {
          cloudOnlyLoading: "Loading data from Supabase…",
          cloudOnlyHint: "Business data is stored in Supabase only (not browser cache). All PCs see the same data after login.",
          cloudOnlyMode: "Cloud database mode",
          importSyncRetry: "Retry sync",
          importStatusDismiss: "Dismiss",
          importSyncFail: "Import saved locally but cloud sync failed. Do not refresh until sync succeeds (check sync key / status bar).",
          importSyncOk: "Data synced to cloud — other users will see it shortly.",
          appTitle: "AIRLINK ERP", companyLogoSlot: "Company logo", companyNameSlot: "Company Name", nav_dashboard: "Dashboard", nav_management_report: "Management Report", nav_overview: "Overview", nav_clients: "Clients", nav_quotation: "Quotation", nav_job: "Job List", nav_ongoing: "Ongoing List", nav_ar: "AR List", nav_ap: "AP List", nav_monthly_report: "Monthly Report", nav_vendors: "Vendors", nav_si: "SI List", nav_settings: "Settings",
          baseCurrency: "Base Currency", hintClickRow: "Click a row or use View/Edit to manage records", newTransaction: "+ New Transaction",
          view: "View", edit: "Edit", save: "Save", cancel: "Cancel", all: "All", groupByClient: "Group by Client", groupByVendor: "Group by Vendor", displayMode: "Display", listView: "Flat list (all)", groupHint: "Filter Client shows only that client. Grouped view adds section headers while keeping all visible.", jobStatus: "Job Status",
          filterClient: "Filter Client", filterJob: "Filter Job", filterType: "Filter Type", filterClientHint: "Only show records for the selected client", filterJobHint: "Only show AP for the selected job",
          jobFlowHint: "Client → Job (quote) → Customer PO → Vendor quotes → Our PO → Delivery & invoices → Pay Vendor (AP) → Job Complete → Invoice Client (AR)",
          payeeType: "Payee Type", payeeVendor: "Vendor", payeeSi: "Legacy SI", apBills: "AP Bills",
          apPayeeAutoHint: "Auto-filled from Vendor master when you select payee. You can still edit for this bill.",
          markApPaid: "Mark Paid", apPaid: "Paid", apUnpaid: "Unpaid", issueInvoiceToClient: "Issue Invoice to Client",
          jobCompletedBanner: "Job completed — you can now issue an invoice to the client.", autoApPaidHint: "Entering pay date marks this bill as paid (green).",
          settingsTitle: "Settings", language: "Language", langEn: "English", langZhTw: "Traditional Chinese (Taiwan)", settingsSaved: "Settings saved",
          clientDetail: "Client Details", jobDetail: "Job Details", vendorDetail: "Vendor Details", siDetail: "SI Details", arDetail: "AR Details", apDetail: "AP Details",
          relatedJobs: "Related Jobs", relatedInvoices: "AR Invoices (to Client)", relatedPo: "Related PO", relatedAp: "AP Bills (Vendor costs)", noJobs: "No jobs", noInvoices: "No invoices", noPo: "No PO", noAp: "No AP bills",
          ownedClient: "Linked Client", linkedJob: "Linked Job", markPaid: "Mark Paid", paymentStatus: "Payment Status", paid: "Paid", awaitingPayment: "Awaiting Payment", overdue: "Overdue",
          addClient: "+ Add Client", addJob: "+ Add Job", addVendor: "+ Add Vendor", addSi: "+ Add SI", createAr: "+ Create AR Invoice", createAp: "+ Create AP Bill",
          vendorNo: "Vendor No.", siNo: "SI No.",
          actions: "Actions", invoices: "Invoices", clientGroup: "Client (Group)", type: "Type", status: "Status",
          quotationNo: "Quotation No.", amount: "Amount", currency: "Currency", description: "Description",
          addQuotation: "+ Add Quotation", addQuotationTitle: "Create Quotation", editQuotation: "Edit Quotation", quotationDetail: "Quotation Details",
          quotationDate: "Quotation Date", validUntil: "Valid Until", quotationStatus: "Status",
          ongoingHint: "Currency & PO Amt from AR invoices (summed per job). Click Edit to change Billed Amt or Remarks, then Save.",
          ongoingListTitle: "Ongoing Projects List",
          ongoingColJobNo: "Job No.", ongoingColNature: "Job Nature", ongoingColCustomer: "Customer",
          ongoingColProject: "Project Name", ongoingColPo: "Customer PO", ongoingColStart: "Start Date",
          ongoingColCurrency: "Currency", ongoingColPoAmt: "PO Amt", ongoingColBilled: "Billed Amt",
          ongoingColBalance: "PO Balance", ongoingColBalanceUsd: "Balance Amt in USD", ongoingColRemarks: "Remarks",
          ongoingNatureHint: "(Maintenance/Service/Project/Trade)",
          ongoingTotalPoUsd: "Total PO Amount in USD", ongoingTotalMaintUsd: "Total Amount for Maintenance in USD",
          ongoingTotalServiceUsd: "Total Amount for Service in USD", ongoingTotalProjectUsd: "Total Amount for Project in USD",
          ongoingTotalTradeUsd: "Total Amount for Trade in USD", exportPdf: "Export PDF",
          quotationSelectHint: "Type Accepted quotation no. and press Enter or blur to add. One job can link multiple quotations.",
          quotationNotAccepted: "Quotation not found or status is not Accepted.",
          quotationNotFound: "No matching quotation in this region.",
          quotationAlreadyLinked: "This quotation is already linked to job {job}.",
          quotationMultiJobHint: "Also linked to jobs: {jobs}",
          jobQuotationLinesTitle: "Linked Quotations",
          jobQuotationAdd: "+ Add Quotation",
          jobQuotationInputHint: "Enter Accepted quotation no., then blur or click Add.",
          quotationAmountTotal: "Quotation Amount (Total)",
          invoiceAmountTotal: "Invoice Amount (Total)",
          customerPoAmountTotal: "Customer PO Amount (Total)",
          jobCompleteCol: "Job Complete",
          arPaymentStatusAutoHint: "Status is automatic: Paid when payment date is entered; Overdue after due date; otherwise Awaiting Payment.",
          apPaymentStatusAutoHint: "Status is automatic: Paid when pay date is entered; Overdue after due date; otherwise Awaiting Payment.",
          jobPoLinesTitle: "Customer PO lines",
          jobPoAddLine: "+ Add PO",
          jobPoColNo: "PO No.",
          jobPoColAmount: "Amount",
          jobPoColDate: "Received Date",
          jobPoColRemarks: "Remarks",
          paymentStage: "Payment Stage",
          paymentStageHint: "Deposit / Progress / Balance — use multiple AR or AP invoices per job for installments.",
          duplicateInvoiceNo: "Invoice No. {no} already exists in {module}.",
          duplicateQuotationNo: "Quotation No. {no} already exists.",
          duplicateJobNo: "Job No. {no} already exists.",
          jobAlreadyHasQuotation: "This job is already linked to quotation {quotation}.",
          quotationLinkedOk: "Linked to Accepted quotation.",
          monthlyReportTitle: "Monthly Report",
          monthlyReportHint: "Full-year PO received, invoices issued, and AR expected (next-month due dates).",
          monthlyReportYear: "Report Year",
          monthlyPoReceived: "PO Received (Current Month)",
          monthlyInvoiceIssued: "Invoice Issued (Current Month)",
          monthlyArExpected: "AR Expected",
          monthlyArExpectedHint: "Expected receivables for the following month (e.g. July shows August due dates). Unpaid overdue items are included; Paid/Cancel excluded. Past months are frozen once the next month starts.",
          monthlyReportChartTitle: "{region} Monthly Report {year}",
          monthlyChartPo: "PO Received", monthlyChartInvoice: "Invoice Issued", monthlyChartArExpected: "AR Expected",
          monthlyTabReport: "Report", monthlyTabPoReceived: "PO Received", monthlyTabArInvoice: "AR Invoice",
          monthlyColJobNo: "Job#", monthlyColOtherCurrency: "Currency", monthlyColSourceAmount: "Amount", monthlyTotal: "TOTAL",
          monthlySyncFromJobs: "Sync from Jobs", monthlySyncFromAr: "Sync from AR",
          monthlySyncReport: "Sync Report",
          monthlySyncReportDone: "Report updated from Jobs and AR. PO / Invoice rows refreshed. Past months' AR Expected stay frozen.",
          monthlyReportSyncHint: "Report does not auto-update. After editing Jobs or AR, click Sync Report (or sync on PO Received / AR Invoice tabs).",
          monthlyPoDetailHint: "PO by month — Job#, currency, original PO amount (Customer PO Amount), then converted to {currency}. Click Sync from Jobs after editing jobs. Report tab uses the {currency} TOTAL.",
          monthlyArDetailHint: "AR by month — Job#, currency, original invoice amount, then converted to {currency}. Report tab uses the {currency} TOTAL.",
          monthlySyncSaved: "Synced and saved to this browser. Data remains after refresh (F5).",
          monthlyStorageHint: "Data is saved in your browser (localStorage), not a server database. Use Settings → Backup to save a copy. Import applies to the current region filter.",
          monthlyExportPdf: "Export PDF",
          allRegions: "All Regions",
          paymentCancel: "Cancel",
          colJobNo: "Job No.", colCustomer: "Customer", colCustomerPo: "Customer PO", colCustomerPoAmount: "Customer PO Amount",
          colInvoiceNo: "Invoice No.", colInvoiceDate: "Invoice Date", colInvoiceReceivedDate: "Invoice Received Date",
          colInvoiceCurrency: "Invoice Currency", colInvoiceAmt: "Invoice Amt", colDueDate: "Due Date", colPayDate: "Pay Date",
          colCompanyName: "Company Name", colPaymentTerms: "Payment Terms", colStartDate: "Start Date", colProjectName: "Project Name",
          colPayee: "Payee", colRemarks: "Remarks", colOverdueDays: "Overdue Days", colDays: "Days", colAmt: "Amt", colDate: "Date",
          colAccountDeptContact: "Account Dept Contact", colBankCharge: "Bank Charge", colSwiftCode: "SWIFT CODE",
          colPaymentAdviceEmail: "Payment Advice email", colCustomerNo: "Customer No.", colBu: "BU", colBuNo: "BU No.", clientBuHint: "Same company, different contact/job — enter BU No.", colFinanceContact: "Finance Contact", colFinanceEmail: "Finance Email",
          colCompany: "Company", colGstNo: "GST#", colPrimaryContact: "Primary Contact",
          colCompanyPhone: "Company Phone #", colMobilePhone: "Mobile Phone #", colEmail: "E-mail", colPostalCode: "Postal Code",
          colName: "Name", colBank: "Bank", colSwift: "SWIFT", colContact: "Contact",
          importReading: "Reading file…",
          importPreviewReady: "Import preview ready — please confirm.",
          storageSaveFailed: "Could not save to browser storage. Check privacy settings or disk space.",
          arOutstanding: "AR Outstanding", apOutstanding: "AP Outstanding", netPosition: "Net Position", dueSoon: "Due Soon", overdueCount: "overdue", pendingBills: "bills pending", dueSoonHint: "unpaid, due soon",
          topOverdue: "Top Overdue Invoices", approvalQueue: "Approval Queue", cashFlow: "Cash In vs Cash Out (90 Days)", arAging: "A/R Aging",
          dashSubtitle: "Executive summary — cash position & what needs attention",
          mgmtSubtitle: "Boss BI — profit, client margins, job-type mix & 3-month cash forecast",
          mgmtFilterClient: "Client", mgmtAllClients: "All Clients", mgmtGroupCurrency: "Group Base Currency", mgmtCompareCurrency: "Compare Currency",
          mgmtGrossRevenue: "Gross Revenue", mgmtTotalCogs: "Total COGS", mgmtNetProfit: "Net Profit", mgmtGrossMargin: "Gross Margin",
          mgmtRevenueHint: "AR from Completed jobs / Accepted quotations", mgmtCogsHint: "Vendor & SI AP on qualifying jobs",
          mgmtClientProfitShare: "This client accounts for {pct}% of company profit",
          mgmtAggregatedRegions: "Aggregated across {n} regions · FX to {currency}",
          mgmtTopClientsTitle: "Top Clients by Profit", mgmtTopClientsHint: "Bar length = net profit · click to drill down · margin % on the right",
          mgmtJobTypeTitle: "Profit by Job Type", mgmtJobTypeHint: "Share of net profit by Maintenance / Service / Project / Trade",
          mgmtCashForecastTitle: "3-Month Cash Flow Forecast", mgmtCashForecastHint: "Unpaid AR due dates vs unpaid AP due dates",
          mgmtCashIn: "Cash In", mgmtCashOut: "Cash Out", mgmtNetCashFlow: "Net Cash Flow",
          mgmtDryRunAlert: "Cash out exceeds cash in — prepare funding this month",
          mgmtClientRiskTitle: "Client Collection Risk", mgmtClientRiskHint: "Unpaid & overdue AR for this client over the next 3 months",
          mgmtOverdueAr: "Overdue AR", mgmtExpectedInMonth: "Expected in month",
          mgmtRegionBreakdownTitle: "Profit by Region & Client", mgmtLocalAmount: "Local", mgmtConvertedAmount: "Converted",
          mgmtMarginPct: "Margin", mgmtNatureProject: "Project", mgmtNatureService: "Service", mgmtNatureMaintenance: "Maintenance", mgmtNatureTrade: "Trade",
          mgmtNoQualifyingData: "No Completed / Accepted job data in this filter scope.",
          mgmtClearClient: "Back to all clients", mgmtShowingClient: "Drill-down: {name}",
          dashCashIn: "Cash In (AR)", dashCashOut: "Cash Out (AP)", dashCashInHint: "Unpaid client invoices", dashCashOutHint: "Unpaid vendor / SI bills", dashNetHint: "Receivables minus payables",
          dashActiveJobs: "Active Jobs", dashActiveJobsHint: "Open or in progress",
          dashIncomeOut: "Expected Income vs Out", dashIncomeOutHint: "Grouped by due date · full calendar year · unpaid only",
          dashIncomeOutH1: "Jan – Jun", dashIncomeOutH2: "Jul – Dec",
          dashCashInLegend: "Income (AR due)", dashCashOutLegend: "Out (AP due)",
          dashAllPaidHint: "Records exist but all AR/AP are marked paid for this region.",
          dashRegionHint: "AR/AP records exist in other regions. Try switching Region to ALL or the matching region.",
          dashArDueSoon: "AR Due Within 14 Days", dashApDueSoon: "AP Due Within 14 Days",
          dashOverdueAr: "Overdue AR — Collect", dashOverdueAp: "Overdue AP — Pay",
          dashDueToday: "Due today", dashDueInDays: "Due in {n}d", dashDaysOverdue: "{n}d overdue",
          dashNoDueSoon: "Nothing due in the next 14 days", dashNoOverdue: "No overdue items",
          dashAgingCurrent: "Not yet due", dashAging1_30: "1–30 days overdue", dashAging31_60: "31–60 days overdue", dashAging61_90: "61–90 days overdue", dashAging90plus: "90+ days overdue",
          dashThisMonthIn: "AR due this month", dashThisMonthOut: "AP due this month", dashViewAll: "View all", dashDueDate: "Due",
          dashGlobalMap: "Global Financial Map", dashGlobalMapHint: "World map · regional AR share below",
          dashMapArTotal: "AR Outstanding", dashMapActiveJobs: "Active Jobs", dashMapOverdueAlert: "Overdue AR — action needed",
          dashMapApTotal: "AP Outstanding", dashMapNetRegion: "Net (AR − AP)",
          dashMapZoomIn: "Zoom in", dashMapZoomOut: "Zoom out", dashMapZoomReset: "Reset", dashMapZoomAsia: "Asia-Pacific",
          dashMapExpandHint: "Click row to expand details",
          dashMapLegendHealthy: "Healthy", dashMapLegendCaution: "AP > AR", dashMapLegendAlert: "Overdue AR", dashMapLegendEmpty: "No activity",
          dashMapGoRegion: "Switch to this region",
          editClient: "Edit Client", addClientTitle: "Add Client", editJob: "Edit Job", addJobTitle: "Add Job", editVendor: "Edit Vendor", addVendorTitle: "Add Vendor", editSi: "Edit SI", addSiTitle: "Add SI",
          editAr: "Edit AR Invoice", createArTitle: "Create AR Invoice", editAp: "Edit AP Bill", createApTitle: "Create AP Bill", newTxTitle: "New Transaction",
          createArDesc: "Create accounts receivable invoice", createApDesc: "Create accounts payable bill",
          selectClientHint: "Select or type client no. / name",
          selectJobHint: "Select or type job no. / company",
          selectJobOptionalHint: "Select or type job no. / company (optional)",
          selectVendorHint: "Select or type vendor no. / name",
          noMatchFound: "No matching results",
          yes: "Yes", no: "No", goToAr: "Go to AR", goToJob: "Go to Job", goToAp: "Go to AP",
          paymentReceived: "Payment Received Date", autoPaidHint: "Entering payment date will set status to Paid",
          arDueAutoHint: "Due date is calculated from Invoice Date + client Payment Terms (editable).",
          overviewTitle: "AR / AP Overview", overviewHint: "Filter by invoice date range · amounts in base currency",
          overviewStartDate: "Start Date", overviewEndDate: "End Date", overviewApply: "Apply",
          overviewInTotal: "Cash In (AR)", overviewOutTotal: "Cash Out (AP)", overviewNet: "Net",
          overviewNoData: "No AR/AP invoices in this date range", overviewByInvoiceDate: "Invoice date within range",
          overviewArCount: "AR invoices", overviewApCount: "AP bills", phoneNo: "Phone", address: "Address",
          overviewPaidTitle: "Collected vs Paid — Full Year", overviewPaidHint: "Grouped by payment date · current calendar year",
          overviewArCollected: "AR Collected", overviewApPaid: "AP Paid", overviewArCollectedTotal: "Total AR Collected", overviewApPaidTotal: "Total AP Paid",
          amtInCurrency: "Amt in {ccy}",
          amountManualOverride: "editable", settingsLangHint: "The interface language updates immediately when you switch.",
          settingsGuide: "User Guide", openUserGuide: "Open User Guide", userGuideHint: "Recommended workflow order to avoid broken links or failed imports.",
          invoiceAmountOrig: "Amount (invoice currency)", invoiceAmountHint: "Original amount on vendor/client invoice. Editable — updates base amount when rate changes.",
          baseAmountHint: "Usually Amount × Exchange Rate. Override here if bank/rounding differs.",
          siPageHint: "SI completes work and sends PO/Invoice to you — pay them via AP List.",
          settingsGeneral: "General", settingsManagement: "User Management", settingsAudit: "Audit Log",
          worldTime: "World Time", worldTimeHint: "Uses selected timezone (not your computer clock).",
          timezone: "Timezone", currentUser: "Signed in as", switchUser: "Switch user (demo)",
          mgmtTitle: "User & Access Management", mgmtHint: "Root admin can create users and assign view / add / edit / delete rights per module.",
          addUser: "+ Add User", editUser: "Edit User", userName: "Display Name", userEmail: "Email", userPassword: "Password", userRole: "Role", roleRoot: "Root (full access)", roleUser: "User",
          userActive: "Active", accessRights: "Access Rights", permView: "View", permAdd: "Add", permEdit: "Edit", permDelete: "Delete",
          auditTitle: "Activity Log", auditHint: "Filter by date, user, module or action — like a report.",
          auditWhen: "When (UTC stored)", auditUser: "User", auditModule: "Module", auditAction: "Action", auditEntity: "Record", auditDetail: "Details",
          auditDateFrom: "From date", auditDateTo: "To date", auditSearch: "Search", auditReset: "Reset filters", auditNoRows: "No log entries match your filters.",
          actionCreate: "Create", actionUpdate: "Update", actionDelete: "Delete", actionMarkPaid: "Mark Paid", actionUserCreate: "User Created", actionUserUpdate: "User Updated",
          confirmDelete: "Delete this record?", noPermission: "You do not have permission for this action.", rootOnly: "Root admin only.",
          delete: "Delete", allUsers: "All users", allModules: "All modules", allActions: "All actions",
          permExport: "Export", permImport: "Import", exportData: "Export Data", exportExcel: "Export Excel", exportTemplate: "Download Template",
          importExcel: "Import Excel", importPreviewTitle: "Import Preview", importConfirm: "Confirm Import",
          importRowCreate: "New", importRowUpdate: "Update", importRowError: "Error", importSummary: "Summary",
          importNoRows: "No valid rows found.", importFileError: "Could not read file.",
          actionExport: "Export", actionImport: "Import",           importModeHint: "Existing keys will be updated; new keys will be created.",
          filterStatus: "Filter Status", searchPlaceholder: "Search...", searchHint: "Matches invoice no., client, job no., PO, etc.",
          searchHintClients: "Matches customer no., company, contact, email, phone, address, etc.",
          searchHintQuotation: "Matches quotation no., client, job no., description, status, etc.",
          searchHintJob: "Matches job no., client, PO, quotation no., description, status, etc.",
          searchHintVendors: "Matches vendor no., name, email, bank, SWIFT, contact, phone, etc.",
          searchHintSi: "Matches SI no., name, email, bank, SWIFT, contact, phone, etc.",
          region: "Region", regionAccess: "Region Access", regionScopeHint: "Data you create is tagged to the active region. ALL shows every region you can access.",
          regionDataHint: "New records are saved under the currently selected region (not ALL).",
          importRegionHint: "Imported under region: {region}. Switch region (top-right) or set ALL to see records.",
          settingsBackup: "Backup & Restore", backupTitle: "Backup & Restore", backupHint: "All lists are stored in this browser only (localStorage). Download a backup file to move data to another PC or browser. Restore replaces all current data.",
          backupDownload: "Download Backup", restoreUpload: "Restore from Backup", restoreConfirm: "Restore will REPLACE all current ERP data in this browser. Continue?",
          restoreDone: "Restore completed. Page will reload.", restoreInvalid: "Invalid backup file.", backupDone: "Backup file downloaded.",
          cloudSyncTitle: "Cloud Database (Supabase)", cloudSyncHint: "Push/Pull uses your server API. Set the same sync key in Vercel → ERP_SYNC_SECRET.",
          cloudSyncKey: "Cloud Sync Key", cloudSyncKeyPlaceholder: "Same as ERP_SYNC_SECRET on server",
          cloudPush: "Push to Cloud", cloudPull: "Pull from Cloud", cloudPushDone: "Uploaded to Supabase.", cloudPullDone: "Downloaded from cloud. Reloading…",
          cloudSyncFail: "Cloud sync failed", cloudNoBackup: "No backup found in database.", cloudNotConfigured: "Server cloud sync not configured.",
          liveSyncTitle: "Live Multi-User Sync", liveSyncHint: "Each record syncs to Supabase automatically. All users see changes within a few seconds. Requires sync key + SQL migration 002.",
          liveSyncEnable: "Enable live sync (auto push & pull)", liveSyncOk: "Live sync active", liveSyncReadOnly: "Live sync active (read-only)", liveSyncSyncing: "Syncing…", liveSyncError: "Sync error",
          liveSyncLast: "Last sync", liveSyncNeedKey: "Enter cloud sync key first.", liveSyncBootstrapped: "Live sync enabled. Data connected to cloud.",
          liveSyncAutoHint: "Sync starts automatically when you open the app. View-only users download only; editors upload changes.",
          loginUsername: "Username", loginPassword: "Password", loginRemember: "Remember me", loginBtn: "Login", loginFail: "Invalid username or password.", loginInactive: "This account is disabled.", logout: "Logout",
          myAccountTitle: "My Account", changePassword: "Change Password", currentPassword: "Current password", newPassword: "New password", confirmPassword: "Confirm new password", passwordChanged: "Password updated.", passwordMismatch: "New passwords do not match.", passwordWrong: "Current password is incorrect.", userLoginId: "Username",
          noDatabaseNote: "No server database yet — Excel Import updates browser storage only. After upload, click Confirm Import in the preview dialog."
        },
        zh_TW: {
          cloudOnlyLoading: "正在從 Supabase 載入資料…",
          cloudOnlyHint: "業務資料只儲存在 Supabase（不使用瀏覽器快取）。所有電腦登入後會看到相同資料。",
          cloudOnlyMode: "雲端資料庫模式",
          importSyncRetry: "重試同步",
          importStatusDismiss: "關閉",
          importSyncFail: "匯入已儲存於本機，但雲端同步失敗。請勿重新整理，先確認同步金鑰及狀態列。",
          importSyncOk: "已同步至雲端，其他用戶稍後會看到。",
          appTitle: "AIRLINK ERP", companyLogoSlot: "公司標誌", companyNameSlot: "公司名稱", nav_dashboard: "儀表板", nav_management_report: "管理層報表", nav_overview: "總覽", nav_clients: "客戶", nav_quotation: "報價單", nav_job: "工作清單", nav_ongoing: "進行中清單", nav_ar: "應收清單", nav_ap: "應付清單", nav_monthly_report: "月報表", nav_vendors: "供應商", nav_si: "分包商清單", nav_settings: "設定",
          baseCurrency: "本位幣", hintClickRow: "點擊列或按「查看／修改」管理資料", newTransaction: "+ 新增交易",
          view: "查看", edit: "修改", save: "儲存", cancel: "取消", all: "全部", groupByClient: "按客戶分組", groupByVendor: "按供應商分組", displayMode: "顯示方式", listView: "列表（全部）", groupHint: "篩選客戶只顯示該客戶；分組模式會加標題分區但仍可顯示全部。", jobStatus: "工作狀態",
          filterClient: "篩選客戶", filterJob: "篩選工作", filterType: "篩選類型", filterClientHint: "只顯示所選客戶的資料", filterJobHint: "只顯示所選工作的應付單",
          jobFlowHint: "客戶 → 工作（報價）→ 客戶 PO → 向供應商詢價 → 開立 PO → 收貨及發票 → 付供應商（應付）→ 工作完成 → 向客戶開立發票（應收）",
          payeeType: "付款對象", payeeVendor: "供應商", payeeSi: "舊分包商 (SI)", apBills: "應付單數",
          apPayeeAutoHint: "選擇供應商後會自動帶入主檔資料，仍可為此單修改。",
          markApPaid: "標記已付款", apPaid: "已付款", apUnpaid: "待付款", issueInvoiceToClient: "向客戶開立發票",
          jobCompletedBanner: "工作已完成 — 可向客戶開立發票收款。", autoApPaidHint: "輸入付款日期會標記為已付款（綠色顯示）。",
          settingsTitle: "設定", language: "介面語言", langEn: "英文", langZhTw: "繁體中文（台灣）", settingsSaved: "設定已儲存",
          clientDetail: "客戶詳情", jobDetail: "工作詳情", vendorDetail: "供應商詳情", siDetail: "分包商詳情", arDetail: "應收詳情", apDetail: "應付詳情",
          relatedJobs: "相關工作", relatedInvoices: "應收發票（向客戶）", relatedPo: "相關採購單", relatedAp: "應付單（供應商成本）", noJobs: "暫無工作", noInvoices: "暫無發票", noPo: "暫無 PO", noAp: "暫無應付單",
          ownedClient: "所屬客戶", linkedJob: "關聯工作", markPaid: "標記已收款", paymentStatus: "付款狀態", paid: "已收款", awaitingPayment: "待收款", overdue: "逾期",
          addClient: "+ 新增客戶", addJob: "+ 新增工作", addVendor: "+ 新增供應商", addSi: "+ 新增分包商", createAr: "+ 新增應收發票", createAp: "+ 新增應付單",
          vendorNo: "供應商編號", siNo: "分包商編號",
          actions: "操作", invoices: "發票數", clientGroup: "客戶（分組）", type: "類型", status: "狀態",
          quotationNo: "報價單號", amount: "金額", currency: "幣別", description: "說明",
          addQuotation: "+ 新增報價", addQuotationTitle: "新增報價單", editQuotation: "修改報價單", quotationDetail: "報價詳情",
          quotationDate: "報價日期", validUntil: "有效期至", quotationStatus: "狀態",
          ongoingHint: "幣別與 PO 金額來自應收發票（同一工作會加總）。點擊「修改」可變更已開票金額或備註，儲存後鎖定。",
          ongoingListTitle: "進行中項目清單",
          ongoingColJobNo: "工作編號", ongoingColNature: "工作性質", ongoingColCustomer: "客戶",
          ongoingColProject: "專案名稱", ongoingColPo: "客戶 PO", ongoingColStart: "開始日期",
          ongoingColCurrency: "幣別", ongoingColPoAmt: "PO 金額", ongoingColBilled: "已開票金額",
          ongoingColBalance: "PO 餘額", ongoingColBalanceUsd: "餘額（USD）", ongoingColRemarks: "備註",
          ongoingNatureHint: "（維護／服務／專案／貿易）",
          ongoingTotalPoUsd: "PO 總額（USD）", ongoingTotalMaintUsd: "維護類總額（USD）",
          ongoingTotalServiceUsd: "服務類總額（USD）", ongoingTotalProjectUsd: "專案類總額（USD）",
          ongoingTotalTradeUsd: "貿易類總額（USD）", exportPdf: "匯出 PDF",
          quotationSelectHint: "輸入狀態為 Accepted 的報價單號，離開欄位或按新增。一個工作可連結多張報價單。",
          quotationNotAccepted: "找不到報價單，或狀態不是 Accepted。",
          quotationNotFound: "此區域沒有這張報價單。",
          quotationAlreadyLinked: "這張報價單已連結至工作 {job}。",
          quotationMultiJobHint: "亦已連結工作：{jobs}",
          jobQuotationLinesTitle: "已連結報價單",
          jobQuotationAdd: "+ 新增報價單",
          jobQuotationInputHint: "輸入 Accepted 報價單號，再離開欄位或按新增。",
          quotationAmountTotal: "報價金額（總計）",
          invoiceAmountTotal: "發票金額（總計）",
          customerPoAmountTotal: "客戶 PO 金額（總計）",
          jobCompleteCol: "工作完成",
          arPaymentStatusAutoHint: "狀態自動更新：有收款日期 = 已收款；過 due date = 逾期；否則 = 待收款。",
          apPaymentStatusAutoHint: "狀態自動更新：有付款日期 = 已付款；過 due date = 逾期；否則 = 待付款。",
          jobPoLinesTitle: "客戶 PO 明細",
          jobPoAddLine: "+ 新增 PO",
          jobPoColNo: "PO 編號",
          jobPoColAmount: "金額",
          jobPoColDate: "收單日期",
          jobPoColRemarks: "備註",
          paymentStage: "付款階段",
          paymentStageHint: "Deposit / Progress / Balance — 同一 Job 可開多張 AR 或 AP 分期單。",
          duplicateInvoiceNo: "Invoice 編號 {no} 已存在於 {module}。",
          duplicateQuotationNo: "報價單號 {no} 已存在。",
          duplicateJobNo: "工作編號 {no} 已存在。",
          jobAlreadyHasQuotation: "這個工作已連結報價單 {quotation}。",
          quotationLinkedOk: "已對應 Accepted 報價單。",
          monthlyReportTitle: "月報表",
          monthlyReportHint: "全年各月 PO 收單、開票及預計應收（AR Expected = 下月到期日）。",
          monthlyReportYear: "報表年份",
          monthlyPoReceived: "本月 PO 收單",
          monthlyInvoiceIssued: "本月開立發票",
          monthlyArExpected: "預計應收",
          monthlyArExpectedHint: "預計下月應收（例如 7 月欄顯示 8 月到期日）。未收的逾期項目會計入；已收款／已取消不計。月份結束後凍結，之後收款不會改變歷史月份。",
          monthlyReportChartTitle: "{region} 月報表 {year}",
          monthlyChartPo: "PO 收單", monthlyChartInvoice: "開立發票", monthlyChartArExpected: "預計應收",
          monthlyTabReport: "報表", monthlyTabPoReceived: "PO 收單", monthlyTabArInvoice: "AR 發票",
          monthlyColJobNo: "工作編號", monthlyColOtherCurrency: "幣別", monthlyColSourceAmount: "原幣金額", monthlyTotal: "合計",
          monthlySyncFromJobs: "從工作同步", monthlySyncFromAr: "從應收同步",
          monthlySyncReport: "同步月報",
          monthlySyncReportDone: "已從工作及應收同步月報，PO／發票列已更新。已結束月份的預計應收維持凍結。",
          monthlyReportSyncHint: "報表不會自動更新。修改工作或應收後，請點擊「同步月報」（或在 PO 收單／AR 發票頁籤分別同步）。",
          monthlyPoDetailHint: "按月 PO — 工作編號、幣別、原幣 PO 金額（客戶 PO 金額）、換算成 {currency}。修改工作後請點擊「從工作同步」。報表頁使用 {currency} 合計。",
          monthlyArDetailHint: "按月 AR — 工作編號、幣別、原幣發票金額、換算成 {currency}。報表頁使用 {currency} 合計。",
          monthlySyncSaved: "已同步並儲存至本機瀏覽器，重新整理（F5）後仍會保留。",
          monthlyStorageHint: "資料儲存在瀏覽器（localStorage），並非伺服器資料庫。請至「設定 → 備份」下載備份。匯入會套用目前右上角所選地區。",
          monthlyExportPdf: "匯出 PDF",
          allRegions: "全部地區",
          paymentCancel: "已取消",
          colJobNo: "工作編號", colCustomer: "客戶", colCustomerPo: "客戶 PO", colCustomerPoAmount: "客戶 PO 金額",
          colInvoiceNo: "發票編號", colInvoiceDate: "發票日期", colInvoiceReceivedDate: "發票收到日期",
          colInvoiceCurrency: "發票幣別", colInvoiceAmt: "發票金額", colDueDate: "到期日", colPayDate: "付款日期",
          colCompanyName: "公司名稱", colPaymentTerms: "付款條件", colStartDate: "開始日期", colProjectName: "專案名稱",
          colPayee: "付款對象", colRemarks: "備註", colOverdueDays: "逾期天數", colDays: "天數", colAmt: "金額", colDate: "日期",
          colAccountDeptContact: "會計聯絡人", colBankCharge: "銀行手續費", colSwiftCode: "SWIFT 代碼",
          colPaymentAdviceEmail: "付款通知電子郵件", colCustomerNo: "客戶編號", colBu: "BU", colBuNo: "BU 編號", clientBuHint: "同一公司、不同聯絡人或不同工作時可勾選，並填寫 BU 編號。", colFinanceContact: "財務聯絡人", colFinanceEmail: "財務電子郵件",
          colCompany: "公司名稱", colGstNo: "GST／統一編號", colPrimaryContact: "主要聯絡人",
          colCompanyPhone: "公司電話", colMobilePhone: "手機", colEmail: "電子郵件", colPostalCode: "郵遞區號",
          colName: "名稱", colBank: "銀行", colSwift: "SWIFT", colContact: "聯絡人",
          importReading: "正在讀取檔案…",
          importPreviewReady: "匯入預覽已準備好，請確認。",
          storageSaveFailed: "無法儲存到瀏覽器，請檢查隱私權設定或儲存空間。",
          arOutstanding: "應收未結", apOutstanding: "應付未結", netPosition: "淨部位", dueSoon: "即將到期", overdueCount: "筆已逾期", pendingBills: "筆待處理", dueSoonHint: "未收款、即將到期",
          topOverdue: "逾期發票排行", approvalQueue: "審批佇列", cashFlow: "現金流入 vs 流出（90 天）", arAging: "應收帳齡",
          dashSubtitle: "管理層總覽 — 現金狀況及需要跟進的項目",
          mgmtSubtitle: "管理層 BI — 毛利、客戶排行、工作類型利潤與 3 個月現金流預測",
          mgmtFilterClient: "客戶", mgmtAllClients: "全部客戶", mgmtGroupCurrency: "集團本位幣", mgmtCompareCurrency: "對照貨幣",
          mgmtGrossRevenue: "總收入", mgmtTotalCogs: "項目總成本", mgmtNetProfit: "淨利潤", mgmtGrossMargin: "毛利率",
          mgmtRevenueHint: "已完成工作／Accepted 報價相關應收", mgmtCogsHint: "相關工作的供應商／分包商應付",
          mgmtClientProfitShare: "此客戶佔公司利潤 {pct}%",
          mgmtAggregatedRegions: "已匯總 {n} 個地區 · 折算為 {currency}",
          mgmtTopClientsTitle: "客戶利潤排行榜", mgmtTopClientsHint: "條形長度 = 純利潤 · 點擊穿透分析 · 右側括號為毛利率",
          mgmtJobTypeTitle: "按工作類型分類利潤", mgmtJobTypeHint: "維護／服務／專案／貿易各佔淨利潤比例",
          mgmtCashForecastTitle: "3 個月現金流預測", mgmtCashForecastHint: "未收應收到期日 vs 未付應付到期日",
          mgmtCashIn: "預計流入", mgmtCashOut: "預計流出", mgmtNetCashFlow: "淨現金流",
          mgmtDryRunAlert: "流出大於流入 — 本月需準備資金，避免資金斷裂",
          mgmtClientRiskTitle: "客戶收款風險", mgmtClientRiskHint: "此客戶未來 3 個月預計欠款與逾期應收",
          mgmtOverdueAr: "逾期應收", mgmtExpectedInMonth: "當月預計收款",
          mgmtRegionBreakdownTitle: "地區與客戶利潤明細", mgmtLocalAmount: "當地幣", mgmtConvertedAmount: "折算",
          mgmtMarginPct: "毛利率", mgmtNatureProject: "專案", mgmtNatureService: "服務", mgmtNatureMaintenance: "維護", mgmtNatureTrade: "貿易",
          mgmtNoQualifyingData: "此篩選範圍內沒有已完成／Accepted 的工作資料。",
          mgmtClearClient: "返回全部客戶", mgmtShowingClient: "穿透分析：{name}",
          dashCashIn: "現金流入（應收）", dashCashOut: "現金流出（應付）", dashCashInHint: "客戶未收款發票", dashCashOutHint: "供應商／分包商未付款", dashNetHint: "應收減應付",
          dashActiveJobs: "進行中工作", dashActiveJobsHint: "開啟或進行中",
          dashIncomeOut: "預期收入 vs 支出", dashIncomeOutHint: "按到期日 · 全年（1–12月）· 只計未結清",
          dashIncomeOutH1: "1 – 6 月", dashIncomeOutH2: "7 – 12 月",
          dashCashInLegend: "收入（應收到期）", dashCashOutLegend: "支出（應付到期）",
          dashAllPaidHint: "此區域有記錄，但應收／應付已全部標記為已付款。",
          dashRegionHint: "應收／應付資料在其他區域。請將「區域」改為 ALL 或對應區域。",
          dashArDueSoon: "14 日內應收到期", dashApDueSoon: "14 日內應付到期",
          dashOverdueAr: "逾期應收 — 催收", dashOverdueAp: "逾期應付 — 付款",
          dashDueToday: "今日到期", dashDueInDays: "{n} 日後到期", dashDaysOverdue: "逾期 {n} 日",
          dashNoDueSoon: "14 日內無到期項目", dashNoOverdue: "無逾期項目",
          dashAgingCurrent: "未到期", dashAging1_30: "逾期 1–30 日", dashAging31_60: "逾期 31–60 日", dashAging61_90: "逾期 61–90 日", dashAging90plus: "逾期 90 日以上",
          dashThisMonthIn: "本月應收到期", dashThisMonthOut: "本月應付到期", dashViewAll: "查看全部", dashDueDate: "到期",
          dashGlobalMap: "跨地區財務世界地圖", dashGlobalMapHint: "全球地圖 · 下方為各地區應收佔比",
          dashMapArTotal: "應收未結", dashMapActiveJobs: "進行中工作", dashMapOverdueAlert: "逾期應收 — 需跟進",
          dashMapApTotal: "應付未結", dashMapNetRegion: "淨額（應收 − 應付）",
          dashMapZoomIn: "放大", dashMapZoomOut: "縮小", dashMapZoomReset: "重設", dashMapZoomAsia: "亞太",
          dashMapExpandHint: "點擊列展開詳情",
          dashMapLegendHealthy: "健康", dashMapLegendCaution: "應付 > 應收", dashMapLegendAlert: "逾期應收", dashMapLegendEmpty: "無活動",
          dashMapGoRegion: "切換至此地區",
          editClient: "修改客戶", addClientTitle: "新增客戶", editJob: "修改工作", addJobTitle: "新增工作", editVendor: "修改供應商", addVendorTitle: "新增供應商", editSi: "修改分包商", addSiTitle: "新增分包商",
          editAr: "修改應收發票", createArTitle: "新增應收發票", editAp: "修改應付單", createApTitle: "新增應付單", newTxTitle: "新增交易",
          createArDesc: "建立應收發票", createApDesc: "建立應付單",
          selectClientHint: "可選擇或輸入客戶編號／名稱",
          selectJobHint: "可選擇或輸入工作編號／公司名稱",
          selectJobOptionalHint: "可選擇或輸入工作編號／公司名稱（非必填）",
          selectVendorHint: "可選擇或輸入供應商編號／名稱",
          noMatchFound: "找不到符合的結果",
          yes: "是", no: "否", goToAr: "前往應收", goToJob: "前往工作", goToAp: "前往應付",
          paymentReceived: "收款日期", autoPaidHint: "輸入收款日期會自動設為已收款",
          arDueAutoHint: "到期日依發票日期 + 客戶付款條款自動計算（可手動修改）。",
          overviewTitle: "應收／應付總覽", overviewHint: "按發票日期篩選 · 金額以本位幣顯示",
          overviewStartDate: "開始日期", overviewEndDate: "結束日期", overviewApply: "套用",
          overviewInTotal: "流入（應收）", overviewOutTotal: "流出（應付）", overviewNet: "淨額",
          overviewNoData: "此日期範圍內沒有應收／應付發票", overviewByInvoiceDate: "發票日期在範圍內",
          overviewArCount: "應收發票", overviewApCount: "應付帳單", phoneNo: "電話", address: "地址",
          overviewPaidTitle: "全年已收／已付", overviewPaidHint: "按收款／付款日期 · 本年度各月",
          overviewArCollected: "應收已收", overviewApPaid: "應付已付", overviewArCollectedTotal: "全年應收已收", overviewApPaidTotal: "全年應付已付",
          amtInCurrency: "折合 {ccy}",
          amountManualOverride: "可手動修改", settingsLangHint: "切換後介面會即時更新為繁體中文（台灣）。",
          settingsGuide: "使用教學", openUserGuide: "開啟使用教學", userGuideHint: "說明正確操作順序，避免因步驟顛倒造成資料無法連結或匯入失敗。",
          invoiceAmountOrig: "金額（發票原幣）",           invoiceAmountHint: "供應商／客戶發票上的原幣金額，可以修改；修改後會按匯率重算本位幣。",
          baseAmountHint: "一般 = 金額 × 匯率。若銀行實收、四捨五入不同，可在此手動修改。",
          siPageHint: "分包商 (SI) 完成工作後向你開立 PO／發票，你在應付清單付款給對方。",
          settingsGeneral: "一般", settingsManagement: "用戶管理", settingsAudit: "操作記錄",
          worldTime: "世界時間", worldTimeHint: "依所選時區顯示（不跟電腦本機時間）。",
          timezone: "時區", currentUser: "目前登入", switchUser: "切換用戶（示範）",
          mgmtTitle: "用戶與權限管理", mgmtHint: "Root 管理員可新增用戶，並按模組設定檢視／新增／修改／刪除權限。",
          addUser: "+ 新增用戶", editUser: "修改用戶", userName: "顯示名稱", userEmail: "電子郵件", userPassword: "密碼", userRole: "角色", roleRoot: "Root（最高權限）", roleUser: "一般用戶",
          userActive: "啟用", accessRights: "存取權限", permView: "檢視", permAdd: "新增", permEdit: "修改", permDelete: "刪除",
          auditTitle: "操作記錄", auditHint: "可按日期、用戶、模組、動作篩選，如同報表查詢。",
          auditWhen: "時間（UTC 儲存）", auditUser: "用戶", auditModule: "模組", auditAction: "動作", auditEntity: "記錄", auditDetail: "詳情",
          auditDateFrom: "開始日期", auditDateTo: "結束日期", auditSearch: "搜尋", auditReset: "重設篩選", auditNoRows: "沒有符合條件的記錄。",
          actionCreate: "新增", actionUpdate: "修改", actionDelete: "刪除", actionMarkPaid: "標記已付款", actionUserCreate: "建立用戶", actionUserUpdate: "更新用戶",
          confirmDelete: "確定刪除此記錄？", noPermission: "你沒有此操作權限。", rootOnly: "僅 Root 管理員可用。",
          delete: "刪除", allUsers: "全部用戶", allModules: "全部模組", allActions: "全部動作",
          permExport: "匯出", permImport: "匯入", exportData: "匯出資料", exportExcel: "匯出 Excel", exportTemplate: "下載範本",
          importExcel: "匯入 Excel", importPreviewTitle: "匯入預覽", importConfirm: "確認匯入",
          importRowCreate: "新增", importRowUpdate: "更新", importRowError: "錯誤", importSummary: "摘要",
          importNoRows: "找不到有效資料列。", importFileError: "無法讀取檔案。",
          actionExport: "匯出", actionImport: "匯入",           importModeHint: "已存在的主鍵會更新；新主鍵會新增。",
          filterStatus: "篩選狀態", searchPlaceholder: "搜尋...", searchHint: "可搜發票號、客戶、工作編號、PO 等。",
          searchHintClients: "可搜尋客戶編號、公司名稱、聯絡人、電子郵件、電話、地址等。",
          searchHintQuotation: "可搜尋報價單號、客戶、工作編號、說明、狀態等。",
          searchHintJob: "可搜尋工作編號、客戶、PO、報價單號、說明、狀態等。",
          searchHintVendors: "可搜尋供應商編號、名稱、電子郵件、銀行、SWIFT、聯絡人等。",
          region: "地區", regionAccess: "地區權限",           regionScopeHint: "您新增的資料會標記為目前地區；ALL 顯示您有權限的所有地區。",
          regionDataHint: "新增資料會儲存於目前所選地區（選 ALL 時使用您有權限的第一個地區）。",
          importRegionHint: "已匯入至地區：{region}。請確認右上角「地區」設為 ALL 或正確地區後才會看到資料。",
          settingsBackup: "備份與還原", backupTitle: "備份與還原", backupHint: "所有清單資料儲存在此瀏覽器。下載備份檔可搬移至另一台電腦或瀏覽器。還原會覆蓋目前所有資料。",
          backupDownload: "下載備份", restoreUpload: "還原備份", restoreConfirm: "還原會覆蓋此瀏覽器內所有 ERP 資料，確定繼續？",
          restoreDone: "還原完成，頁面將重新載入。", restoreInvalid: "備份檔案無效。", backupDone: "已下載備份檔。",
          cloudSyncTitle: "雲端資料庫（Supabase）", cloudSyncHint: "上傳／下載經由 API 連接 Supabase。同步金鑰須與 Vercel 環境變數 ERP_SYNC_SECRET 相同。",
          cloudSyncKey: "雲端同步金鑰", cloudSyncKeyPlaceholder: "與伺服器 ERP_SYNC_SECRET 相同",
          cloudPush: "上傳至雲端", cloudPull: "從雲端下載", cloudPushDone: "已上傳至 Supabase。", cloudPullDone: "已從雲端下載，正在重新載入…",
          cloudSyncFail: "雲端同步失敗", cloudNoBackup: "資料庫內沒有備份。", cloudNotConfigured: "伺服器尚未設定雲端同步。",
          liveSyncTitle: "即時多人同步", liveSyncHint: "每筆資料（客戶、工作、應收應付等）自動寫入 Supabase，其他用戶幾秒內會看到更新。需先執行 SQL migration 002。",
          liveSyncEnable: "啟用即時同步（自動上傳及下載）", liveSyncOk: "即時同步運作中", liveSyncReadOnly: "即時同步運作中（唯讀）", liveSyncSyncing: "同步中…", liveSyncError: "同步錯誤",
          liveSyncLast: "上次同步", liveSyncNeedKey: "請先輸入雲端同步金鑰。", liveSyncBootstrapped: "已啟用即時同步，資料已連接雲端。",
          liveSyncAutoHint: "開啟網站即自動同步。唯讀用戶只下載；有編輯權限的用戶才會上傳變更。",
          loginUsername: "用戶名稱", loginPassword: "密碼", loginRemember: "記住我", loginBtn: "登入", loginFail: "用戶名稱或密碼錯誤。", loginInactive: "此帳號已停用。", logout: "登出",
          myAccountTitle: "我的帳號", changePassword: "更改密碼", currentPassword: "現有密碼", newPassword: "新密碼", confirmPassword: "確認新密碼", passwordChanged: "密碼已更新。", passwordMismatch: "兩次輸入的新密碼不一致。", passwordWrong: "現有密碼不正確。", userLoginId: "用戶名稱",
          noDatabaseNote: "目前尚無伺服器資料庫 — Excel 匯入只會更新瀏覽器資料。上傳後請在預覽視窗按「確認匯入」。"
        },
      };

      function paymentStatusLabel(status, t) {
        if (status === "Paid") return t("paid");
        if (status === "Cancel" || status === "Cancelled" || status === "Canceled") return t("paymentCancel");
        if (status === "Overdue") return t("overdue");
        return t("awaitingPayment");
      }

      function deriveArPaymentStatus(r) {
        if (!r) return "Awaiting Payment";
        const stored = String(r.payment_status || "").trim();
        if (stored === "Cancel" || stored === "Cancelled" || stored === "Canceled") return "Cancel";
        if (r.payment_received_date) return "Paid";
        if (r.due_date && calcOverdueDays(r.due_date) > 0) return "Overdue";
        return "Awaiting Payment";
      }

      function sumLinkedQuotationAmounts(job, quotations, targetCurrency) {
        return findLinkedQuotationsForJob(job, quotations).reduce((s, q) => s + convertCurrency(Number(q.amount || 0), q.currency || "USD", targetCurrency), 0);
      }

      function sumJobPoAmount(job) {
        return Number(normalizeJobPoFields(job).customer_po_amount || 0);
      }

      function sumJobArInvoiceAmount(jobNo, arList, targetCurrency) {
        return (arList || []).filter((r) => r.job_no === jobNo).reduce((s, r) => s + resolveArDisplayAmount(r, targetCurrency), 0);
      }

      function paymentStatusClass(status) {
        const base = "inline-flex items-center whitespace-nowrap px-2 py-0.5 rounded text-xs";
        if (status === "Paid") return base + " text-green-700 bg-green-100";
        if (status === "Overdue") return base + " text-red-700 bg-red-100";
        return base + " text-amber-700 bg-amber-100";
      }

      function isJobCompleted(status) {
        return status === "Completed" || status === "Complete";
      }

      function jobStatusClass(status) {
        if (isJobCompleted(status)) return "text-green-800 bg-green-100 font-semibold ring-1 ring-green-300";
        if (status === "In Progress") return "text-blue-700 bg-blue-100";
        if (status === "On Hold") return "text-slate-600 bg-slate-100";
        return "text-amber-700 bg-amber-100";
      }

      function deriveApPaymentStatus(b) {
        if (!b) return "Awaiting Payment";
        const stored = String(b.payment_status || "").trim();
        if (stored === "Cancel" || stored === "Cancelled" || stored === "Canceled") return "Cancel";
        if (b.pay_date) return "Paid";
        if (b.due_date && calcOverdueDays(b.due_date) > 0) return "Overdue";
        return "Awaiting Payment";
      }

      function isApPaid(bill) {
        return deriveApPaymentStatus(bill) === "Paid";
      }

      function isApUnpaid(bill) {
        const s = deriveApPaymentStatus(bill);
        return s === "Awaiting Payment" || s === "Overdue";
      }

      function apPaymentLabel(bill, t) {
        return paymentStatusLabel(deriveApPaymentStatus(bill), t);
      }

      function apPaymentClass(bill) {
        return paymentStatusClass(deriveApPaymentStatus(bill));
      }

      // Lucide-style sidebar icons (stroke 2, rounded caps — distinct from previous Heroicons set)
      const NAV_ICON_SVG = { xmlns: "http://www.w3.org/2000/svg", viewBox: "0 0 24 24", fill: "none", stroke: "currentColor", strokeWidth: 2, strokeLinecap: "round", strokeLinejoin: "round" };

      const NAV_ICONS = {
        dashboard: [
          <rect key="1" width="7" height="9" x="3" y="3" rx="1" />,
          <rect key="2" width="7" height="5" x="14" y="3" rx="1" />,
          <rect key="3" width="7" height="9" x="14" y="12" rx="1" />,
          <rect key="4" width="7" height="5" x="3" y="16" rx="1" />
        ],
        overview: [
          <path key="1" d="M21.21 15.89A10 10 0 1 1 8 2.83" />,
          <path key="2" d="M22 12A10 10 0 0 0 12 2v10z" />
        ],
        management_report: [
          <path key="1" d="M12 20V10" />,
          <path key="2" d="M18 20V4" />,
          <path key="3" d="M6 20v-4" />
        ],
        clients: [
          <path key="1" d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2" />,
          <circle key="2" cx="9" cy="7" r="4" />,
          <path key="3" d="M22 21v-2a4 4 0 0 0-3-3.87" />,
          <path key="4" d="M16 3.13a4 4 0 0 1 0 7.75" />
        ],
        job: [
          <rect key="1" width="8" height="4" x="8" y="2" rx="1" ry="1" />,
          <path key="2" d="M16 4h2a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h2" />,
          <path key="3" d="M12 11h4" />,
          <path key="4" d="M12 16h4" />,
          <path key="5" d="M8 11h.01" />,
          <path key="6" d="M8 16h.01" />
        ],
        quotation: [
          <path key="1" d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z" />,
          <path key="2" d="M14 2v6h6" />,
          <path key="3" d="M16 13H8" />,
          <path key="4" d="M16 17H8" />,
          <path key="5" d="M10 9H8" />
        ],
        ongoing: [
          <circle key="1" cx="12" cy="12" r="10" />,
          <path key="2" d="M12 6v6l4 2" />
        ],
        monthly_report: [
          <path key="1" d="M3 3v18h18" />,
          <path key="2" d="M18 17V9" />,
          <path key="3" d="M13 17V5" />,
          <path key="4" d="M8 17v-3" />
        ],
        ar: [
          <path key="1" d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1Z" />,
          <path key="2" d="M16 8h-6a2 2 0 1 0 0 4h4a2 2 0 1 1 0 4H8" />,
          <path key="3" d="M12 17V7" />
        ],
        ap: [
          <rect key="1" width="20" height="12" x="2" y="6" rx="2" />,
          <circle key="2" cx="12" cy="12" r="2" />,
          <path key="3" d="M6 12h.01M18 12h.01" />
        ],
        vendors: [
          <path key="1" d="M10 12h4" />,
          <path key="2" d="M10 8h4" />,
          <path key="3" d="M14 21v-3a2 2 0 0 0-4 0v3" />,
          <path key="4" d="M6 10H4a2 2 0 0 0-2 2v7a2 2 0 0 0 2 2h16a2 2 0 0 0 2-2V9a2 2 0 0 0-2-2h-2" />,
          <path key="5" d="M6 21V5a2 2 0 0 1 2-2h8a2 2 0 0 1 2 2v16" />
        ],
        settings: [
          <path key="1" d="M12.22 2h-.44a2 2 0 0 0-2 2v.18a2 2 0 0 1-1 1.73l-.43.25a2 2 0 0 1-2 0l-.15-.08a2 2 0 0 0-2.73.73l-.22.38a2 2 0 0 0 .73 2.73l.15.1a2 2 0 0 1 1 1.72v.51a2 2 0 0 1-1 1.74l-.15.09a2 2 0 0 0-.73 2.73l.22.38a2 2 0 0 0 2.73.73l.15-.08a2 2 0 0 1 2 0l.43.25a2 2 0 0 1 1 1.73V20a2 2 0 0 0 2 2h.44a2 2 0 0 0 2-2v-.18a2 2 0 0 1 1-1.73l.43-.25a2 2 0 0 1 2 0l.15.08a2 2 0 0 0 2.73-.73l.22-.39a2 2 0 0 0-.73-2.73l-.15-.08a2 2 0 0 1-1-1.74v-.5a2 2 0 0 1 1-1.74l.15-.09a2 2 0 0 0 .73-2.73l-.22-.38a2 2 0 0 0-2.73-.73l-.15.08a2 2 0 0 1-2 0l-.43-.25a2 2 0 0 1-1-1.73V4a2 2 0 0 0-2-2z" />,
          <circle key="2" cx="12" cy="12" r="3" />
        ]
      };

      function SidebarNavIcon({ id, className }) {
        const icon = NAV_ICONS[id];
        return (
          <svg {...NAV_ICON_SVG} className={className || "w-5 h-5 shrink-0"} aria-hidden="true">
            {icon || <circle cx="12" cy="12" r="2" />}
          </svg>
        );
      }

      const fxUsdMap = { USD: 1, HKD: 7.8, SGD: 1.35, TWD: 32.2, RMB: 7.2, EUR: 0.92, JPY: 160, MYR: 4.7, IDR: 16000, AUD: 1.55 };

      function calcOverdueDays(dueDate) {
        if (!dueDate) return 0;
        const now = new Date();
        const due = new Date(dueDate + "T00:00:00");
        const diff = Math.floor((now - due) / (1000 * 60 * 60 * 24));
        return diff > 0 ? diff : 0;
      }

      function daysUntilDue(dueDate) {
        if (!dueDate) return null;
        const now = new Date();
        now.setHours(0, 0, 0, 0);
        const due = new Date(dueDate + "T00:00:00");
        return Math.floor((due - now) / (1000 * 60 * 60 * 24));
      }

      function dueBadgeLabel(daysUntil, overdueDays, t) {
        if (daysUntil === 0) return t("dashDueToday");
        if (daysUntil !== null && daysUntil > 0) return t("dashDueInDays").replace("{n}", String(daysUntil));
        if (overdueDays > 0) return t("dashDaysOverdue").replace("{n}", String(overdueDays));
        return "-";
      }

      function dueBadgeClass(daysUntil, overdueDays) {
        if (overdueDays > 0) return "bg-red-100 text-red-800 ring-1 ring-red-200";
        if (daysUntil === 0) return "bg-amber-100 text-amber-900 ring-1 ring-amber-200";
        if (daysUntil !== null && daysUntil <= 7) return "bg-amber-50 text-amber-800";
        return "bg-slate-100 text-slate-700";
      }

      function calcDaysBetween(startDate, endDate) {
        if (!startDate || !endDate) return 0;
        const start = new Date(startDate + "T00:00:00");
        const end = new Date(endDate + "T00:00:00");
        const diff = Math.floor((end - start) / (1000 * 60 * 60 * 24));
        return diff > 0 ? diff : 0;
      }

      function parsePaymentTermsDays(terms) {
        if (!terms || !String(terms).trim()) return 30;
        const m = String(terms).match(/(\d+)/);
        return m ? Math.max(0, parseInt(m[1], 10)) : 30;
      }

      function addDaysIso(isoDate, days) {
        if (!isoDate) return "";
        const d = new Date(isoDate + "T12:00:00");
        if (Number.isNaN(d.getTime())) return "";
        d.setDate(d.getDate() + days);
        return d.toISOString().slice(0, 10);
      }

      function findClientForAr(data, clients, jobs) {
        const job = data.job_id ? jobs.find((j) => j.id === Number(data.job_id)) : null;
        if (job) {
          const byId = clients.find((c) => c.id === job.client_id);
          if (byId) return byId;
          const byCo = clients.find((c) => c.company === job.company);
          if (byCo) return byCo;
        }
        if (data.customer) return clients.find((c) => c.company === data.customer || c.invoice_title === data.customer) || null;
        return null;
      }

      function computeArDueDate(data, clients, jobs) {
        if (!data.invoice_date) return data.due_date || "";
        const client = findClientForAr(data, clients, jobs);
        return addDaysIso(data.invoice_date, parsePaymentTermsDays(client ? client.payment_terms : ""));
      }

      function applyArDueDate(data, clients, jobs) {
        return { ...data, due_date: computeArDueDate(data, clients, jobs) };
      }

      function monthChartLabel(year, monthIndex, lang) {
        const locale = i18nLang(lang) === "zh_TW" ? "zh-TW" : "en-US";
        return new Intl.DateTimeFormat(locale, { month: "short", year: "numeric" }).format(new Date(year, monthIndex, 1));
      }

      function monthShortLabel(year, monthIndex, lang) {
        const locale = i18nLang(lang) === "zh_TW" ? "zh-TW" : "en-US";
        return new Intl.DateTimeFormat(locale, { month: "short" }).format(new Date(year, monthIndex, 1));
      }

      function monthTableHeader(year, monthIndex, lang) {
        const locale = i18nLang(lang) === "zh_TW" ? "zh-TW" : "en-US";
        return new Intl.DateTimeFormat(locale, { month: "long" }).format(new Date(year, monthIndex, 1));
      }

      function monthRangeFor(year, monthIndex) {
        const monthNum = monthIndex + 1;
        const key = year + "-" + String(monthNum).padStart(2, "0");
        const from = key + "-01";
        const lastDay = new Date(year, monthNum, 0).getDate();
        const to = key + "-" + String(lastDay).padStart(2, "0");
        return { from, to, key };
      }

      let _erpIdSeq = 0;

      function resetErpIdAllocator() {
        _erpIdSeq = 0;
      }

      function erpScanMaxId(...sources) {
        let max = 0;
        sources.forEach((src) => {
          (Array.isArray(src) ? src : []).forEach((r) => {
            const id = Number(r?.id);
            if (Number.isFinite(id) && id > max) max = id;
          });
        });
        return max;
      }

      function nextErpRecordId(...sources) {
        if (_erpIdSeq === 0) _erpIdSeq = erpScanMaxId(...sources);
        _erpIdSeq += 1;
        return _erpIdSeq;
      }

      function nextMonthRange(year, monthIndex) {
        if (monthIndex >= 11) return monthRangeFor(year + 1, 0);
        return monthRangeFor(year, monthIndex + 1);
      }

      function monthEndIso(year, monthIndex) {
        return monthRangeFor(year, monthIndex).to;
      }

      function reportMonthHasEnded(year, monthIndex, todayIso) {
        return todayIso > monthEndIso(year, monthIndex);
      }

      function arExpectedSnapshotKey(year, monthIndex, region) {
        return year + "-" + String(monthIndex + 1).padStart(2, "0") + "-" + region;
      }

      function isArCancelled(r) {
        const s = String(r && r.payment_status || "").toLowerCase();
        return s === "cancel" || s === "cancelled" || s === "canceled";
      }

      function wasArUnpaidAsOf(r, asOfIso) {
        if (!r || isArCancelled(r)) return false;
        const paidDate = r.payment_received_date || "";
        if (paidDate) return paidDate > asOfIso;
        if (r.payment_status === "Paid") {
          const fallback = r.invoice_date || r.due_date || "";
          return !fallback || fallback > asOfIso;
        }
        return true;
      }

      function arCountsForExpectedMonth(r, year, monthIndex, asOfIso) {
        if (!wasArUnpaidAsOf(r, asOfIso) || !r.due_date) return false;
        const { from: nextFrom, to: nextTo } = nextMonthRange(year, monthIndex);
        if (dateInRange(r.due_date, nextFrom, nextTo)) return true;
        if (r.due_date < nextFrom && r.due_date <= asOfIso) return true;
        return false;
      }

      function computeArExpectedForReportMonth(year, monthIndex, arInvoices, currency, asOfIso) {
        return arInvoices
          .filter((r) => arCountsForExpectedMonth(r, year, monthIndex, asOfIso))
          .reduce((s, r) => s + resolveArDisplayAmount(r, currency), 0);
      }

      function buildMonthlyReportSeries(year, currency, lang, poMonthTotals, invoiceMonthTotals, arInvoices, snapshots, region, todayIso) {
        const y = Number(year);
        return Array.from({ length: 12 }, (_, m) => {
          const { key } = monthRangeFor(y, m);
          const snapKey = arExpectedSnapshotKey(y, m, region);
          const monthEnded = reportMonthHasEnded(y, m, todayIso);
          const asOfIso = monthEnded ? monthEndIso(y, m) : todayIso;
          const arExpected = monthEnded && snapshots[snapKey] != null
            ? snapshots[snapKey]
            : computeArExpectedForReportMonth(y, m, arInvoices, currency, asOfIso);
          return {
            monthIndex: m,
            key,
            label: monthChartLabel(y, m, lang),
            header: monthTableHeader(y, m, lang),
            currency,
            poReceived: poMonthTotals[m] || 0,
            invoiceIssued: invoiceMonthTotals[m] || 0,
            arExpected
          };
        });
      }

      function buildMonthDetailGrid(lines, year, lang) {
        const y = Number(year);
        const months = Array.from({ length: 12 }, (_, m) => {
          const monthLines = lines.filter((l) => l.year === y && l.month === m + 1);
          const total = monthLines.reduce((s, l) => s + Number(l.base_amount || 0), 0);
          return { monthIndex: m, header: monthShortLabel(y, m, lang), rows: monthLines, total };
        });
        const maxRows = Math.max(1, ...months.map((m) => m.rows.length));
        return { months, maxRows };
      }

      function syncPoLinesFromJobs(jobs, year, baseCurrency, existing) {
        const y = Number(year);
        const next = [...existing];
        jobs.map(normalizeJobPoFields).forEach((job) => {
          const lineCurrency = REGION_CURRENCY[job.region] || baseCurrency;
          const jobCurrency = job.currency || "USD";
          const poEntries = (job.po_lines && job.po_lines.length)
            ? job.po_lines.map((pl) => ({ po_no: pl.po_no || "", sourceAmt: Number(pl.amount || 0), date: pl.received_date || job.start_date || "" }))
            : [{ po_no: job.customer_po || "", sourceAmt: Number(job.customer_po_amount ?? job.amount ?? 0), date: job.start_date || "" }];
          for (let m = 0; m < 12; m++) {
            const { from, to } = monthRangeFor(y, m);
            poEntries.forEach((pe, poIdx) => {
              if (!pe.date || !dateInRange(pe.date, from, to) || !pe.sourceAmt) return;
              const sourceKey = "job-" + job.id + "-" + y + "-" + (m + 1) + "-po" + poIdx;
              const idx = next.findIndex((l) => l.source_key === sourceKey);
              const baseAmt = convertCurrency(pe.sourceAmt, jobCurrency, lineCurrency);
              const displayJobNo = pe.po_no ? job.job_no + " / " + pe.po_no : job.job_no;
              const patch = {
                region: job.region || "HK",
                year: y,
                month: m + 1,
                job_no: displayJobNo,
                source_key: sourceKey,
                other_currency: jobCurrency,
                source_amount: pe.sourceAmt,
                base_currency: lineCurrency,
                base_amount: baseAmt
              };
              if (idx >= 0) {
                next[idx] = { ...next[idx], ...patch, other_currency: next[idx].other_currency || patch.other_currency };
              } else {
                next.push({ id: nextErpRecordId(next), ...patch });
              }
            });
          }
        });
        return next;
      }

      function syncArLinesFromInvoices(arInvoices, year, baseCurrency, existing) {
        const y = Number(year);
        const next = [...existing];
        for (let m = 0; m < 12; m++) {
          const { from, to } = monthRangeFor(y, m);
          arInvoices.filter((r) => r.invoice_date && dateInRange(r.invoice_date, from, to)).forEach((r) => {
            const sourceKey = "ar-" + r.id;
            const idx = next.findIndex((l) => l.source_key === sourceKey);
            const lineCurrency = REGION_CURRENCY[r.region] || baseCurrency;
            const invCurrency = r.invoice_currency || "USD";
            const sourceAmt = Number(r.invoice_amt || 0);
            const baseAmt = resolveArDisplayAmount(r, lineCurrency);
            const patch = {
              region: r.region || "HK",
              year: y,
              month: m + 1,
              job_no: r.job_no || "",
              source_key: sourceKey,
              other_currency: invCurrency,
              source_amount: sourceAmt,
              base_currency: lineCurrency,
              base_amount: baseAmt
            };
            if (idx >= 0) {
              next[idx] = {
                ...next[idx],
                ...patch,
                other_currency: next[idx].other_currency || patch.other_currency
              };
            } else {
              next.push({ id: nextErpRecordId(next), ...patch });
            }
          });
        }
        return next;
      }

      function dateInRange(iso, from, to) {
        if (!iso || !from || !to) return false;
        return iso >= from && iso <= to;
      }

      function convertCurrency(amount, fromCurrency, baseCurrency) {
        if (!amount) return 0;
        const fromRate = fxUsdMap[fromCurrency] || 1;
        const baseRate = fxUsdMap[baseCurrency] || 1;
        return (Number(amount) / fromRate) * baseRate;
      }

      function isArUnpaid(r) {
        return deriveArPaymentStatus(r) === "Awaiting Payment" || deriveArPaymentStatus(r) === "Overdue";
      }

      function resolveArDisplayAmount(r, targetCurrency) {
        const invAmt = Number(r.invoice_amt);
        if (invAmt > 0) {
          return convertCurrency(invAmt, r.invoice_currency || "USD", targetCurrency);
        }
        const base = Number(r.base_amount_locked ?? r.base_amount);
        if (base > 0) {
          const recordCur = REGION_CURRENCY[r.region] || r.invoice_currency || "HKD";
          return convertCurrency(base, recordCur, targetCurrency);
        }
        return 0;
      }

      function resolveApDisplayAmount(b, targetCurrency) {
        const amt = Number(b.amount);
        if (amt > 0) {
          return convertCurrency(amt, b.currency || "USD", targetCurrency);
        }
        const base = Number(b.base_amount_locked ?? b.base_amount);
        if (base > 0) {
          const recordCur = REGION_CURRENCY[b.region] || b.currency || "HKD";
          return convertCurrency(base, recordCur, targetCurrency);
        }
        return 0;
      }

      function dueDateMonthKey(dueDate, chartYear) {
        if (!dueDate) return null;
        const s = String(dueDate).trim();
        let y;
        let m;
        if (/^\d{4}-\d{2}-\d{2}/.test(s)) {
          y = s.slice(0, 4);
          m = s.slice(5, 7);
        } else {
          const d = new Date(s.includes("T") ? s : s + "T00:00:00");
          if (Number.isNaN(d.getTime())) return null;
          y = String(d.getFullYear());
          m = String(d.getMonth() + 1).padStart(2, "0");
        }
        if (chartYear && y !== String(chartYear)) return chartYear + "-" + m;
        return y + "-" + m;
      }

      function money(value) {
        return Number(value || 0).toLocaleString(undefined, { maximumFractionDigits: 2 });
      }

      function nowIso() { return new Date().toISOString(); }

      function emptyClient() {
        return { customer_no: "", company: "", invoice_title: "", gst_no: "", primary_contact: "", company_phone: "", mobile_phone: "", email: "", address: "", postal_code: "", account_dept_contact: "", payment_terms: "", is_bu: false, bu_no: "" };
      }

      function nextCustomerNo(clientList) {
        let bestPrefix = "SC";
        let bestNum = 0;
        let bestWidth = 3;
        (clientList || []).forEach((c) => {
          const m = String(c.customer_no || "").trim().match(/^([A-Za-z]*)(\d+)$/);
          if (!m) return;
          const prefix = m[1];
          const num = parseInt(m[2], 10);
          if (Number.isNaN(num)) return;
          const width = m[2].length;
          if (num > bestNum) {
            bestNum = num;
            bestPrefix = prefix || "SC";
            bestWidth = width;
          }
        });
        if (bestNum === 0) return bestPrefix + String(1).padStart(bestWidth, "0");
        return bestPrefix + String(bestNum + 1).padStart(bestWidth, "0");
      }

      function parseBuFlag(value) {
        if (value === true) return true;
        const v = String(value || "").trim().toLowerCase();
        return v === "y" || v === "yes" || v === "1" || v === "true";
      }

      function emptyJob() {
        return { job_no: "", client_id: "", company: "", job_type: "Project", customer_po: "", customer_po_amount: "", quotation_no: "", linked_quotation_nos: [], amount: "", currency: "USD", status: "Open", description: "", remarks: "", start_date: "", po_lines: [] };
      }

      function emptyPoLine() {
        return { id: 0, po_no: "", amount: "", received_date: "", remarks: "" };
      }

      function normalizeJobPoFields(job) {
        const j = { ...job };
        if (!Array.isArray(j.po_lines) || !j.po_lines.length) {
          if (j.customer_po || j.customer_po_amount) {
            j.po_lines = [{ id: 1, po_no: j.customer_po || "", amount: Number(j.customer_po_amount || 0), received_date: j.start_date || "", remarks: "" }];
          } else {
            j.po_lines = [];
          }
        } else {
          j.po_lines = j.po_lines.map((l, i) => ({
            id: l.id || i + 1,
            po_no: l.po_no || "",
            amount: Number(l.amount || 0),
            received_date: l.received_date || "",
            remarks: l.remarks || ""
          }));
        }
        const total = j.po_lines.reduce((s, l) => s + Number(l.amount || 0), 0);
        if (j.po_lines.length) {
          j.customer_po_amount = total;
          j.customer_po = j.po_lines.map((l) => l.po_no).filter(Boolean).join(", ");
          if (!j.start_date) {
            const dates = j.po_lines.map((l) => l.received_date).filter(Boolean).sort();
            if (dates.length) j.start_date = dates[0];
          }
        }
        return j;
      }

      function ensureJobsPoLines(arr) {
        return (arr || []).map(normalizeJobPoFields);
      }

      function jobQuotationNos(j) {
        if (!j) return [];
        if (Array.isArray(j.linked_quotation_nos) && j.linked_quotation_nos.length) {
          return j.linked_quotation_nos.map((n) => String(n || "").trim()).filter(Boolean);
        }
        if (j.quotation_no) {
          return String(j.quotation_no).split(",").map((s) => s.trim()).filter(Boolean);
        }
        return [];
      }

      function addQuotationToJob(j, quotationNo) {
        const q = String(quotationNo || "").trim();
        if (!q) return j;
        const nos = [...jobQuotationNos(j)];
        if (!nos.some((n) => n.toLowerCase() === q.toLowerCase())) nos.push(q);
        return { ...j, linked_quotation_nos: nos, quotation_no: nos.join(", ") };
      }

      function removeQuotationFromJob(j, quotationNo) {
        const k = String(quotationNo || "").trim().toLowerCase();
        const nos = jobQuotationNos(j).filter((n) => n.toLowerCase() !== k);
        return { ...j, linked_quotation_nos: nos, quotation_no: nos.join(", ") };
      }

      function normalizeJobQuotations(job) {
        const j = { ...job };
        const nos = jobQuotationNos(j);
        j.linked_quotation_nos = nos;
        j.quotation_no = nos.join(", ");
        return j;
      }

      function ensureJobsQuotations(arr) {
        return ensureJobsPoLines(arr).map(normalizeJobQuotations);
      }

      function quotationJobNos(q) {
        if (!q) return [];
        if (Array.isArray(q.linked_job_nos) && q.linked_job_nos.length) return q.linked_job_nos.filter(Boolean);
        if (q.job_no) return String(q.job_no).split(",").map((s) => s.trim()).filter(Boolean);
        return [];
      }

      function quotationJobIds(q) {
        if (!q) return [];
        if (Array.isArray(q.linked_job_ids) && q.linked_job_ids.length) return q.linked_job_ids;
        if (q.job_id != null) return [q.job_id];
        return [];
      }

      function addJobToQuotation(q, jobId, jobNo) {
        const ids = [...quotationJobIds(q)];
        const nos = [...quotationJobNos(q)];
        if (jobId != null && !ids.includes(jobId)) ids.push(jobId);
        if (jobNo && !nos.includes(jobNo)) nos.push(jobNo);
        return { ...q, linked_job_ids: ids, linked_job_nos: nos, job_id: ids[0] || null, job_no: nos.join(", ") };
      }

      function removeJobFromQuotation(q, jobId, jobNo) {
        const ids = quotationJobIds(q).filter((id) => id !== jobId);
        const nos = quotationJobNos(q).filter((n) => n !== jobNo);
        return { ...q, linked_job_ids: ids, linked_job_nos: nos, job_id: ids[0] || null, job_no: nos.join(", ") };
      }

      function findDuplicateInvoiceNo(list, invoiceNo, excludeId) {
        const k = String(invoiceNo || "").trim().toLowerCase();
        if (!k) return null;
        return (list || []).find((r) => r.id !== excludeId && String(r.invoice_no || "").trim().toLowerCase() === k) || null;
      }

      function quotationLinksJob(q, jobNo) {
        if (!q || !jobNo) return false;
        return quotationJobNos(q).some((n) => n.toLowerCase() === String(jobNo).trim().toLowerCase());
      }

      function paymentStageLabel(stage, t) {
        if (!stage) return "-";
        return stage;
      }

      function jobNatureCategory(jobType) {
        const s = String(jobType || "");
        if (/trading/i.test(s)) return "Trade";
        if (s === "Maintenance") return "Maintenance";
        if (s === "Service") return "Service";
        return "Project";
      }

      const MGMT_GROUP_CURRENCIES = ["HKD", "TWD", "SGD", "USD"];
      const MGMT_JOB_NATURES = ["Maintenance", "Service", "Project", "Trade"];
      const MGMT_JOB_TYPE_COLORS = {
        Maintenance: { from: "#1e40af", to: "#60a5fa", light: "#93c5fd" },
        Service: { from: "#5b21b6", to: "#a78bfa", light: "#c4b5fd" },
        Project: { from: "#047857", to: "#34d399", light: "#6ee7b7" },
        Trade: { from: "#c2410c", to: "#fb923c", light: "#fdba74" }
      };
      const MGMT_CARD_GRADIENTS = {
        revenue: "linear-gradient(148deg, #0c4a3e 0%, #0f766e 38%, #2dd4bf 100%)",
        cogs: "linear-gradient(148deg, #431407 0%, #b45309 42%, #fbbf24 100%)",
        profit: "linear-gradient(148deg, #1e1b4b 0%, #4338ca 48%, #818cf8 100%)",
        loss: "linear-gradient(148deg, #450a0a 0%, #b91c1c 50%, #f87171 100%)"
      };

      function polarToCartesian(cx, cy, r, angle) {
        return { x: cx + r * Math.cos(angle), y: cy + r * Math.sin(angle) };
      }

      function describeDonutSlice(cx, cy, outerR, innerR, startAngle, endAngle) {
        const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
        const oStart = polarToCartesian(cx, cy, outerR, startAngle);
        const oEnd = polarToCartesian(cx, cy, outerR, endAngle);
        const iStart = polarToCartesian(cx, cy, innerR, endAngle);
        const iEnd = polarToCartesian(cx, cy, innerR, startAngle);
        return [
          "M", oStart.x, oStart.y,
          "A", outerR, outerR, 0, largeArc, 1, oEnd.x, oEnd.y,
          "L", iStart.x, iStart.y,
          "A", innerR, innerR, 0, largeArc, 0, iEnd.x, iEnd.y,
          "Z"
        ].join(" ");
      }

      function MgmtDonutChart({ breakdown, size, centerLabel, centerSub }) {
        const chartSize = size || 168;
        const cx = chartSize / 2;
        const cy = chartSize / 2;
        const outerR = chartSize / 2 - 6;
        const innerR = outerR * 0.58;
        const slices = breakdown.filter((x) => x.profit > 0);
        const total = slices.reduce((s, x) => s + x.profit, 0) || 1;
        const gap = slices.length > 1 ? 0.045 : 0;
        let cursor = -Math.PI / 2;

        const segments = slices.map((slice) => {
          const sweep = (slice.profit / total) * Math.PI * 2;
          const start = cursor + gap / 2;
          const end = cursor + sweep - gap / 2;
          cursor += sweep;
          const mid = (start + end) / 2;
          const palette = MGMT_JOB_TYPE_COLORS[slice.type] || { from: "#475569", to: "#94a3b8", light: "#cbd5e1" };
          return { slice, start, end, mid, palette, path: describeDonutSlice(cx, cy, outerR, innerR, start, end) };
        });

        return (
          <div className="relative shrink-0" style={{ width: chartSize, height: chartSize }}>
            <svg width={chartSize} height={chartSize} viewBox={"0 0 " + chartSize + " " + chartSize} className="drop-shadow-lg">
              <defs>
                <filter id="mgmtDonutShadow" x="-20%" y="-20%" width="140%" height="140%">
                  <feDropShadow dx="0" dy="3" stdDeviation="4" floodColor="#0f172a" floodOpacity="0.22" />
                </filter>
                {segments.map((seg) => (
                  <linearGradient key={seg.slice.type + "-grad"} id={"mgmtGrad-" + seg.slice.type} x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stopColor={seg.palette.from} />
                    <stop offset="55%" stopColor={seg.palette.to} />
                    <stop offset="100%" stopColor={seg.palette.light} />
                  </linearGradient>
                ))}
                <radialGradient id="mgmtDonutCenterGlow" cx="50%" cy="45%" r="55%">
                  <stop offset="0%" stopColor="#ffffff" />
                  <stop offset="100%" stopColor="#f1f5f9" />
                </radialGradient>
              </defs>
              <circle cx={cx} cy={cy} r={outerR + 2} fill="none" stroke="#e2e8f0" strokeWidth="1" opacity="0.6" />
              <g filter="url(#mgmtDonutShadow)">
                {segments.map((seg) => (
                  <path
                    key={seg.slice.type}
                    d={seg.path}
                    fill={"url(#mgmtGrad-" + seg.slice.type + ")"}
                    stroke="#ffffff"
                    strokeWidth="2"
                    strokeLinejoin="round"
                  />
                ))}
              </g>
              <circle cx={cx} cy={cy} r={innerR - 3} fill="url(#mgmtDonutCenterGlow)" stroke="#e2e8f0" strokeWidth="1" />
              {centerLabel && (
                <text x={cx} y={cy - 2} textAnchor="middle" className="fill-slate-800" style={{ fontSize: "11px", fontWeight: 700 }}>
                  {centerLabel}
                </text>
              )}
              {centerSub && (
                <text x={cx} y={cy + 12} textAnchor="middle" className="fill-slate-500" style={{ fontSize: "9px", fontWeight: 600 }}>
                  {centerSub}
                </text>
              )}
            </svg>
          </div>
        );
      }

      function jobQualifiesForMgmt(job, quotations) {
        if (!job) return false;
        if (isJobCompleted(job.status)) return true;
        return findLinkedQuotationsForJob(job, quotations).some((q) => q.status === "Accepted");
      }

      function arQualifiesForMgmt(r, jobs, quotations) {
        if (!r || deriveArPaymentStatus(r) === "Cancel") return false;
        const job = jobs.find((j) => j.id === r.job_id || (r.job_no && j.job_no === r.job_no));
        if (!job) return false;
        return jobQualifiesForMgmt(job, quotations);
      }

      function apQualifiesForMgmt(b, jobs, quotations) {
        if (!b || deriveApPaymentStatus(b) === "Cancel") return false;
        const job = jobs.find((j) => j.id === b.job_id || (b.job_no && j.job_no === b.job_no));
        if (!job) return false;
        return jobQualifiesForMgmt(job, quotations);
      }

      function mgmtMonthBuckets(count, todayIso) {
        const today = todayIso ? new Date(todayIso) : new Date();
        const buckets = [];
        const d = new Date(today.getFullYear(), today.getMonth(), 1);
        for (let i = 0; i < count; i++) {
          const y = d.getFullYear();
          const m = d.getMonth();
          buckets.push({
            year: y,
            monthIndex: m,
            key: y + "-" + String(m + 1).padStart(2, "0")
          });
          d.setMonth(d.getMonth() + 1);
        }
        return buckets;
      }

      function buildManagementReport({ jobs, quotations, arInvoices, apBills, clientFilter, displayCurrency, compareCurrency, todayIso, lang, activeRegion, regionCount }) {
        const today = todayIso || new Date().toISOString().slice(0, 10);
        const cf = clientFilter && clientFilter !== "all" ? clientFilter : "all";

        const qualAr = arInvoices.filter((r) => {
          if (!arQualifiesForMgmt(r, jobs, quotations)) return false;
          if (cf !== "all" && r.customer !== cf) return false;
          return true;
        });
        const qualAp = apBills.filter((b) => {
          if (!apQualifiesForMgmt(b, jobs, quotations)) return false;
          if (cf !== "all") {
            const job = jobs.find((j) => j.id === b.job_id || j.job_no === b.job_no);
            if (!job || job.company !== cf) return false;
          }
          return true;
        });

        const grossRevenue = qualAr.reduce((s, r) => s + resolveArDisplayAmount(r, displayCurrency), 0);
        const totalCogs = qualAp.reduce((s, b) => s + resolveApDisplayAmount(b, displayCurrency), 0);
        const netProfit = grossRevenue - totalCogs;

        const clientMap = {};
        qualAr.forEach((r) => {
          const company = r.customer || "-";
          const job = jobs.find((j) => j.id === r.job_id || j.job_no === r.job_no);
          if (!clientMap[company]) clientMap[company] = { company, region: (job && job.region) || r.region || "HK", revenue: 0, cogs: 0 };
          clientMap[company].revenue += resolveArDisplayAmount(r, displayCurrency);
        });
        qualAp.forEach((b) => {
          const job = jobs.find((j) => j.id === b.job_id || j.job_no === b.job_no);
          const company = job ? job.company : "-";
          if (!clientMap[company]) clientMap[company] = { company, region: (job && job.region) || b.region || "HK", revenue: 0, cogs: 0 };
          clientMap[company].cogs += resolveApDisplayAmount(b, displayCurrency);
        });

        const clientMargins = Object.values(clientMap).map((c) => {
          const profit = c.revenue - c.cogs;
          const marginPct = c.revenue > 0 ? (profit / c.revenue) * 100 : 0;
          return { ...c, profit, marginPct };
        }).sort((a, b) => b.profit - a.profit);

        const allClientMargins = (() => {
          const map = {};
          arInvoices.filter((r) => arQualifiesForMgmt(r, jobs, quotations)).forEach((r) => {
            const company = r.customer || "-";
            const job = jobs.find((j) => j.id === r.job_id || j.job_no === r.job_no);
            if (!map[company]) map[company] = { company, revenue: 0, cogs: 0 };
            map[company].revenue += resolveArDisplayAmount(r, displayCurrency);
          });
          apBills.filter((b) => apQualifiesForMgmt(b, jobs, quotations)).forEach((b) => {
            const job = jobs.find((j) => j.id === b.job_id || j.job_no === b.job_no);
            const company = job ? job.company : "-";
            if (!map[company]) map[company] = { company, revenue: 0, cogs: 0 };
            map[company].cogs += resolveApDisplayAmount(b, displayCurrency);
          });
          return Object.values(map).map((c) => ({ ...c, profit: c.revenue - c.cogs }));
        })();
        const companyTotalProfit = allClientMargins.reduce((s, c) => s + c.profit, 0);
        const selectedClientProfit = cf !== "all" ? (clientMargins[0] ? clientMargins[0].profit : 0) : netProfit;
        const clientProfitSharePct = cf !== "all" && companyTotalProfit > 0 ? (selectedClientProfit / companyTotalProfit) * 100 : null;

        const typeProfits = {};
        MGMT_JOB_NATURES.forEach((n) => { typeProfits[n] = 0; });
        jobs.forEach((j) => {
          if (!jobQualifiesForMgmt(j, quotations)) return;
          if (cf !== "all" && j.company !== cf) return;
          const nature = jobNatureCategory(j.job_type);
          const rev = qualAr.filter((r) => r.job_id === j.id || r.job_no === j.job_no).reduce((s, r) => s + resolveArDisplayAmount(r, displayCurrency), 0);
          const cost = qualAp.filter((b) => b.job_id === j.id || b.job_no === j.job_no).reduce((s, b) => s + resolveApDisplayAmount(b, displayCurrency), 0);
          typeProfits[nature] = (typeProfits[nature] || 0) + (rev - cost);
        });
        const jobTypeBreakdown = MGMT_JOB_NATURES.map((type) => {
          const profit = typeProfits[type] || 0;
          const sharePct = netProfit !== 0 ? (profit / netProfit) * 100 : 0;
          return { type, profit, sharePct };
        }).filter((x) => x.profit !== 0 || grossRevenue > 0);

        const monthBuckets = mgmtMonthBuckets(3, today);
        const cashForecast = monthBuckets.map((mo) => {
          const label = monthChartLabel(mo.year, mo.monthIndex, lang);
          const cashIn = arInvoices.filter((r) => {
            if (!isArUnpaid(r)) return false;
            if (cf !== "all" && r.customer !== cf) return false;
            if (!r.due_date) return false;
            return r.due_date.slice(0, 7) === mo.key;
          }).reduce((s, r) => s + resolveArDisplayAmount(r, displayCurrency), 0);
          const cashOut = apBills.filter((b) => {
            if (!isApUnpaid(b)) return false;
            if (cf !== "all") {
              const job = jobs.find((j) => j.id === b.job_id || j.job_no === b.job_no);
              if (!job || job.company !== cf) return false;
            }
            if (!b.due_date) return false;
            return b.due_date.slice(0, 7) === mo.key;
          }).reduce((s, b) => s + resolveApDisplayAmount(b, displayCurrency), 0);
          const netCash = cashIn - cashOut;
          return { ...mo, label, cashIn, cashOut, netCash, alert: cashOut > cashIn };
        });

        const overdueAr = arInvoices.filter((r) => {
          if (!isArUnpaid(r)) return false;
          if (cf !== "all" && r.customer !== cf) return false;
          return calcOverdueDays(r.due_date) > 0;
        }).reduce((s, r) => s + resolveArDisplayAmount(r, displayCurrency), 0);

        const regionBreakdown = activeRegion === "ALL" ? DATA_REGION_IDS.map((regionId) => {
          const localCcy = REGION_CURRENCY[regionId] || "HKD";
          const regAr = arInvoices.filter((r) => (r.region || "HK") === regionId && arQualifiesForMgmt(r, jobs, quotations));
          const regAp = apBills.filter((b) => (b.region || "HK") === regionId && apQualifiesForMgmt(b, jobs, quotations));
          const rev = regAr.reduce((s, r) => s + resolveArDisplayAmount(r, localCcy), 0);
          const cogs = regAp.reduce((s, b) => s + resolveApDisplayAmount(b, localCcy), 0);
          const profit = rev - cogs;
          const clientTree = {};
          regAr.forEach((r) => {
            const company = r.customer || "-";
            if (!clientTree[company]) clientTree[company] = { company, revenue: 0, cogs: 0 };
            clientTree[company].revenue += resolveArDisplayAmount(r, localCcy);
          });
          regAp.forEach((b) => {
            const job = jobs.find((j) => j.id === b.job_id || j.job_no === b.job_no);
            const company = job ? job.company : "-";
            if (!clientTree[company]) clientTree[company] = { company, revenue: 0, cogs: 0 };
            clientTree[company].cogs += resolveApDisplayAmount(b, localCcy);
          });
          const clients = Object.values(clientTree).map((c) => {
            const p = c.revenue - c.cogs;
            const marginPct = c.revenue > 0 ? (p / c.revenue) * 100 : 0;
            return { ...c, profit: p, marginPct };
          }).filter((c) => c.revenue || c.cogs).sort((a, b) => b.profit - a.profit);
          return {
            regionId,
            localCurrency: localCcy,
            localProfit: profit,
            localRevenue: rev,
            convertedProfit: convertCurrency(profit, localCcy, compareCurrency),
            convertedRevenue: convertCurrency(rev, localCcy, compareCurrency),
            clients
          };
        }).filter((r) => r.localRevenue || r.localProfit) : [];

        const hasData = grossRevenue > 0 || totalCogs > 0 || clientMargins.length > 0;

        return {
          grossRevenue,
          totalCogs,
          netProfit,
          clientMargins,
          clientProfitSharePct,
          jobTypeBreakdown,
          cashForecast,
          overdueAr,
          regionBreakdown,
          hasData,
          regionCount: regionCount || DATA_REGION_IDS.length,
          displayCurrency,
          compareCurrency
        };
      }

      function mgmtNatureLabel(type, t) {
        if (type === "Maintenance") return t("mgmtNatureMaintenance");
        if (type === "Service") return t("mgmtNatureService");
        if (type === "Trade") return t("mgmtNatureTrade");
        return t("mgmtNatureProject");
      }

      function MgmtMarginBar({ row, maxProfit, onSelect, t }) {
        const profitAmt = Math.max(0, row.profit);
        const width = maxProfit > 0 ? Math.max(4, (profitAmt / maxProfit) * 100) : 0;
        const danger = row.marginPct < 10;
        const warn = row.marginPct < 20 && !danger;
        const barColor = danger ? "#dc2626" : warn ? "#ea580c" : "#2563eb";
        return (
          <button type="button" onClick={() => onSelect && onSelect(row.company)} className="w-full flex items-center gap-2 py-2.5 px-1 hover:bg-slate-50 rounded-lg text-left group">
            <span className="w-[32%] min-w-0 text-sm font-medium text-slate-800 truncate">
              {row.company}
              <span className="ml-1 text-slate-400 group-hover:text-blue-600 transition-colors">→</span>
              <span className="ml-1.5 text-[10px] font-semibold text-slate-500 bg-slate-100 px-1.5 py-0.5 rounded">{REGION_FLAGS[row.region] || ""} {row.region}</span>
            </span>
            <span className="flex-1 h-3 bg-slate-100 rounded-full overflow-hidden min-w-[3rem]">
              <span className="block h-full rounded-full transition-all" style={{ width: width + "%", background: barColor }} />
            </span>
            <span className={"min-w-[7.5rem] text-right text-sm font-semibold tabular-nums whitespace-nowrap " + (danger ? "text-red-600" : warn ? "text-orange-600" : "text-slate-700")}>
              {money(row.profit)} ({row.marginPct.toFixed(1)}%)
            </span>
          </button>
        );
      }

      function findQuotationByNo(quotationNo, quotations) {
        const k = String(quotationNo || "").trim().toLowerCase();
        if (!k) return null;
        return quotations.find((q) => String(q.quotation_no || "").toLowerCase() === k) || null;
      }

      function canLinkQuotationToJob(q, jobModal, jobs) {
        if (!q) return { ok: false, key: "quotationNotFound" };
        if (q.status !== "Accepted") return { ok: false, key: "quotationNotAccepted" };
        if (jobModal && jobModal.mode === "edit") {
          const job = jobs.find((j) => j.id === jobModal.id);
          if (job && quotationLinksJob(q, job.job_no)) return { ok: true };
        }
        return { ok: true };
      }

      function fieldsFromQuotation(q) {
        return {
          quotation_no: q.quotation_no,
          client_id: String(q.client_id || ""),
          company: q.company,
          amount: String(q.amount ?? ""),
          currency: q.currency || "USD",
          description: q.description || ""
        };
      }

      function quotationAccessibleInCurrentScope(q, allowedRegions) {
        if (!q) return false;
        const rg = q.region || "HK";
        return allowedRegions.includes("ALL") || allowedRegions.includes(rg);
      }

      function findLinkedQuotationForJob(job, quotations) {
        const all = findLinkedQuotationsForJob(job, quotations);
        return all[0] || null;
      }

      function findLinkedQuotationsForJob(job, quotations) {
        if (!job) return [];
        const jobNo = String(job.job_no || "").trim().toLowerCase();
        const matched = new Map();
        jobQuotationNos(job).forEach((qNo) => {
          const q = quotations.find((x) => String(x.quotation_no || "").trim().toLowerCase() === qNo.toLowerCase());
          if (q) matched.set(q.id, q);
        });
        quotations.forEach((q) => {
          if (job.id != null && quotationJobIds(q).includes(job.id)) matched.set(q.id, q);
          else if (jobNo && quotationJobNos(q).some((n) => n.toLowerCase() === jobNo)) matched.set(q.id, q);
        });
        return [...matched.values()];
      }

      function syncJobsOnQuotationSave(list, jobId, jobNo, quotationNo, oldJobNo, oldQuotationNo) {
        const nextJobNo = String(jobNo || "").trim();
        const nextQuotationNo = String(quotationNo || "").trim();
        const prevJobNo = String(oldJobNo || "").trim();
        const prevQuotationNo = String(oldQuotationNo || "").trim();
        if (!nextJobNo && !prevJobNo) return list;
        return list.map((j) => {
          if (prevJobNo && j.job_no === prevJobNo && prevQuotationNo && (!nextJobNo || j.job_no !== nextJobNo)) {
            return removeQuotationFromJob(j, prevQuotationNo);
          }
          if (nextJobNo && j.job_no === nextJobNo && nextQuotationNo) {
            return addQuotationToJob(j, nextQuotationNo);
          }
          return j;
        });
      }

      function resolveJsPDF() {
        if (window.jspdf && window.jspdf.jsPDF) return window.jspdf.jsPDF;
        if (typeof window.jsPDF === "function") return window.jsPDF;
        return null;
      }

      let pdfLibLoadPromise = null;
      function ensurePdfLibs() {
        if (resolveJsPDF()) return Promise.resolve(true);
        if (pdfLibLoadPromise) return pdfLibLoadPromise;
        pdfLibLoadPromise = new Promise(function (resolve) {
          var urls = ["https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js"];
          var autoUrls = ["https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.4/dist/jspdf.plugin.autotable.min.js"];
          function loadScript(src) {
            return new Promise(function (res, rej) {
              var s = document.createElement("script");
              s.src = src;
              s.onload = function () { res(); };
              s.onerror = function () { rej(new Error(src)); };
              document.head.appendChild(s);
            });
          }
          function tryLoad(list, i) {
            if (i >= list.length) return Promise.reject(new Error("PDF library not loaded"));
            return loadScript(list[i]).catch(function () { return tryLoad(list, i + 1); });
          }
          tryLoad(urls, 0).then(function () { return tryLoad(autoUrls, 0); }).then(function () {
            resolve(!!resolveJsPDF());
          }).catch(function () { resolve(false); });
        });
        return pdfLibLoadPromise;
      }

      function exportOngoingPdfReport({ rows, totals, title, t, moneyFn }) {
        const JsPDF = resolveJsPDF();
        if (!JsPDF) { alert("PDF library not loaded"); return; }
        const doc = new JsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
        doc.setFontSize(13);
        doc.text(title, 40, 32);
        const head = [[
          t("ongoingColJobNo"), t("ongoingColNature") + " " + t("ongoingNatureHint"), t("ongoingColCustomer"), t("ongoingColProject"),
          t("ongoingColPo"), t("ongoingColStart"), t("ongoingColCurrency"), t("ongoingColPoAmt"),
          t("ongoingColBilled"), t("ongoingColBalance"), t("ongoingColBalanceUsd"), t("ongoingColRemarks")
        ]];
        const body = rows.map((j) => [
          j.job_no, j.nature, j.company, j.description || "", j.customer_po || "", j.start_date || "",
          j.arCurrency || "", j.poAmtLabel || moneyFn(j.poAmt), moneyFn(j.billedAmt), moneyFn(j.poBalance), moneyFn(j.balanceUsd), j.remarks || ""
        ]);
        doc.autoTable({
          head, body, startY: 42,
          styles: { fontSize: 7.5, cellPadding: 3 },
          headStyles: { fillColor: [218, 227, 243], textColor: [30, 41, 59], fontStyle: "bold" },
          columnStyles: {
            7: { halign: "right" }, 8: { halign: "right" }, 9: { halign: "right" }, 10: { halign: "right" }
          }
        });
        const summaryY = doc.lastAutoTable.finalY + 16;
        doc.autoTable({
          body: [
            [t("ongoingTotalPoUsd"), moneyFn(totals.totalPoUsd)],
            [t("ongoingTotalMaintUsd"), moneyFn(totals.Maintenance)],
            [t("ongoingTotalServiceUsd"), moneyFn(totals.Service)],
            [t("ongoingTotalProjectUsd"), moneyFn(totals.Project)],
            [t("ongoingTotalTradeUsd"), moneyFn(totals.Trade)]
          ],
          startY: summaryY,
          margin: { left: 420 },
          styles: { fontSize: 9, cellPadding: 4 },
          columnStyles: { 0: { fontStyle: "bold", cellWidth: 220 }, 1: { halign: "right", cellWidth: 100 } },
          theme: "plain"
        });
        const safeName = title.replace(/[^\w\u4e00-\u9fff-]+/g, "_").slice(0, 80);
        doc.save(safeName + ".pdf");
      }

      function exportMonthlyReportPdf({ series, title, t, moneyFn }) {
        const JsPDF = resolveJsPDF();
        if (!JsPDF) { alert("PDF library not loaded"); return; }
        const doc = new JsPDF({ orientation: "landscape", unit: "pt", format: "a4" });
        doc.setFontSize(13);
        doc.text(title, 40, 32);
        const monthCount = series.length;
        const margin = { left: 28, right: 28 };
        const pageWidth = doc.internal.pageSize.getWidth();
        const tableWidth = pageWidth - margin.left - margin.right;
        const labelWidth = 132;
        const monthColWidth = (tableWidth - labelWidth) / monthCount;
        const fmtAmt = (v) => (v ? moneyFn(v) : "");
        const head = [
          [""].concat(series.map((m) => m.header)),
          [""].concat(series.map((m) => m.currency))
        ];
        const body = [
          [t("monthlyPoReceived")].concat(series.map((m) => fmtAmt(m.poReceived))),
          [t("monthlyInvoiceIssued")].concat(series.map((m) => fmtAmt(m.invoiceIssued))),
          [t("monthlyArExpected")].concat(series.map((m) => fmtAmt(m.arExpected)))
        ];
        const columnStyles = { 0: { cellWidth: labelWidth, halign: "left" } };
        for (let i = 1; i <= monthCount; i++) {
          columnStyles[i] = { cellWidth: monthColWidth, halign: "right" };
        }
        doc.autoTable({
          head,
          body,
          startY: 42,
          theme: "grid",
          margin,
          tableWidth,
          styles: {
            fontSize: 7.5,
            cellPadding: { top: 4, right: 5, bottom: 4, left: 5 },
            lineWidth: 0.1,
            lineColor: [190, 198, 210],
            valign: "middle"
          },
          headStyles: {
            fillColor: [218, 227, 243],
            textColor: [30, 41, 59],
            fontStyle: "bold",
            halign: "center"
          },
          columnStyles,
          didParseCell: function (data) {
            if (data.column.index === 0) {
              data.cell.styles.halign = "left";
              if (data.section === "body") data.cell.styles.fontStyle = "bold";
              return;
            }
            if (data.section === "head") {
              data.cell.styles.halign = "center";
              return;
            }
            if (data.section === "body") {
              data.cell.styles.halign = "right";
              if (data.row.index === 2) data.cell.styles.textColor = [185, 28, 28];
            }
          }
        });
        const safeName = title.replace(/[^\w\u4e00-\u9fff-]+/g, "_").slice(0, 80);
        doc.save(safeName + ".pdf");
      }

      function emptyQuotation() {
        return { quotation_no: "", client_id: "", company: "", quotation_date: "", valid_until: "", amount: "", currency: "USD", status: "Draft", description: "", job_no: "" };
      }

      function emptyVendor() { return { vendor_no: "", name: "", email: "", bank: "", swift_code: "", charge: "OUR", contact: "", phone: "", address: "", finance_contact: "", finance_email: "" }; }

      function ensureVendorNumbers(list) {
        let n = 1;
        return list.map((v) => {
          if (v.vendor_no) return v;
          const no = "VEND-" + String(n++).padStart(4, "0");
          return { ...v, vendor_no: no };
        });
      }

      const LIST_DEFAULT_SORT = {
        clients: { field: "customer_no", dir: "asc" },
        quotation: { field: "quotation_no", dir: "asc" },
        job: { field: "job_no", dir: "asc" },
        vendors: { field: "vendor_no", dir: "asc" },
        ar: { field: "invoice_no", dir: "asc" },
        ap: { field: "invoice_no", dir: "asc" },
      };

      function sortRecords(list, field, dir) {
        const mult = dir === "desc" ? -1 : 1;
        return [...list].sort((a, b) => {
          const sa = a[field] == null ? "" : String(a[field]).trim();
          const sb = b[field] == null ? "" : String(b[field]).trim();
          if (!sa && !sb) return 0;
          if (!sa) return 1 * mult;
          if (!sb) return -1 * mult;
          const na = Number(sa.replace(/,/g, ""));
          const nb = Number(sb.replace(/,/g, ""));
          if (!Number.isNaN(na) && !Number.isNaN(nb) && /^-?[\d,]+(\.\d+)?$/.test(sa) && /^-?[\d,]+(\.\d+)?$/.test(sb)) {
            return mult * (na - nb);
          }
          return mult * sa.localeCompare(sb, undefined, { numeric: true, sensitivity: "base" });
        });
      }

      function emptyApForm() {
        return { job_id: "", job_no: "", payee_type: "Vendor", company_name: "", si_id: "", invoice_no: "", invoice_date: "", invoice_received_date: "", due_date: "", currency: "USD", amount: "", exchange_rate: "", base_amount: "", override_reason: "", pay_date: "", payment_status: "Awaiting Payment", remarks: "", bank: "", charge: "", swift_code: "", payment_advice_email: "", payment_stage: "" };
      }

      function patchApFromPayeeMaster(companyName, vendorList) {
        const v = vendorList.find((x) => x.name === companyName);
        if (!v) return { company_name: companyName || "", bank: "", swift_code: "", payment_advice_email: "", charge: "" };
        return {
          company_name: v.name || "",
          bank: v.bank || "",
          swift_code: v.swift_code || "",
          payment_advice_email: v.email || "",
          charge: v.charge || "OUR"
        };
      }

      function Modal({ title, children, onClose, wide }) {
        return (
          <div className="fixed inset-0 z-50 bg-black/40 p-4 flex items-center justify-center" onMouseDown={(e) => { if (e.target === e.currentTarget) onClose(); }}>
            <div className={"w-full rounded-xl bg-white border border-slate-200 shadow-xl " + (wide ? "max-w-5xl" : "max-w-4xl")} onMouseDown={(e) => e.stopPropagation()} onClick={(e) => e.stopPropagation()}>
              <div className="px-5 py-4 border-b flex justify-between items-center">
                <h3 className="font-semibold">{title}</h3>
                <button onClick={onClose} className="text-slate-500 hover:text-slate-900">×</button>
              </div>
              <div className="p-5 max-h-[78vh] overflow-y-auto">{children}</div>
            </div>
          </div>
        );
      }

      function Field({ label, children }) {
        return (
          <label className="block">
            <span className="block mb-1 text-xs font-medium text-slate-600">{label}</span>
            {children}
          </label>
        );
      }

      function Input(props) {
        return <input {...props} className={"w-full rounded-lg border border-slate-300 px-3 py-2 text-sm " + (props.className || "")} />;
      }

      function Select(props) {
        return <select {...props} className={"w-full rounded-lg border border-slate-300 px-3 py-2 text-sm " + (props.className || "")} />;
      }

      function SearchableSelect({ value, options, onChange, placeholder, required, noResultsText }) {
        const listId = React.useMemo(() => "erp-searchable-" + Math.random().toString(36).slice(2), []);
        const selected = options.find((opt) => String(opt.value) === String(value));
        const [query, setQuery] = useState(selected ? selected.label : "");

        useEffect(() => {
          setQuery(selected ? selected.label : "");
        }, [selected ? selected.label : "", value]);

        const filtered = useMemo(() => {
          const q = String(query || "").trim().toLowerCase();
          if (!q) return options;
          return options.filter((opt) => String(opt.searchText || opt.label || "").toLowerCase().includes(q));
        }, [options, query]);

        function commit(nextQuery) {
          const normalized = String(nextQuery || "").trim().toLowerCase();
          if (!normalized) {
            setQuery("");
            onChange("", null);
            return;
          }
          const exact = options.find((opt) => {
            const label = String(opt.label || "").toLowerCase();
            const searchText = String(opt.searchText || "").toLowerCase();
            return label === normalized || searchText === normalized;
          });
          if (exact) {
            setQuery(exact.label);
            onChange(String(exact.value), exact);
          }
        }

        return (
          <div className="relative">
            <input
              list={listId}
              value={query}
              onChange={(e) => {
                const next = e.target.value;
                setQuery(next);
                const exact = options.find((opt) => String(opt.label || "").toLowerCase() === next.toLowerCase());
                if (exact) onChange(String(exact.value), exact);
                else if (!next.trim()) onChange("", null);
              }}
              onBlur={() => commit(query)}
              placeholder={placeholder}
              className="w-full rounded-lg border border-slate-300 px-3 py-2 text-sm"
            />
            <datalist id={listId}>
              {filtered.map((opt) => (
                <option key={opt.value} value={opt.label}>
                  {opt.searchText && opt.searchText !== opt.label ? opt.searchText : opt.label}
                </option>
              ))}
            </datalist>
            {required ? <input tabIndex={-1} value={value || ""} onChange={() => {}} required className="absolute opacity-0 pointer-events-none h-0 w-0" aria-hidden="true" /> : null}
            {!filtered.length && query.trim() ? <p className="text-[10px] text-slate-400 mt-1">{noResultsText || "No matching results"}</p> : null}
          </div>
        );
      }

      function ActionButtons({ onView, onEdit, onDelete, viewLabel, editLabel, deleteLabel, canEdit = true, canDelete = false, extra }) {
        return (
          <div className="flex gap-1 flex-wrap">
            <button type="button" onClick={(e) => { e.stopPropagation(); onView(); }} className="px-2 py-1 text-xs rounded bg-slate-100 hover:bg-slate-200">{viewLabel}</button>
            {canEdit && <button type="button" onClick={(e) => { e.stopPropagation(); onEdit(); }} className="px-2 py-1 text-xs rounded bg-amber-100 text-amber-800 hover:bg-amber-200">{editLabel}</button>}
            {canDelete && onDelete && <button type="button" onClick={(e) => { e.stopPropagation(); onDelete(); }} className="px-2 py-1 text-xs rounded bg-red-100 text-red-800 hover:bg-red-200">{deleteLabel || "Delete"}</button>}
            {extra}
          </div>
        );
      }

      function DetailActions({ onEdit, onDelete, editLabel, deleteLabel, canEdit = true, canDelete = false, extra, children }) {
        const hasActions = (canEdit && onEdit) || (canDelete && onDelete) || extra || children;
        if (!hasActions) return null;
        return (
          <div className="flex flex-wrap items-center gap-2 pt-4 mt-4 border-t border-slate-200">
            {canEdit && onEdit && (
              <button type="button" onClick={onEdit} className="px-3 py-1.5 text-xs rounded-lg bg-amber-100 text-amber-900 hover:bg-amber-200 font-medium">{editLabel}</button>
            )}
            {canDelete && onDelete && (
              <button type="button" onClick={onDelete} className="px-3 py-1.5 text-xs rounded-lg bg-red-100 text-red-900 hover:bg-red-200 font-medium">{deleteLabel}</button>
            )}
            {extra}
            {children}
          </div>
        );
      }

      function LinkBtn({ children, onClick }) {
        return <button type="button" onClick={(e) => { e.stopPropagation(); onClick(); }} className="text-blue-600 hover:underline font-medium">{children}</button>;
      }

      function ListToolbar({ t, mode, clients, jobs, viewMode, setViewMode, clientFilter, setClientFilter, jobFilter, setJobFilter, statusFilter, setStatusFilter, typeFilter, setTypeFilter, searchQ, setSearchQ, lang }) {
        const showSearch = ["job", "ar", "ap", "clients", "quotation", "vendors"].includes(mode);
        const showDisplayMode = ["job", "ar", "ap"].includes(mode);
        const searchHintKey = { ar: "searchHint", ap: "searchHint", job: "searchHintJob", clients: "searchHintClients", quotation: "searchHintQuotation", vendors: "searchHintVendors" }[mode] || "filterClientHint";
        const clientOptions = useMemo(() => [
          { value: "all", label: t("all"), searchText: t("all") },
          ...sortRecords(clients || [], "customer_no", "asc").map((c) => ({
            value: mode === "ar" ? (c.invoice_title || c.company) : c.company,
            label: mode === "ar"
              ? `${c.customer_no} · ${c.invoice_title || c.company}${c.bu_no ? " · BU " + c.bu_no : ""}`
              : `${c.customer_no} · ${c.company}${c.bu_no ? " · BU " + c.bu_no : ""}`,
            searchText: (mode === "ar"
              ? [c.customer_no, c.invoice_title, c.company, c.primary_contact, c.bu_no]
              : [c.customer_no, c.company, c.primary_contact, c.bu_no]
            ).filter(Boolean).join(" ")
          }))
        ], [clients, t, mode]);
        const jobOptions = useMemo(() => [
          { value: "all", label: t("all"), searchText: t("all") },
          ...sortRecords(jobs || [], "job_no", "asc").map((j) => ({
            value: String(j.id),
            label: `${j.job_no} · ${j.company}`,
            searchText: [j.job_no, j.company, j.job_type, j.customer_po].filter(Boolean).join(" ")
          }))
        ], [jobs, t]);
        return (
          <div className="p-4 border-b flex flex-wrap gap-3 items-center bg-slate-50/80">
            {(mode === "job" || mode === "ar" || mode === "ap") && (
              <label className="text-sm flex items-center gap-2">
                <span className="text-slate-600">{t("filterClient")}:</span>
                <div className="w-72">
                  <SearchableSelect
                    value={clientFilter}
                    options={clientOptions}
                    placeholder={t("selectClientHint")}
                    noResultsText={t("noMatchFound")}
                    onChange={(nextValue, option) => setClientFilter(nextValue || (option ? option.value : "all") || "all")}
                  />
                </div>
              </label>
            )}
            {mode === "job" && (
              <label className="text-sm flex items-center gap-2">
                <span className="text-slate-600">{t("filterStatus")}:</span>
                <Select value={statusFilter} onChange={(e) => setStatusFilter(e.target.value)} className="w-40 bg-white">
                  <option value="all">{t("all")}</option>
                  {JOB_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                </Select>
              </label>
            )}
            {mode === "job" && (
              <label className="text-sm flex items-center gap-2">
                <span className="text-slate-600">{t("filterType")}:</span>
                <Select value={typeFilter} onChange={(e) => setTypeFilter(e.target.value)} className="w-48 bg-white">
                  <option value="all">{t("all")}</option>
                  {JOB_TYPES.map((s) => <option key={s} value={s}>{s}</option>)}
                </Select>
              </label>
            )}
            {mode === "ap" && (
              <label className="text-sm flex items-center gap-2">
                <span className="text-slate-600">{t("filterJob")}:</span>
                <div className="w-72">
                  <SearchableSelect
                    value={jobFilter}
                    options={jobOptions}
                    placeholder={t("selectJobHint")}
                    noResultsText={t("noMatchFound")}
                    onChange={(nextValue, option) => setJobFilter(nextValue || (option ? option.value : "all") || "all")}
                  />
                </div>
              </label>
            )}
            {showSearch && setSearchQ && (
              <label className="text-sm flex items-center gap-2 flex-1 min-w-[12rem]">
                <span className="text-slate-600 shrink-0">{t("auditSearch")}:</span>
                <Input value={searchQ || ""} onChange={(e) => setSearchQ(e.target.value)} placeholder={t("searchPlaceholder")} className="bg-white" />
              </label>
            )}
            {showDisplayMode && (
              <label className="text-sm flex items-center gap-2">
                <span className="text-slate-600">{t("displayMode")}:</span>
                <Select value={viewMode} onChange={(e) => setViewMode(e.target.value)} className="w-52 bg-white">
                  <option value="list">{t("listView")}</option>
                  <option value="grouped">{mode === "ap" ? t("groupByVendor") : t("groupByClient")}</option>
                </Select>
              </label>
            )}
            <span className="text-xs text-slate-500">
              {showSearch ? t(searchHintKey) : mode === "ap" && jobFilter !== "all" ? t("filterJobHint") : t("filterClientHint")}
              {showDisplayMode ? " · " + t("groupHint") : ""}
            </span>
          </div>
        );
      }

      const initialClients = [];
      const initialJobs = [];
      const initialQuotations = [];
      const initialVendors = [];
      const initialAr = [];
      const initialAp = [];

      const initialApprovals = [
        { id: 1, doc: "PO-2026-4411 / IT Hardware", requester: "Benson", status: "Pending CFO" },
        { id: 2, doc: "BILL-2026-2289 / Marketing Agency", requester: "Amy", status: "Pending Finance" },
        { id: 3, doc: "Expense Batch / Travel Reimbursement", requester: "Sales Team", status: "Over SLA" }
      ];

      function normalizeStoredUsers(loaded) {
        if (!Array.isArray(loaded) || !loaded.length) return initialUsers.map((u) => ({ ...u }));
        const users = loaded.map((u) => {
          let email = String(u.email || "").trim();
          if (u.role === "root" && (email.toLowerCase() === "root@erp.local" || email.toLowerCase() === "root@example.com")) email = "root";
          return {
            ...u,
            email,
            permissions: u.role === "root" ? fullPermissions() : normalizePermissions(u.permissions),
            allowed_regions: u.role === "root" ? ["ALL"] : (u.allowed_regions && u.allowed_regions.length ? u.allowed_regions : ["HK"])
          };
        });
        if (!users.some((u) => u.role === "root")) {
          const rootSeed = initialUsers.find((u) => u.role === "root");
          if (rootSeed) users.unshift({ ...rootSeed });
        }
        return users;
      }

      function getErpUsers() {
        const loaded = loadJson("erp_users", null);
        const users = normalizeStoredUsers(loaded);
        if (!loaded || JSON.stringify(loaded) !== JSON.stringify(users)) saveErpStorage("erp_users", users);
        return users;
      }

      async function fetchAllCloudUsersForLogin() {
        const batches = [];
        try {
          const statusRes = await fetch("/api/sync/status");
          if (statusRes.ok) {
            const status = await statusRes.json();
            if (status.entity_count > 0) {
              const fullRes = await fetch("/api/sync/full");
              if (fullRes.ok) {
                const full = await fullRes.json();
                const syncUsers = full.entities && full.entities.users;
                if (Array.isArray(syncUsers) && syncUsers.length) batches.push(syncUsers);
              }
            }
          }
        } catch { /* ignore */ }
        try {
          const cfgRes = await fetch("/api/sync/client-config");
          if (cfgRes.ok) {
            const cfg = await cfgRes.json();
            if (cfg.sync_key) {
              const backupRes = await fetch("/api/backup", {
                headers: { "X-ERP-Sync-Key": cfg.sync_key }
              });
              if (backupRes.ok) {
                const body = await backupRes.json();
                const backupUsers = body.payload && body.payload.data && body.payload.data.users;
                if (Array.isArray(backupUsers) && backupUsers.length) batches.push(backupUsers);
              }
            }
          }
        } catch { /* ignore */ }
        if (!batches.length) return null;
        let merged = [];
        batches.forEach((batch) => { merged = mergeUsersForLogin(merged, batch); });
        return merged;
      }

      function mergeUsersForLogin(localUsers, cloudUsers) {
        const byId = new Map((normalizeStoredUsers(localUsers) || []).map((u) => [u.id, { ...u }]));
        (cloudUsers || []).forEach((u) => {
          if (!u || u.id == null) return;
          byId.set(u.id, { ...u });
        });
        return normalizeStoredUsers(Array.from(byId.values()));
      }

      async function getUsersForLogin() {
        const local = getErpUsers();
        const cloudUsers = await fetchAllCloudUsersForLogin();
        if (!cloudUsers) return local;
        const merged = mergeUsersForLogin(local, cloudUsers);
        saveErpStorage("erp_users", merged);
        return merged;
      }

      function findUserForLogin(users, id, pw) {
        const key = String(id || "").trim().toLowerCase();
        const pwStr = String(pw || "");
        if (!key || !pwStr) return null;
        return users.find((u) => {
          if (String(u.password || "") !== pwStr) return false;
          const email = String(u.email || "").trim().toLowerCase();
          const name = String(u.name || "").trim().toLowerCase();
          if (email === key || name === key) return true;
          if (u.role === "root" && (key === "root" || key === "root@erp.local" || key === "administrator")) return true;
          return false;
        }) || null;
      }

      const initialUsers = [
        { id: 1, name: "Administrator", email: "root", password: "root", role: "root", is_active: true, permissions: fullPermissions(), allowed_regions: ["ALL"] },
        {
          id: 2, name: "Finance User", email: "finance@erp.local", password: "demo", role: "user", is_active: true, allowed_regions: ["HK"],
          permissions: {
            dashboard: { view: true, add: false, edit: false, delete: false, export: false, import: false },
            management_report: { view: true, add: false, edit: false, delete: false, export: false, import: false },
            overview: { view: true, add: false, edit: false, delete: false, export: true, import: false },
            clients: { view: true, add: false, edit: false, delete: false, export: true, import: true },
            quotation: { view: true, add: true, edit: true, delete: false, export: true, import: true },
            job: { view: true, add: true, edit: true, delete: false, export: true, import: true },
            ongoing: { view: true, add: false, edit: false, delete: false, export: true, import: false },
            ar: { view: true, add: true, edit: true, delete: false, export: true, import: true },
            ap: { view: true, add: true, edit: true, delete: false, export: true, import: true },
            monthly_report: { view: true, add: false, edit: true, delete: false, export: true, import: false },
            vendors: { view: true, add: false, edit: false, delete: false, export: true, import: true },
            settings: { view: true, add: false, edit: false, delete: false, export: false, import: false }
          }
        }
      ];

      const initialAuditLogs = [];

      function loadJson(key, fallback) {
        try {
          const raw = localStorage.getItem(key);
          return raw ? JSON.parse(raw) : fallback;
        } catch {
          return fallback;
        }
      }

      function saveErpStorage(key, data) {
        try {
          localStorage.setItem(key, JSON.stringify(data));
          return true;
        } catch {
          return false;
        }
      }

      const ERP_BUSINESS_CACHE_KEYS = [
        "erp_clients", "erp_jobs", "erp_quotations", "erp_vendors",
        "erp_ar", "erp_ap", "erp_monthly_po_lines", "erp_monthly_ar_lines", "erp_monthly_ar_expected"
      ];

      function isCloudOnlyPreferred() {
        return localStorage.getItem("erp_cloud_only") === "1" || !!localStorage.getItem("erp_sync_key");
      }

      function clearLocalBusinessCache() {
        ERP_BUSINESS_CACHE_KEYS.forEach((k) => localStorage.removeItem(k));
      }

      function initialBusinessList(key, normalize, fallback) {
        if (isCloudOnlyPreferred()) return normalize ? normalize([]) : [];
        const raw = loadJson(key, fallback);
        return normalize ? normalize(raw) : raw;
      }

      function OverviewDonutChart({ inTotal, outTotal, currency, t }) {
        const total = inTotal + outTotal;
        const size = 220;
        const r = 72;
        const stroke = 32;
        const cx = size / 2;
        const cy = size / 2;
        const circ = 2 * Math.PI * r;
        const inPct = total > 0 ? inTotal / total : 0;
        const outPct = total > 0 ? outTotal / total : 0;
        const inLen = total > 0 ? circ * inPct : 0;
        const outLen = total > 0 ? circ * outPct : circ;

        return (
          <div className="flex flex-col sm:flex-row items-center justify-center gap-8 py-4">
            <div className="relative shrink-0" style={{ width: size, height: size }}>
              <svg width={size} height={size} viewBox={"0 0 " + size + " " + size} className="-rotate-90">
                <circle cx={cx} cy={cy} r={r} fill="none" stroke="#e2e8f0" strokeWidth={stroke} />
                {total > 0 ? (
                  <>
                    <circle cx={cx} cy={cy} r={r} fill="none" stroke="#2563eb" strokeWidth={stroke} strokeLinecap="butt"
                      strokeDasharray={inLen + " " + (circ - inLen)} />
                    <circle cx={cx} cy={cy} r={r} fill="none" stroke="#ea580c" strokeWidth={stroke} strokeLinecap="butt"
                      strokeDasharray={outLen + " " + (circ - outLen)} strokeDashoffset={-inLen} />
                  </>
                ) : (
                  <circle cx={cx} cy={cy} r={r} fill="none" stroke="#cbd5e1" strokeWidth={stroke} strokeDasharray="4 8" />
                )}
              </svg>
              <div className="absolute inset-0 flex flex-col items-center justify-center text-center px-4">
                {total > 0 ? (
                  <>
                    <p className="text-[10px] text-slate-500 uppercase tracking-wide">{t("overviewNet")}</p>
                    <p className={"text-lg font-bold tabular-nums " + (inTotal >= outTotal ? "text-emerald-700" : "text-amber-700")}>{money(inTotal - outTotal)}</p>
                    <p className="text-[10px] text-slate-400 mt-0.5">{currency}</p>
                  </>
                ) : (
                  <p className="text-xs text-slate-400">{t("overviewNoData")}</p>
                )}
              </div>
            </div>
            <div className="space-y-3 text-sm min-w-[10rem]">
              <div className="flex items-center gap-2">
                <span className="w-3 h-3 rounded-full bg-blue-600 shrink-0" />
                <div className="flex-1">
                  <p className="text-slate-500 text-xs">{t("overviewInTotal")}</p>
                  <p className="font-bold text-blue-700 tabular-nums">{money(inTotal)} <span className="text-xs font-normal text-slate-400">{currency}</span></p>
                </div>
              </div>
              <div className="flex items-center gap-2">
                <span className="w-3 h-3 rounded-full bg-orange-600 shrink-0" />
                <div className="flex-1">
                  <p className="text-slate-500 text-xs">{t("overviewOutTotal")}</p>
                  <p className="font-bold text-orange-600 tabular-nums">{money(outTotal)} <span className="text-xs font-normal text-slate-400">{currency}</span></p>
                </div>
              </div>
              {total > 0 && (
                <p className="text-[10px] text-slate-400 pt-1 border-t border-slate-100">
                  {Math.round(inPct * 100)}% / {Math.round(outPct * 100)}%
                </p>
              )}
            </div>
          </div>
        );
      }

      function isArPaid(r) {
        return !!(r && (r.payment_status === "Paid" || r.payment_received_date));
      }

      function arPaymentDate(r) {
        if (r.payment_received_date) return r.payment_received_date;
        if (r.payment_status === "Paid") return r.invoice_date || r.due_date || "";
        return "";
      }

      function MonthlyDetailGrid({ grid, baseCurrency, canEdit, moneyFn, t, onOtherCurrencyChange }) {
        const { months, maxRows } = grid;
        return (
          <div className="erp-list-scroll erp-list-scroll--ongoing">
            <table className="text-xs min-w-[1100px]">
              <thead>
                <tr className="bg-[#dae3f3] text-slate-800">
                  {months.map((m, i) => (
                    <React.Fragment key={"mh-" + m.monthIndex}>
                      {i > 0 && <th className="w-2 border-0 bg-white"></th>}
                      <th colSpan={4} className="p-2 border border-slate-300 text-center font-semibold whitespace-nowrap">{m.header}</th>
                    </React.Fragment>
                  ))}
                </tr>
                <tr className="bg-slate-50 text-slate-600">
                  {months.map((m, i) => (
                    <React.Fragment key={"sh-" + m.monthIndex}>
                      {i > 0 && <th className="w-2 border-0 bg-white"></th>}
                      <th className="p-1.5 border border-slate-300 font-medium whitespace-nowrap">{t("monthlyColJobNo")}</th>
                      <th className="p-1.5 border border-slate-300 font-medium whitespace-nowrap">{t("monthlyColOtherCurrency")}</th>
                      <th className="p-1.5 border border-slate-300 font-medium text-right whitespace-nowrap">{t("monthlyColSourceAmount")}</th>
                      <th className="p-1.5 border border-slate-300 font-medium text-right whitespace-nowrap">{baseCurrency}</th>
                    </React.Fragment>
                  ))}
                </tr>
              </thead>
              <tbody>
                {Array.from({ length: maxRows }, (_, rowIdx) => (
                  <tr key={"r-" + rowIdx}>
                    {months.map((m, i) => {
                      const row = m.rows[rowIdx];
                      return (
                        <React.Fragment key={"c-" + m.monthIndex + "-" + rowIdx}>
                          {i > 0 && <td className="border-0 bg-white"></td>}
                          <td className="p-1.5 border border-slate-200 whitespace-nowrap">{row ? row.job_no : ""}</td>
                          <td className="p-1.5 border border-slate-200">
                            {row && canEdit ? (
                              <input
                                type="text"
                                className="w-full min-w-[3.5rem] text-xs border border-slate-300 rounded px-1 py-0.5 uppercase"
                                value={row.other_currency || ""}
                                onChange={(e) => onOtherCurrencyChange(row.id, e.target.value.toUpperCase())}
                              />
                            ) : (row ? row.other_currency : "")}
                          </td>
                          <td className="p-1.5 border border-slate-200 text-right tabular-nums whitespace-nowrap">{row ? moneyFn(row.source_amount ?? 0) : ""}</td>
                          <td className="p-1.5 border border-slate-200 text-right tabular-nums whitespace-nowrap">{row ? moneyFn(row.base_amount || 0) : ""}</td>
                        </React.Fragment>
                      );
                    })}
                  </tr>
                ))}
                <tr className="font-semibold bg-slate-50">
                  {months.map((m, i) => (
                    <React.Fragment key={"t-" + m.monthIndex}>
                      {i > 0 && <td className="border-0 bg-white"></td>}
                      <td className="p-1.5 border border-slate-300"></td>
                      <td className="p-1.5 border border-slate-300"></td>
                      <td className="p-1.5 border border-slate-300">{t("monthlyTotal")}</td>
                      <td className="p-1.5 border border-slate-300 text-right tabular-nums border-t-2 border-slate-500">{m.total ? moneyFn(m.total) : "-"}</td>
                    </React.Fragment>
                  ))}
                </tr>
              </tbody>
            </table>
          </div>
        );
      }

      function MonthlyReportBarChart({ series, maxBar, t, moneyFn }) {
        const barAreaPx = 140;
        function renderMonth(p) {
          const poH = p.poReceived > 0 ? Math.max(4, Math.round((p.poReceived / maxBar) * barAreaPx)) : 2;
          const invH = p.invoiceIssued > 0 ? Math.max(4, Math.round((p.invoiceIssued / maxBar) * barAreaPx)) : 2;
          const arH = p.arExpected > 0 ? Math.max(4, Math.round((p.arExpected / maxBar) * barAreaPx)) : 2;
          return (
            <div key={p.key} className="flex-1 flex flex-col items-center min-w-[3.25rem]">
              <div className="w-full flex items-end justify-center gap-0.5" style={{ height: barAreaPx }}>
                <div title={t("monthlyChartPo") + ": " + moneyFn(p.poReceived)} className="w-[30%] max-w-5 rounded-t-sm" style={{ height: poH, opacity: p.poReceived ? 1 : 0.15, background: "linear-gradient(180deg, #60a5fa 0%, #2563eb 100%)" }} />
                <div title={t("monthlyChartInvoice") + ": " + moneyFn(p.invoiceIssued)} className="w-[30%] max-w-5 rounded-t-sm" style={{ height: invH, opacity: p.invoiceIssued ? 1 : 0.15, background: "linear-gradient(180deg, #fdba74 0%, #ea580c 100%)" }} />
                <div title={t("monthlyChartArExpected") + ": " + moneyFn(p.arExpected)} className="w-[30%] max-w-5 rounded-t-sm" style={{ height: arH, opacity: p.arExpected ? 1 : 0.15, background: "linear-gradient(180deg, #cbd5e1 0%, #64748b 100%)" }} />
              </div>
              <span className="text-[8px] sm:text-[9px] text-slate-500 font-medium mt-1.5 leading-tight text-center px-0.5">{p.label}</span>
            </div>
          );
        }
        const firstHalf = series.slice(0, 6);
        const secondHalf = series.slice(6, 12);
        return (
          <div className="space-y-3">
            <div className="flex items-end gap-1 w-full overflow-x-auto pb-1">{firstHalf.map(renderMonth)}</div>
            <div className="flex items-end gap-1 w-full overflow-x-auto pb-1 border-t border-slate-100 pt-3">{secondHalf.map(renderMonth)}</div>
            <div className="flex flex-wrap items-center justify-center gap-4 pt-2 text-xs text-slate-600">
              <span className="inline-flex items-center gap-1.5"><span className="w-3 h-3 rounded-sm bg-blue-600" />{t("monthlyChartPo")}</span>
              <span className="inline-flex items-center gap-1.5"><span className="w-3 h-3 rounded-sm bg-orange-600" />{t("monthlyChartInvoice")}</span>
              <span className="inline-flex items-center gap-1.5"><span className="w-3 h-3 rounded-sm bg-slate-500" />{t("monthlyChartArExpected")}</span>
            </div>
          </div>
        );
      }

      function MonthlyBarChart({ firstHalf, secondHalf, maxBar, t, inLegend, outLegend, h1Label, h2Label, totals, netLabel }) {
        const barAreaPx = 80;
        function renderMonths(series) {
          return series.map((p) => {
            const inH = p.inBase > 0 ? Math.max(6, Math.round((p.inBase / maxBar) * barAreaPx)) : 2;
            const outH = p.outBase > 0 ? Math.max(6, Math.round((p.outBase / maxBar) * barAreaPx)) : 2;
            return (
              <div key={p.key} className="flex-1 flex flex-col items-center min-w-0">
                <div className="w-full flex items-end justify-center gap-0.5" style={{ height: barAreaPx }}>
                  <div
                    title={inLegend + ": " + money(p.inBase)}
                    className="w-[42%] max-w-8 rounded-t-md shadow-sm"
                    style={{
                      height: inH,
                      opacity: p.inBase ? 1 : 0.2,
                      background: "linear-gradient(180deg, #38bdf8 0%, #1d4ed8 100%)"
                    }}
                  />
                  <div
                    title={outLegend + ": " + money(p.outBase)}
                    className="w-[42%] max-w-8 rounded-t-md shadow-sm"
                    style={{
                      height: outH,
                      opacity: p.outBase ? 1 : 0.2,
                      background: "linear-gradient(180deg, #fbbf24 0%, #ea580c 100%)"
                    }}
                  />
                </div>
                <span className="text-[8px] sm:text-[9px] text-slate-500 font-medium mt-1.5 leading-tight text-center px-0.5">{p.label}</span>
                {(p.inBase > 0 || p.outBase > 0) && (
                  <div className="text-[8px] leading-tight mt-0.5 text-center w-full tabular-nums hidden sm:block">
                    {p.inBase > 0 && <span className="text-blue-600 block truncate">{money(p.inBase)}</span>}
                    {p.outBase > 0 && <span className="text-orange-600 block truncate">{money(p.outBase)}</span>}
                  </div>
                )}
              </div>
            );
          });
        }
        return (
          <div className="flex flex-col gap-3">
            <div>
              <p className="text-[10px] font-medium text-slate-400 mb-1.5">{h1Label}</p>
              <div className="flex items-end gap-1 w-full">{renderMonths(firstHalf)}</div>
            </div>
            <div className="border-t border-slate-100 pt-2">
              <p className="text-[10px] font-medium text-slate-400 mb-1.5">{h2Label}</p>
              <div className="flex items-end gap-1 w-full">{renderMonths(secondHalf)}</div>
            </div>
            {totals && (
              <div className="mt-1 pt-3 border-t border-slate-100 grid grid-cols-3 gap-2 text-xs">
                <div className="text-center rounded-lg bg-blue-50/80 px-2 py-2">
                  <p className="text-slate-500 text-[10px] leading-tight">{inLegend}</p>
                  <p className="font-bold text-blue-700 tabular-nums mt-0.5">{money(totals.in)}</p>
                </div>
                <div className="text-center rounded-lg bg-orange-50/80 px-2 py-2">
                  <p className="text-slate-500 text-[10px] leading-tight">{outLegend}</p>
                  <p className="font-bold text-orange-600 tabular-nums mt-0.5">{money(totals.out)}</p>
                </div>
                <div className="text-center rounded-lg bg-slate-50 px-2 py-2">
                  <p className="text-slate-500 text-[10px] leading-tight">{netLabel}</p>
                  <p className={"font-bold tabular-nums mt-0.5 " + (totals.in >= totals.out ? "text-emerald-700" : "text-amber-700")}>{money(totals.in - totals.out)}</p>
                </div>
              </div>
            )}
          </div>
        );
      }

      function GlobalFinancialMap({ stats, lang, currency, onRegionClick, t }) {
        const [hoverId, setHoverId] = useState(null);
        const [selectedId, setSelectedId] = useState(null);
        const [mapScale, setMapScale] = useState(1);
        const [mapFocus, setMapFocus] = useState(null);
        const hoverTimer = useRef(null);
        const mapWheelRef = useRef(null);
        const world = window.ERP_WORLD_MAP;
        const view = worldMapView();
        const maxAr = Math.max(1, ...stats.map((s) => s.arTotal));
        const totalAr = stats.reduce((s, x) => s + x.arTotal, 0);
        const sorted = [...stats].sort((a, b) => b.arTotal - a.arTotal);
        const displayId = selectedId || hoverId;

        const pathsByCountry = useMemo(() => {
          if (!world || !world.items) return {};
          const grouped = {};
          world.items.forEach((item) => {
            const id = String(item.id);
            if (!grouped[id]) grouped[id] = [];
            grouped[id].push(item.d);
          });
          return grouped;
        }, [world]);

        const focus = mapFocus || { x: view.width / 2, y: view.height / 2 };
        const vbW = view.width / mapScale;
        const vbH = view.height / mapScale;
        const vbX = Math.max(0, Math.min(view.width - vbW, focus.x - vbW / 2));
        const vbY = Math.max(0, Math.min(view.height - vbH, focus.y - vbH / 2));
        const viewBoxStr = vbX + " " + vbY + " " + vbW + " " + vbH;

        useEffect(() => {
          const el = mapWheelRef.current;
          if (!el) return;
          const onWheel = (e) => {
            e.preventDefault();
            e.stopPropagation();
            setMapScale((s) => Math.min(4, Math.max(1, +(s + (e.deltaY > 0 ? -0.15 : 0.15)).toFixed(2))));
          };
          el.addEventListener("wheel", onWheel, { passive: false });
          return () => el.removeEventListener("wheel", onWheel);
        }, []);

        function enterHover(id) {
          if (hoverTimer.current) clearTimeout(hoverTimer.current);
          setHoverId(id);
        }
        function leaveHover() {
          hoverTimer.current = setTimeout(() => setHoverId(null), 400);
        }
        function selectRegion(id) {
          if (hoverTimer.current) clearTimeout(hoverTimer.current);
          setSelectedId((prev) => (prev === id ? null : id));
          setHoverId(id);
        }
        function zoomIn(e) { e.stopPropagation(); setMapScale((s) => Math.min(4, +(s + 0.4).toFixed(2))); }
        function zoomOut(e) { e.stopPropagation(); setMapScale((s) => Math.max(1, +(s - 0.4).toFixed(2))); }
        function resetZoom(e) { e.stopPropagation(); setMapScale(1); setMapFocus(null); }
        function focusAsia(e) { e.stopPropagation(); setMapScale(2.5); setMapFocus({ x: 720, y: 205 }); }

        function statForCountry(countryId) {
          return stats.find((s) => REGION_COUNTRY_IDS[s.regionId] === String(countryId));
        }

        function countryFill(countryId) {
          const stat = statForCountry(countryId);
          const isHoverCountry = displayId && REGION_COUNTRY_IDS[displayId] === String(countryId);
          if (!stat) return isHoverCountry ? "#cbd5e1" : "#e8ecf2";
          if (stat.hasOverdueAr) return isHoverCountry ? "#ef4444" : "#fca5a5";
          if (stat.arTotal === 0 && stat.activeJobs === 0) return "#e8ecf2";
          if (stat.arTotal < stat.apTotal) return isHoverCountry ? "#f59e0b" : "#fcd34d";
          const tRatio = stat.arTotal / maxAr;
          const shades = ["#dbeafe", "#bfdbfe", "#93c5fd", "#60a5fa", "#3b82f6", "#2563eb"];
          return shades[Math.min(5, Math.floor(tRatio * 6))];
        }

        function pinColor(stat) {
          if (stat.hasOverdueAr) return "#ef4444";
          if (stat.arTotal === 0 && stat.activeJobs === 0) return "#94a3b8";
          if (stat.arTotal >= stat.apTotal) return "#2563eb";
          return "#f59e0b";
        }

        function pinRadius(stat) {
          if (stat.arTotal === 0 && stat.activeJobs === 0) return 3.5;
          return 4 + Math.min(6, (stat.arTotal / maxAr) * 6);
        }

        const display = stats.find((s) => s.regionId === displayId);

        return (
          <article className="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm flex flex-col h-full">
            <div className="mb-2 flex items-start justify-between gap-2">
              <div>
                <h3 className="font-semibold text-base text-slate-800">{t("dashGlobalMap")}</h3>
                <p className="text-xs text-slate-500 mt-0.5">{t("dashMapExpandHint")}</p>
              </div>
            </div>

            <div
              className="grid grid-cols-1 sm:grid-cols-[minmax(0,1fr)_minmax(0,1.15fr)] gap-3 flex-1 min-h-0"
              onMouseEnter={() => { if (hoverTimer.current) clearTimeout(hoverTimer.current); }}
              onMouseLeave={leaveHover}
            >
              <div className="relative min-h-0">
                <div className="absolute top-1 right-1 z-10 flex flex-wrap gap-0.5 justify-end">
                  <button type="button" onClick={zoomIn} title={t("dashMapZoomIn")} className="w-6 h-6 rounded bg-white/90 border border-slate-200 text-slate-600 text-sm font-bold hover:bg-slate-50 shadow-sm">+</button>
                  <button type="button" onClick={zoomOut} title={t("dashMapZoomOut")} className="w-6 h-6 rounded bg-white/90 border border-slate-200 text-slate-600 text-sm font-bold hover:bg-slate-50 shadow-sm">−</button>
                  <button type="button" onClick={focusAsia} title={t("dashMapZoomAsia")} className="h-6 px-1.5 rounded bg-white/90 border border-slate-200 text-[10px] font-medium text-slate-600 hover:bg-slate-50 shadow-sm">{t("dashMapZoomAsia")}</button>
                  <button type="button" onClick={resetZoom} title={t("dashMapZoomReset")} className="h-6 px-1.5 rounded bg-white/90 border border-slate-200 text-[10px] font-medium text-slate-600 hover:bg-slate-50 shadow-sm">{t("dashMapZoomReset")}</button>
                </div>
                {world ? (
                  <div
                    ref={mapWheelRef}
                    className="rounded-lg overflow-hidden border border-slate-100 bg-slate-50"
                  >
                    <svg viewBox={viewBoxStr} className="w-full h-auto min-h-[9rem] max-h-[11rem]" preserveAspectRatio="xMidYMid meet" aria-label={t("dashGlobalMap")}>
                      <rect x={vbX} y={vbY} width={vbW} height={vbH} fill="#f8fafc" />
                      {Object.entries(pathsByCountry).map(([countryId, paths]) => (
                        <g key={countryId} fill={countryFill(countryId)} stroke="#ffffff" strokeWidth={0.4 / mapScale} strokeLinejoin="round">
                          {paths.map((d, i) => <path key={i} d={d} />)}
                        </g>
                      ))}
                      {stats.map((stat) => {
                        const pinMeta = REGION_MAP_PINS[stat.regionId];
                        if (!pinMeta) return null;
                        const pin = mapPinXY(pinMeta);
                        const r = pinRadius(stat);
                        const color = pinColor(stat);
                        const active = displayId === stat.regionId;
                        const hitR = Math.max(14, r + 8);
                        return (
                          <g
                            key={stat.regionId}
                            className="cursor-pointer"
                            onMouseEnter={() => enterHover(stat.regionId)}
                            onClick={() => selectRegion(stat.regionId)}
                          >
                            <circle cx={pin.x} cy={pin.y} r={hitR} fill="transparent" />
                            {stat.hasOverdueAr && (
                              <g pointerEvents="none">
                                <circle cx={pin.x} cy={pin.y} r={r + 4} fill="#ef4444" opacity="0.35">
                                  <animate attributeName="r" values={(r + 2) + ";" + (r + 9) + ";" + (r + 2)} dur="1.5s" repeatCount="indefinite" />
                                  <animate attributeName="opacity" values="0.4;0.08;0.4" dur="1.5s" repeatCount="indefinite" />
                                </circle>
                              </g>
                            )}
                            <circle cx={pin.x} cy={pin.y} r={r + 1.2} fill="white" pointerEvents="none" />
                            <circle cx={pin.x} cy={pin.y} r={r} fill={color} stroke="white" strokeWidth={1.2 / mapScale} opacity={active ? 1 : 0.95} pointerEvents="none" />
                          </g>
                        );
                      })}
                    </svg>
                  </div>
                ) : (
                  <p className="text-xs text-slate-400 text-center py-6">Map data loading…</p>
                )}
                {display && (
                  <div className="absolute bottom-1 left-1 right-1 z-10 bg-white/95 backdrop-blur-sm border border-slate-200 rounded-lg p-2.5 shadow-md text-xs">
                    <p className="font-semibold text-sm text-slate-800">{REGION_FLAGS[display.regionId]} {regionLabel(display.regionId, lang)}</p>
                    <div className="mt-1.5 flex flex-wrap gap-x-3 gap-y-1 text-slate-600">
                      <span>{t("dashMapArTotal")}: <strong className="text-blue-700">{money(display.arTotal)} {currency}</strong></span>
                      <span>{t("dashMapApTotal")}: <strong className="text-orange-600">{money(display.apTotal)} {currency}</strong></span>
                      <span>{t("dashMapActiveJobs")}: <strong>{display.activeJobs}</strong></span>
                      <span>{t("dashMapNetRegion")}: <strong className={display.arTotal >= display.apTotal ? "text-emerald-600" : "text-amber-600"}>{money(display.arTotal - display.apTotal)}</strong></span>
                    </div>
                    {display.hasOverdueAr && <p className="mt-1.5 text-sm text-red-600 font-medium">{t("dashMapOverdueAlert")}</p>}
                  </div>
                )}
              </div>

              <ul className="space-y-1 flex-1 min-h-0 overflow-y-auto pr-0.5">
                {sorted.map((stat) => {
                  const pct = totalAr > 0 ? Math.round((stat.arTotal / totalAr) * 100) : 0;
                  const barW = Math.max(stat.arTotal > 0 ? 12 : 0, pct);
                  const barCls = stat.hasOverdueAr ? "bg-red-500" : stat.arTotal >= stat.apTotal ? "bg-blue-600" : "bg-amber-500";
                  const active = displayId === stat.regionId;
                  const expanded = selectedId === stat.regionId;
                  return (
                    <li key={stat.regionId}>
                      <div
                        className={"flex items-center gap-2 rounded-lg px-1.5 py-1 cursor-pointer transition-colors " + (active ? "bg-blue-50/80 ring-1 ring-blue-200/60" : "hover:bg-slate-50")}
                        onMouseEnter={() => enterHover(stat.regionId)}
                        onClick={() => selectRegion(stat.regionId)}
                      >
                        <span className={"text-slate-400 shrink-0 transition-transform " + (expanded ? "rotate-90" : "")} aria-hidden="true">▸</span>
                        <span className="text-lg leading-none w-6 text-center shrink-0">{REGION_FLAGS[stat.regionId]}</span>
                        <div className="flex-1 min-w-0">
                          <p className="text-sm font-semibold text-slate-700 truncate">{regionLabel(stat.regionId, lang)}</p>
                          <div className="h-5 bg-slate-100 rounded-full overflow-hidden mt-1">
                            <div className={"h-full rounded-full flex items-center justify-end pr-2 " + barCls} style={{ width: barW + "%", minWidth: stat.arTotal > 0 ? "2.25rem" : 0 }}>
                              {stat.arTotal > 0 && <span className="text-[10px] font-bold text-white tabular-nums">{pct}%</span>}
                            </div>
                          </div>
                        </div>
                        <div className="text-right shrink-0">
                          <p className="text-sm font-bold text-slate-800 tabular-nums">{money(stat.arTotal)}</p>
                          <p className="text-xs text-slate-500">{stat.activeJobs} {t("dashMapActiveJobs")}</p>
                        </div>
                      </div>
                      {expanded && (
                        <div className="ml-7 mr-1 mb-1.5 mt-0.5 p-3 bg-slate-50 border border-slate-100 rounded-lg text-sm">
                          <div className="grid grid-cols-3 gap-2 mb-2">
                            <div>
                              <p className="text-xs text-slate-500">{t("dashMapArTotal")}</p>
                              <p className="font-bold text-blue-700 tabular-nums">{money(stat.arTotal)}</p>
                            </div>
                            <div>
                              <p className="text-xs text-slate-500">{t("dashMapApTotal")}</p>
                              <p className="font-bold text-orange-600 tabular-nums">{money(stat.apTotal)}</p>
                            </div>
                            <div>
                              <p className="text-xs text-slate-500">{t("dashMapActiveJobs")}</p>
                              <p className="font-bold text-slate-800">{stat.activeJobs}</p>
                            </div>
                          </div>
                          <p className="text-xs text-slate-500 mb-2">{t("dashMapNetRegion")}: <strong className={stat.arTotal >= stat.apTotal ? "text-emerald-600" : "text-amber-600"}>{money(stat.arTotal - stat.apTotal)} {currency}</strong></p>
                          {stat.hasOverdueAr && <p className="text-xs text-red-600 font-medium mb-2">{t("dashMapOverdueAlert")}</p>}
                          <button
                            type="button"
                            className="text-xs font-medium text-blue-600 hover:text-blue-800 hover:underline"
                            onClick={(e) => { e.stopPropagation(); onRegionClick && onRegionClick(stat.regionId); }}
                          >
                            {t("dashMapGoRegion")} →
                          </button>
                        </div>
                      )}
                    </li>
                  );
                })}
              </ul>
            </div>
          </article>
        );
      }

      function LoginPage({ onLogin }) {
        const lang = localStorage.getItem("erp_lang") || "zh_TW";
        const t = (key) => (I18N[i18nLang(lang)] && I18N[i18nLang(lang)][key]) || I18N.en[key] || key;
        const [username, setUsername] = useState(() => localStorage.getItem("erp_saved_username") || "");
        const [password, setPassword] = useState("");
        const [remember, setRemember] = useState(() => localStorage.getItem("erp_remember_login") === "1");
        const [showPw, setShowPw] = useState(false);
        const [error, setError] = useState("");
        const [submitting, setSubmitting] = useState(false);
        const [cloudReady, setCloudReady] = useState(false);

        useEffect(() => {
          let cancelled = false;
          getUsersForLogin().finally(() => { if (!cancelled) setCloudReady(true); });
          return () => { cancelled = true; };
        }, []);

        async function handleSubmit(e) {
          e.preventDefault();
          setError("");
          const id = String(username || "").trim();
          const pw = String(password || "");
          if (!id || !pw) { setError(t("loginFail")); return; }
          setSubmitting(true);
          let users = getErpUsers();
          let found = findUserForLogin(users, id, pw);
          if (!found) {
            users = await getUsersForLogin();
            found = findUserForLogin(users, id, pw);
          }
          if (!found) { setError(t("loginFail")); setSubmitting(false); return; }
          if (!found.is_active) { setError(t("loginInactive")); setSubmitting(false); return; }
          if (remember) localStorage.setItem("erp_saved_username", id);
          else localStorage.removeItem("erp_saved_username");
          persistAuthSession(found.id, remember);
          onLogin(found.id);
        }

        return (
          <div className="erp-login-shell">
            <div className="erp-login-card">
              <div className="erp-login-brand">
                <img src={AIRLINK_BRAND.loginLogo} alt="AIRLINK ERP" />
              </div>
              <form onSubmit={handleSubmit} className="erp-login-form" aria-label="Login">
                <div className="erp-login-field">
                  <svg className="w-5 h-5 text-slate-400 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M16 7a4 4 0 11-8 0 4 4 0 018 0zM12 14a7 7 0 00-7 7h14a7 7 0 00-7-7z" /></svg>
                  <input
                    type="text"
                    autoComplete="username"
                    placeholder={t("loginUsername")}
                    value={username}
                    onChange={(e) => setUsername(e.target.value)}
                  />
                </div>
                <div className="erp-login-field">
                  <svg className="w-5 h-5 text-slate-400 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M12 15v2m-6 4h12a2 2 0 002-2v-6a2 2 0 00-2-2H6a2 2 0 00-2 2v6a2 2 0 002 2zm10-10V7a4 4 0 00-8 0v4h8z" /></svg>
                  <input
                    type={showPw ? "text" : "password"}
                    autoComplete="current-password"
                    placeholder={t("loginPassword")}
                    value={password}
                    onChange={(e) => setPassword(e.target.value)}
                  />
                  <button type="button" onClick={() => setShowPw((v) => !v)} className="text-slate-400 hover:text-slate-600 p-1 shrink-0" tabIndex={-1} aria-label="Toggle password">
                    {showPw ? (
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M13.875 18.825A10.05 10.05 0 0112 19c-4.478 0-8.268-2.943-9.543-7a9.97 9.97 0 011.563-3.029m5.858 3.05a3 3 0 114.243 4.243M9.878 9.878l4.242 4.242M9.88 9.88l-3.29-3.29m7.532 7.532l3.29 3.29M3 3l3.59 3.59m0 0A9.953 9.953 0 0112 5c4.478 0 8.268 2.943 9.543 7a10.025 10.025 0 01-4.132 5.411m0 0L21 21" /></svg>
                    ) : (
                      <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M15 12a3 3 0 11-6 0 3 3 0 016 0z" /><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="1.75" d="M2.458 12C3.732 7.943 7.523 5 12 5c4.478 0 8.268 2.943 9.542 7-1.274 4.057-5.064 7-9.542 7-4.477 0-8.268-2.943-9.542-7z" /></svg>
                    )}
                  </button>
                </div>
                <label className="erp-login-remember cursor-pointer select-none">
                  <input type="checkbox" className="rounded border-slate-300 text-blue-600" checked={remember} onChange={(e) => setRemember(e.target.checked)} />
                  {t("loginRemember")}
                </label>
                {error && <p className="erp-login-error">{error}</p>}
                <button type="submit" className="erp-login-btn" disabled={submitting || !cloudReady}>{submitting ? "…" : (!cloudReady ? "…" : t("loginBtn"))}</button>
              </form>
            </div>
          </div>
        );
      }

      function resolveAuthUserId() {
        const raw = sessionStorage.getItem("erp_session_user_id")
          || (localStorage.getItem("erp_remember_login") === "1" ? localStorage.getItem("erp_session_user_id") : null);
        if (!raw) return null;
        const users = getErpUsers();
        const u = users.find((x) => x.id === Number(raw) && x.is_active);
        return u ? u.id : null;
      }

      function persistAuthSession(userId, remember) {
        sessionStorage.setItem("erp_session_user_id", String(userId));
        localStorage.setItem("erp_current_user_id", String(userId));
        if (remember) {
          localStorage.setItem("erp_remember_login", "1");
          localStorage.setItem("erp_session_user_id", String(userId));
        } else {
          localStorage.removeItem("erp_remember_login");
          localStorage.removeItem("erp_session_user_id");
        }
      }

      function clearAuthSession() {
        sessionStorage.removeItem("erp_session_user_id");
        localStorage.removeItem("erp_session_user_id");
        localStorage.removeItem("erp_remember_login");
        localStorage.removeItem("erp_current_user_id");
      }

      function ErpRoot() {
        const [authUserId, setAuthUserId] = useState(() => resolveAuthUserId());
        if (!authUserId) {
          return <LoginPage onLogin={(id) => setAuthUserId(id)} />;
        }
        return (
          <App
            authUserId={authUserId}
            onLogout={() => {
              clearAuthSession();
              setAuthUserId(null);
            }}
          />
        );
      }

      function App({ authUserId, onLogout }) {
        const [page, setPage] = useState("dashboard");

        useEffect(() => {
          if (page === "si") setPage("clients");
        }, [page]);
        const [lang, setLang] = useState(() => localStorage.getItem("erp_lang") || "zh_TW");
        const [baseCurrency, setBaseCurrency] = useState("HKD");
        const [customCurrency, setCustomCurrency] = useState("");
        const [clients, setClients] = useState(() => initialBusinessList("erp_clients", ensureRegionOnRecords, initialClients));
        const [jobs, setJobs] = useState(() => initialBusinessList("erp_jobs", (rows) => ensureJobsQuotations(ensureRegionOnRecords(rows)), initialJobs));
        const [quotations, setQuotations] = useState(() => initialBusinessList("erp_quotations", ensureRegionOnRecords, initialQuotations));
        const [vendors, setVendors] = useState(() => initialBusinessList("erp_vendors", (rows) => ensureVendorNumbers(ensureRegionOnRecords(rows)), initialVendors));
        const [arInvoices, setArInvoices] = useState(() => initialBusinessList("erp_ar", ensureRegionOnRecords, initialAr));
        const [apBills, setApBills] = useState(() => initialBusinessList("erp_ap", ensureRegionOnRecords, initialAp));
        const [approvals] = useState(initialApprovals);

        const [newTxOpen, setNewTxOpen] = useState(false);
        const [detailPanel, setDetailPanel] = useState(null);
        const [highlightId, setHighlightId] = useState(null);

        const [jobViewMode, setJobViewMode] = useState("list");
        const [arViewMode, setArViewMode] = useState("list");
        const [apViewMode, setApViewMode] = useState("list");
        const [jobClientFilter, setJobClientFilter] = useState("all");
        const [jobStatusFilter, setJobStatusFilter] = useState("all");
        const [jobTypeFilter, setJobTypeFilter] = useState("all");
        const [jobSearch, setJobSearch] = useState("");
        const [clientsSearch, setClientsSearch] = useState("");
        const [quotationSearch, setQuotationSearch] = useState("");
        const [vendorsSearch, setVendorsSearch] = useState("");
        const [mgmtClientFilter, setMgmtClientFilter] = useState("all");
        const [mgmtGroupCurrency, setMgmtGroupCurrency] = useState("HKD");
        const [mgmtCompareCurrency, setMgmtCompareCurrency] = useState("USD");
        const [arClientFilter, setArClientFilter] = useState("all");
        const [arSearch, setArSearch] = useState("");
        const [apClientFilter, setApClientFilter] = useState("all");
        const [apJobFilter, setApJobFilter] = useState("all");
        const [apSearch, setApSearch] = useState("");
        const [sidebarCollapsed, setSidebarCollapsed] = useState(() => localStorage.getItem("erp_sidebar_collapsed") === "1");
        const [activeRegion, setActiveRegion] = useState(() => localStorage.getItem("erp_active_region") || "HK");

        const [clientModal, setClientModal] = useState(null);
        const [jobModal, setJobModal] = useState(null);
        const [quotationModal, setQuotationModal] = useState(null);
        const [vendorModal, setVendorModal] = useState(null);
        const [arModal, setArModal] = useState(null);
        const [apModal, setApModal] = useState(null);

        const [users, setUsers] = useState(() => getErpUsers());
        const [pwForm, setPwForm] = useState({ current: "", next: "", confirm: "" });
        const sessionUserId = authUserId;
        const [auditLogs, setAuditLogs] = useState(() => (isCloudOnlyPreferred() ? [] : loadJson("erp_audit_logs", initialAuditLogs)));
        const [worldTimeZone, setWorldTimeZone] = useState(() => localStorage.getItem("erp_world_timezone") || "Asia/Taipei");
        const [worldClock, setWorldClock] = useState(() => new Date());
        const [settingsTab, setSettingsTab] = useState("general");
        const [companyName, setCompanyName] = useState(() => localStorage.getItem("erp_company_name") || AIRLINK_BRAND.name);
        const sidebarLogo = AIRLINK_BRAND.logo;

        useEffect(() => {
          localStorage.setItem("erp_company_logo", sidebarLogo);
          localStorage.setItem("erp_company_name", AIRLINK_BRAND.name);
        }, [sidebarLogo]);
        const [erpSyncKey, setErpSyncKey] = useState(() => localStorage.getItem("erp_sync_key") || "");
        const [cloudOnlyMode, setCloudOnlyMode] = useState(() => isCloudOnlyPreferred());
        const [cloudDbStatus, setCloudDbStatus] = useState(null);
        const [liveSyncEnabled, setLiveSyncEnabled] = useState(false);
        const [liveSyncReady, setLiveSyncReady] = useState(false);
        const [liveSyncStatus, setLiveSyncStatus] = useState("off");
        const [liveSyncLastAt, setLiveSyncLastAt] = useState(null);
        const liveServerVersionRef = useRef(0);
        const liveSyncedSnapshotRef = useRef({});
        const liveApplyingRemoteRef = useRef(false);
        const livePushTimerRef = useRef(null);
        const livePollRef = useRef(null);
        const importModuleRef = useRef(null);
        const importFileInputRef = useRef(null);
        const [userModal, setUserModal] = useState(null);
        const [importPreview, setImportPreview] = useState(null);
        const [importLoading, setImportLoading] = useState(false);
        const [importStatus, setImportStatus] = useState("");
        const [tableSort, setTableSort] = useState({});
        const ERP_BUILD_ID = "airlink-2026-07-10d";
        const [ongoingEditId, setOngoingEditId] = useState(null);
        const [ongoingDraft, setOngoingDraft] = useState({ billedAmt: "", remarks: "" });
        const [auditFilters, setAuditFilters] = useState({ dateFrom: "", dateTo: "", userId: "all", module: "all", action: "all", q: "" });
        const [overviewFrom, setOverviewFrom] = useState(() => {
          const d = new Date();
          return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-01";
        });
        const [overviewTo, setOverviewTo] = useState(() => new Date().toISOString().slice(0, 10));
        const [monthlyReportYear, setMonthlyReportYear] = useState(() => String(new Date().getFullYear()));
        const [monthlyReportTab, setMonthlyReportTab] = useState("report");
        const [monthlyPoLines, setMonthlyPoLines] = useState(() => initialBusinessList("erp_monthly_po_lines", null, []));
        const [monthlyArLines, setMonthlyArLines] = useState(() => initialBusinessList("erp_monthly_ar_lines", null, []));
        const [monthlyArExpectedSnapshots, setMonthlyArExpectedSnapshots] = useState(() => initialBusinessList("erp_monthly_ar_expected", null, {}));

        const t = (key) => (I18N[i18nLang(lang)] && I18N[i18nLang(lang)][key]) || I18N.en[key] || key;

        function isImportStatusSticky(msg) {
          if (!msg) return false;
          return msg === t("importSyncFail") || msg.includes("雲端同步失敗") || msg.toLowerCase().includes("cloud sync failed");
        }

        useEffect(() => {
          if (!importStatus || isImportStatusSticky(importStatus)) return;
          const tmr = setTimeout(() => setImportStatus(""), 8000);
          return () => clearTimeout(tmr);
        }, [importStatus, lang]);

        function patchArModalData(data) {
          const next = applyArDueDate(data, clients, jobs);
          return { ...next, payment_status: deriveArPaymentStatus(next) };
        }

        useEffect(() => {
          const timer = setInterval(() => setWorldClock(new Date()), 1000);
          return () => clearInterval(timer);
        }, []);

        useEffect(() => { localStorage.setItem("erp_users", JSON.stringify(users)); }, [users]);
        useEffect(() => { if (!cloudOnlyMode) localStorage.setItem("erp_audit_logs", JSON.stringify(auditLogs)); }, [auditLogs, cloudOnlyMode]);
        useEffect(() => { if (!cloudOnlyMode) localStorage.setItem("erp_clients", JSON.stringify(clients)); }, [clients, cloudOnlyMode]);
        useEffect(() => { if (!cloudOnlyMode) localStorage.setItem("erp_jobs", JSON.stringify(jobs)); }, [jobs, cloudOnlyMode]);
        useEffect(() => { if (!cloudOnlyMode) localStorage.setItem("erp_quotations", JSON.stringify(quotations)); }, [quotations, cloudOnlyMode]);
        useEffect(() => { if (!cloudOnlyMode) localStorage.setItem("erp_vendors", JSON.stringify(vendors)); }, [vendors, cloudOnlyMode]);
        useEffect(() => { if (!cloudOnlyMode) localStorage.setItem("erp_ar", JSON.stringify(arInvoices)); }, [arInvoices, cloudOnlyMode]);
        useEffect(() => { if (!cloudOnlyMode) localStorage.setItem("erp_ap", JSON.stringify(apBills)); }, [apBills, cloudOnlyMode]);
        useEffect(() => {
          if (cloudOnlyMode) return;
          if (saveErpStorage("erp_monthly_po_lines", monthlyPoLines)) localStorage.setItem("erp_monthly_po_saved_at", new Date().toISOString());
        }, [monthlyPoLines, cloudOnlyMode]);
        useEffect(() => {
          if (cloudOnlyMode) return;
          if (saveErpStorage("erp_monthly_ar_lines", monthlyArLines)) localStorage.setItem("erp_monthly_ar_saved_at", new Date().toISOString());
        }, [monthlyArLines, cloudOnlyMode]);
        useEffect(() => { if (!cloudOnlyMode) saveErpStorage("erp_monthly_ar_expected", monthlyArExpectedSnapshots); }, [monthlyArExpectedSnapshots, cloudOnlyMode]);
        useEffect(() => { localStorage.setItem("erp_current_user_id", String(sessionUserId)); }, [sessionUserId]);
        useEffect(() => { localStorage.setItem("erp_world_timezone", worldTimeZone); }, [worldTimeZone]);
        useEffect(() => { localStorage.setItem("erp_sidebar_collapsed", sidebarCollapsed ? "1" : "0"); }, [sidebarCollapsed]);
        useEffect(() => { localStorage.setItem("erp_active_region", activeRegion); }, [activeRegion]);
        useEffect(() => { localStorage.setItem("erp_sync_key", erpSyncKey); }, [erpSyncKey]);
        useEffect(() => {
          console.info("AIRLINK ERP build:", ERP_BUILD_ID);
        }, []);

        useEffect(() => {
          fetch("/api/health")
            .then((r) => (r.ok ? r.json() : Promise.reject(new Error("health"))))
            .then(async (j) => {
              setCloudDbStatus(j);
              if (!j.cloud_sync_configured) return;
              try {
                const cfgRes = await fetch("/api/sync/client-config");
                const cfg = cfgRes.ok ? await cfgRes.json() : {};
                if (cfg.sync_key) setErpSyncKey(cfg.sync_key);
                if (cfg.auto_sync) {
                  setLiveSyncEnabled(true);
                  setCloudOnlyMode(true);
                  localStorage.setItem("erp_cloud_only", "1");
                  clearLocalBusinessCache();
                }
              } catch (_) { /* ignore */ }
            })
            .catch(() => setCloudDbStatus(null));
        }, []);

        useEffect(() => {
          if (!liveSyncEnabled || !liveSyncReady) return;
          if (livePollRef.current) clearInterval(livePollRef.current);
          livePollRef.current = setInterval(() => { pollLiveSync(); }, 3000);
          return () => { if (livePollRef.current) clearInterval(livePollRef.current); };
        }, [liveSyncEnabled, liveSyncReady]);

        function scheduleLiveSyncPush(immediate) {
          if (!liveSyncEnabled || !liveSyncReady || !canLiveSyncPush() || liveApplyingRemoteRef.current) return;
          if (!erpSyncKey.trim()) return;
          if (livePushTimerRef.current) clearTimeout(livePushTimerRef.current);
          if (immediate) {
            flushLiveSyncPush();
            return;
          }
          livePushTimerRef.current = setTimeout(() => { flushLiveSyncPush(); }, 400);
        }

        useEffect(() => {
          if (!liveSyncEnabled || !liveSyncReady || !canLiveSyncPush() || liveApplyingRemoteRef.current) return;
          if (!erpSyncKey.trim()) return;
          scheduleLiveSyncPush(false);
          return () => { if (livePushTimerRef.current) clearTimeout(livePushTimerRef.current); };
        }, [clients, jobs, quotations, vendors, arInvoices, apBills, users, auditLogs, monthlyPoLines, monthlyArLines, monthlyArExpectedSnapshots, lang, activeRegion, worldTimeZone, companyName, sessionUserId, sidebarCollapsed, liveSyncEnabled, liveSyncReady, erpSyncKey, sessionUserId]);

        useEffect(() => {
          if (liveSyncEnabled && !liveSyncReady) {
            bootstrapLiveSync().then((ok) => {
              if (!ok) setLiveSyncEnabled(false);
            });
          }
        }, [liveSyncEnabled]);

        useEffect(() => {
          if (!liveSyncEnabled) return;
          setLiveSyncReady(false);
          liveSyncedSnapshotRef.current = {};
          bootstrapLiveSync();
        }, [sessionUserId]);

        function getCurrentUser() {
          return users.find((u) => u.id === sessionUserId && u.is_active) || null;
        }

        function buildSharedSettings() {
          return { lang, activeRegion, worldTimeZone, sidebarCollapsed };
        }

        function applySharedSettings(s) {
          if (!s) return;
          if (s.lang) { localStorage.setItem("erp_lang", s.lang); setLang(s.lang); }
          if (s.activeRegion) { localStorage.setItem("erp_active_region", s.activeRegion); setActiveRegion(s.activeRegion); }
          if (s.worldTimeZone) setWorldTimeZone(s.worldTimeZone);
          if (s.sidebarCollapsed != null) setSidebarCollapsed(!!s.sidebarCollapsed);
        }

        function getUserRegions() {
          const u = getCurrentUser();
          if (!u) return ["HK"];
          if (u.role === "root" || (u.allowed_regions && u.allowed_regions.includes("ALL"))) return ERP_REGIONS.map((r) => r.id);
          return u.allowed_regions && u.allowed_regions.length ? u.allowed_regions : ["HK"];
        }

        function recordInActiveRegion(rec) {
          const rg = rec.region || "HK";
          const allowed = getUserRegions();
          if (activeRegion === "ALL") {
            if (allowed.includes("ALL")) return true;
            return allowed.includes(rg);
          }
          return rg === activeRegion && (allowed.includes("ALL") || allowed.includes(rg));
        }

        function regionForNewRecord() {
          if (activeRegion !== "ALL") return activeRegion;
          const allowed = getUserRegions().filter((x) => x !== "ALL");
          return allowed[0] || "HK";
        }

        function matchesSearch(row, q, keys) {
          if (!q.trim()) return true;
          const s = q.toLowerCase();
          return keys.some((k) => String(row[k] || "").toLowerCase().includes(s));
        }

        function applyTableSort(list, module) {
          const def = LIST_DEFAULT_SORT[module];
          if (!def) return list;
          const cur = tableSort[module] || def;
          return sortRecords(list, cur.field, cur.dir);
        }

        function toggleListSort(module, field) {
          setTableSort((prev) => {
            const def = LIST_DEFAULT_SORT[module] || { field, dir: "asc" };
            const cur = prev[module] || def;
            if (cur.field === field) {
              return { ...prev, [module]: { field, dir: cur.dir === "asc" ? "desc" : "asc" } };
            }
            return { ...prev, [module]: { field, dir: "asc" } };
          });
        }

        useEffect(() => {
          const allowed = getUserRegions();
          if (!allowed.includes(activeRegion)) {
            setActiveRegion(allowed.includes("ALL") ? "ALL" : allowed[0]);
          }
        }, [sessionUserId, users]);

        function isRoot() {
          const u = getCurrentUser();
          return !!(u && u.role === "root");
        }

        function can(module, action) {
          const u = getCurrentUser();
          if (!u || !u.is_active) return false;
          if (u.role === "root") return true;
          const p = u.permissions && u.permissions[module];
          if (!p) return false;
          return !!p[action];
        }

        function canLiveSyncPush() {
          if (isRoot()) return true;
          const mods = ["clients", "quotation", "job", "ar", "ap", "vendors", "monthly_report", "management_report", "dashboard", "overview", "ongoing", "settings"];
          return mods.some((m) => can(m, "add") || can(m, "edit") || can(m, "delete") || can(m, "import"));
        }

        function actionLabel(action) {
          const map = {
            create: t("actionCreate"), update: t("actionUpdate"), delete: t("actionDelete"),
            mark_paid: t("actionMarkPaid"), user_create: t("actionUserCreate"), user_update: t("actionUserUpdate"),
            user_delete: t("actionUserDelete"),
            export: t("actionExport"), import: t("actionImport")
          };
          return map[action] || action;
        }

        function logAudit(module, action, entity, detail) {
          const u = getCurrentUser();
          const entry = {
            id: nextErpRecordId(auditLogs),
            at: new Date().toISOString(),
            user_id: u ? u.id : null,
            user_name: u ? u.name : "-",
            user_email: u ? u.email : "-",
            module,
            action,
            entity: entity || "-",
            detail: detail || ""
          };
          setAuditLogs((prev) => [entry, ...prev].slice(0, 500));
        }

        function guardPermission(module, action) {
          if (!can(module, action)) {
            alert(t("noPermission"));
            return false;
          }
          return true;
        }

        const filteredAuditLogs = useMemo(() => {
          return auditLogs.filter((row) => {
            if (auditFilters.userId !== "all" && String(row.user_id) !== auditFilters.userId) return false;
            if (auditFilters.module !== "all" && row.module !== auditFilters.module) return false;
            if (auditFilters.action !== "all" && row.action !== auditFilters.action) return false;
            if (auditFilters.dateFrom && row.at.slice(0, 10) < auditFilters.dateFrom) return false;
            if (auditFilters.dateTo && row.at.slice(0, 10) > auditFilters.dateTo) return false;
            if (auditFilters.q) {
              const q = auditFilters.q.toLowerCase();
              const blob = [row.user_name, row.user_email, row.entity, row.detail, row.module, row.action].join(" ").toLowerCase();
              if (!blob.includes(q)) return false;
            }
            return true;
          });
        }, [auditLogs, auditFilters]);

        function changeLang(next) {
          setLang(next);
          localStorage.setItem("erp_lang", next);
        }

        useEffect(() => {
          const L = i18nLang(lang);
          document.documentElement.lang = L === "zh_TW" ? "zh-TW" : "en";
          document.title = (I18N[L] && I18N[L].appTitle) || I18N.en.appTitle;
        }, [lang]);

        function navigateTo(targetPage, opts = {}) {
          setDetailPanel(null);
          setPage(targetPage);
          if (opts.client) {
            setJobClientFilter(opts.client);
            const byCompany = clients.find((c) => c.company === opts.client);
            const byTitle = clients.find((c) => c.invoice_title === opts.client);
            const arClientVal = (byCompany?.invoice_title || byTitle?.invoice_title || opts.client);
            setArClientFilter(arClientVal);
            setApClientFilter(opts.client);
          }
          if (opts.jobId) setApJobFilter(String(opts.jobId));
          if (opts.highlightId) {
            setHighlightId(opts.highlightId);
            setTimeout(() => setHighlightId(null), 4000);
          }
        }

        function navigateToArByPo(customer, po) {
          const client = clients.find((c) => c.company === customer);
          const invoiceTitle = client?.invoice_title || customer;
          const match = arInvoices.find((r) => r.customer === customer && r.customer_po === po) ||
            arInvoices.find((r) => r.customer === invoiceTitle && r.customer_po === po);
          navigateTo("ar", { highlightId: match ? match.id : null });
        }

        function navigateToJobByPo(company, po) {
          const match = jobs.find((j) => j.company === company && j.customer_po === po);
          navigateTo("job", { highlightId: match ? match.id : null });
        }

        function markArPaid(id) {
          if (!guardPermission("ar", "edit")) return;
          const row = arInvoices.find((r) => r.id === id);
          const today = new Date().toISOString().slice(0, 10);
          setArInvoices((prev) => prev.map((r) => (r.id === id ? { ...r, payment_status: "Paid", payment_received_date: r.payment_received_date || today } : r)));
          logAudit("ar", "mark_paid", row ? row.invoice_no : String(id), `Marked AR ${row ? row.invoice_no : id} as paid`);
        }

        function markApPaid(id) {
          if (!guardPermission("ap", "edit")) return;
          const row = apBills.find((b) => b.id === id);
          const today = new Date().toISOString().slice(0, 10);
          setApBills((prev) => prev.map((b) => (b.id === id ? { ...b, pay_date: b.pay_date || today, payment_status: "Paid" } : b)));
          logAudit("ap", "mark_paid", row ? row.invoice_no : String(id), `Marked AP ${row ? row.invoice_no : id} as paid`);
        }

        function openArFromJob(j) {
          if (!guardPermission("ar", "add")) return;
          setDetailPanel(null);
          const client = j.client_id ? clients.find((c) => c.id === Number(j.client_id)) : clients.find((c) => c.company === j.company);
          const customerTitle = client?.invoice_title || j.company;
          setArModal({
            mode: "add",
            data: {
              job_id: String(j.id), job_no: j.job_no, customer: customerTitle, customer_po: j.customer_po || "",
              invoice_no: "", invoice_currency: "USD", invoice_amt: "", exchange_rate: "", base_amount: "",
              override_reason: "", invoice_date: "", due_date: "", payment_received_date: "", payment_status: "Awaiting Payment"
            }
          });
        }

        function rowHighlightClass(id) {
          return highlightId === id ? "ring-2 ring-inset ring-blue-500 bg-blue-50" : "";
        }

        const activeBaseCurrency = baseCurrency === "OTHER" ? (customCurrency || "USD").toUpperCase() : baseCurrency;
        const regionListCurrency = currencyForRegion(activeRegion, activeBaseCurrency);
        const amtInLabel = amtInCurrencyLabel(regionListCurrency, t);

        const scopedClients = useMemo(() => clients.filter(recordInActiveRegion), [clients, activeRegion, users, sessionUserId]);
        const scopedJobs = useMemo(() => jobs.filter(recordInActiveRegion), [jobs, activeRegion, users, sessionUserId]);
        const scopedQuotations = useMemo(() => quotations.filter(recordInActiveRegion), [quotations, activeRegion, users, sessionUserId]);
        const scopedVendors = useMemo(() => vendors.filter(recordInActiveRegion), [vendors, activeRegion, users, sessionUserId]);
        const scopedArInvoices = useMemo(() => arInvoices.filter(recordInActiveRegion), [arInvoices, activeRegion, users, sessionUserId]);
        const scopedApBills = useMemo(() => apBills.filter(recordInActiveRegion), [apBills, activeRegion, users, sessionUserId]);
        const sortedScopedClients = useMemo(() => sortRecords(scopedClients, "customer_no", "asc"), [scopedClients]);
        const sortedScopedJobs = useMemo(() => sortRecords(scopedJobs, "job_no", "asc"), [scopedJobs]);
        const sortedScopedVendors = useMemo(() => sortRecords(scopedVendors, "vendor_no", "asc"), [scopedVendors]);
        const newRecordRegion = regionForNewRecord();
        const jobFormRegion = jobModal?.mode === "edit"
          ? (jobs.find((j) => j.id === jobModal.id)?.region || newRecordRegion)
          : (jobModal?.data?.region || newRecordRegion);
        const quotationFormRegion = quotationModal?.mode === "edit" ? (quotations.find((q) => q.id === quotationModal.id)?.region || newRecordRegion) : newRecordRegion;
        const clientOptionsForJob = useMemo(() => sortedScopedClients
          .filter((c) => c.region === jobFormRegion)
          .map((c) => ({
            value: String(c.id),
            label: `${c.customer_no} · ${c.company}${c.bu_no ? " · BU " + c.bu_no : ""}`,
            searchText: [c.customer_no, c.company, c.primary_contact, c.bu_no].filter(Boolean).join(" ")
          })), [sortedScopedClients, jobFormRegion]);
        const clientOptionsForQuotation = useMemo(() => sortedScopedClients
          .filter((c) => c.region === quotationFormRegion)
          .map((c) => ({
            value: String(c.id),
            label: `${c.customer_no} · ${c.company}${c.bu_no ? " · BU " + c.bu_no : ""}`,
            searchText: [c.customer_no, c.company, c.primary_contact, c.bu_no].filter(Boolean).join(" ")
          })), [sortedScopedClients, quotationFormRegion]);
        const jobOptions = useMemo(() => sortedScopedJobs.map((j) => ({
          value: String(j.id),
          label: `${j.job_no} · ${j.company}`,
          searchText: [j.job_no, j.company, j.job_type, j.customer_po].filter(Boolean).join(" ")
        })), [sortedScopedJobs]);
        const quotationJobOptions = useMemo(() => sortedScopedJobs.map((j) => ({
          value: j.job_no || "",
          label: `${j.job_no} · ${j.company}`,
          searchText: [j.job_no, j.company, j.job_type].filter(Boolean).join(" ")
        })), [sortedScopedJobs]);
        const vendorOptions = useMemo(() => sortedScopedVendors.map((v) => ({
          value: v.name || "",
          label: `${v.vendor_no ? v.vendor_no + " · " : ""}${v.name}`,
          searchText: [v.vendor_no, v.name, v.contact, v.email, v.finance_contact, v.finance_email].filter(Boolean).join(" ")
        })), [sortedScopedVendors]);
        const accessibleQuotations = useMemo(() => {
          const allowed = getUserRegions();
          return quotations.filter((q) => quotationAccessibleInCurrentScope(q, allowed));
        }, [quotations, users, sessionUserId]);
        const quotationsForJobRegion = useMemo(() => {
          const allowed = getUserRegions();
          return quotations.filter((q) => {
            const rg = q.region || "HK";
            if (!(allowed.includes("ALL") || allowed.includes(rg))) return false;
            return rg === jobFormRegion;
          });
        }, [quotations, users, sessionUserId, jobFormRegion]);

        const mgmtDisplayCurrency = activeRegion === "ALL" ? mgmtGroupCurrency : (REGION_CURRENCY[activeRegion] || regionListCurrency);
        const mgmtRegionCount = useMemo(() => {
          if (activeRegion !== "ALL") return 1;
          const allowed = getUserRegions().filter((r) => r !== "ALL");
          return allowed.length || DATA_REGION_IDS.length;
        }, [activeRegion, users, sessionUserId]);
        const mgmtReport = useMemo(() => buildManagementReport({
          jobs: scopedJobs,
          quotations: scopedQuotations,
          arInvoices: scopedArInvoices,
          apBills: scopedApBills,
          clientFilter: mgmtClientFilter,
          displayCurrency: mgmtDisplayCurrency,
          compareCurrency: mgmtCompareCurrency,
          todayIso: new Date().toISOString().slice(0, 10),
          lang,
          activeRegion,
          regionCount: mgmtRegionCount
        }), [scopedJobs, scopedQuotations, scopedArInvoices, scopedApBills, mgmtClientFilter, mgmtDisplayCurrency, mgmtCompareCurrency, lang, activeRegion, mgmtRegionCount]);
        const mgmtClientOptions = useMemo(() => {
          const names = new Set();
          mgmtReport.clientMargins.forEach((c) => names.add(c.company));
          scopedClients.forEach((c) => { if (c.company) names.add(c.company); });
          return [...names].filter(Boolean).sort();
        }, [mgmtReport.clientMargins, scopedClients]);
        const mgmtClientSearchOptions = useMemo(() => {
          const byCompany = new Map(sortedScopedClients.map((c) => [c.company, c]));
          return [
            { value: "all", label: t("mgmtAllClients"), searchText: t("mgmtAllClients") },
            ...mgmtClientOptions.map((name) => {
              const client = byCompany.get(name);
              return {
                value: name,
                label: client ? `${client.customer_no} · ${client.company}${client.bu_no ? " · BU " + client.bu_no : ""}` : name,
                searchText: client ? [client.customer_no, client.company, client.primary_contact, client.bu_no].filter(Boolean).join(" ") : name
              };
            }).sort((a, b) => a.label.localeCompare(b.label, undefined, { numeric: true, sensitivity: "base" }))
          ];
        }, [mgmtClientOptions, sortedScopedClients, t]);

        const arDisplay = useMemo(() => scopedArInvoices.map((r) => {
          const due_date = r.due_date || computeArDueDate(r, clients, jobs);
          const payment_status = deriveArPaymentStatus({ ...r, due_date });
          return { ...r, due_date, payment_status, display_base_amount: resolveArDisplayAmount(r, regionListCurrency) };
        }), [scopedArInvoices, regionListCurrency, clients, jobs]);
        const apDisplay = useMemo(() => scopedApBills.map((b) => ({
          ...b,
          payment_status: deriveApPaymentStatus(b),
          display_base_amount: resolveApDisplayAmount(b, regionListCurrency)
        })), [scopedApBills, regionListCurrency]);
        const arTotal = useMemo(() => arDisplay.filter(isArUnpaid).reduce((a, b) => a + b.display_base_amount, 0), [arDisplay]);
        const apTotal = useMemo(() => apDisplay.filter(isApUnpaid).reduce((a, b) => a + b.display_base_amount, 0), [apDisplay]);
        const scopedArCount = scopedArInvoices.length;
        const scopedApCount = scopedApBills.length;
        const dashShowsAllPaid = scopedArCount + scopedApCount > 0 && arTotal === 0 && apTotal === 0;
        const recordsOutsideRegion = useMemo(() => {
          if (activeRegion === "ALL") return 0;
          const arOther = arInvoices.filter((r) => (r.region || "HK") !== activeRegion).length;
          const apOther = apBills.filter((b) => (b.region || "HK") !== activeRegion).length;
          return arOther + apOther;
        }, [arInvoices, apBills, activeRegion]);
        const dashRegionMismatch = activeRegion !== "ALL" && scopedArCount + scopedApCount === 0 && recordsOutsideRegion > 0;
        const netPosition = arTotal - apTotal;
        const activeJobCount = useMemo(() => scopedJobs.filter((j) => !isJobCompleted(j.status)).length, [scopedJobs]);

        const arDueSoonList = useMemo(() =>
          arDisplay
            .filter((r) => isArUnpaid(r) && r.due_date)
            .map((r) => ({ ...r, days_until: daysUntilDue(r.due_date), overdue_days: calcOverdueDays(r.due_date) }))
            .filter((r) => r.days_until !== null && r.days_until >= 0 && r.days_until <= 14)
            .sort((a, b) => a.days_until - b.days_until)
            .slice(0, 8),
          [arDisplay]
        );

        const apDueSoonList = useMemo(() =>
          apDisplay
            .filter((b) => !isApPaid(b) && b.due_date)
            .map((b) => ({ ...b, days_until: daysUntilDue(b.due_date), overdue_days: calcOverdueDays(b.due_date) }))
            .filter((b) => b.days_until !== null && b.days_until >= 0 && b.days_until <= 14)
            .sort((a, b) => a.days_until - b.days_until)
            .slice(0, 8),
          [apDisplay]
        );

        const arOverdueList = useMemo(() =>
          arDisplay
            .map((r) => ({ ...r, overdue_days: calcOverdueDays(r.due_date), days_until: daysUntilDue(r.due_date) }))
            .filter((r) => r.overdue_days > 0 && isArUnpaid(r))
            .sort((a, b) => b.overdue_days - a.overdue_days)
            .slice(0, 6),
          [arDisplay]
        );

        const apOverdueList = useMemo(() =>
          apDisplay
            .map((b) => ({ ...b, overdue_days: calcOverdueDays(b.due_date), days_until: daysUntilDue(b.due_date) }))
            .filter((b) => b.overdue_days > 0 && !isApPaid(b))
            .sort((a, b) => b.overdue_days - a.overdue_days)
            .slice(0, 6),
          [apDisplay]
        );

        const aging = useMemo(() => {
          const bucket = { current: 0, d1_30: 0, d31_60: 0, d61_90: 0, d90_plus: 0 };
          arDisplay.filter(isArUnpaid).forEach((r) => {
            const od = calcOverdueDays(r.due_date);
            if (od <= 0) bucket.current += r.display_base_amount;
            else if (od <= 30) bucket.d1_30 += r.display_base_amount;
            else if (od <= 60) bucket.d31_60 += r.display_base_amount;
            else if (od <= 90) bucket.d61_90 += r.display_base_amount;
            else bucket.d90_plus += r.display_base_amount;
          });
          return bucket;
        }, [arDisplay]);

        const agingMax = useMemo(() => Math.max(1, aging.current + aging.d1_30 + aging.d31_60 + aging.d61_90 + aging.d90_plus), [aging]);

        const cashFlowSeries = useMemo(() => {
          const year = new Date().getFullYear();
          const months = [];
          for (let m = 0; m < 12; m++) {
            const key = year + "-" + String(m + 1).padStart(2, "0");
            const label = monthChartLabel(year, m, lang);
            months.push({ key, label, inBase: 0, outBase: 0 });
          }
          const monthMap = Object.fromEntries(months.map((x) => [x.key, x]));
          arDisplay.filter((r) => isArUnpaid(r) && r.due_date).forEach((r) => {
            const key = dueDateMonthKey(r.due_date, year);
            if (key && monthMap[key]) monthMap[key].inBase += r.display_base_amount;
          });
          apDisplay.filter((b) => !isApPaid(b) && b.due_date).forEach((b) => {
            const key = dueDateMonthKey(b.due_date, year);
            if (key && monthMap[key]) monthMap[key].outBase += b.display_base_amount;
          });
          return months;
        }, [arDisplay, apDisplay, lang]);

        const cashFlowFirstHalf = useMemo(() => cashFlowSeries.slice(0, 6), [cashFlowSeries]);
        const cashFlowSecondHalf = useMemo(() => cashFlowSeries.slice(6, 12), [cashFlowSeries]);

        const maxBar = useMemo(() => Math.max(1, ...cashFlowSeries.flatMap((p) => [p.inBase, p.outBase])), [cashFlowSeries]);

        const cashFlowTotals = useMemo(() => ({
          in: cashFlowSeries.reduce((s, p) => s + p.inBase, 0),
          out: cashFlowSeries.reduce((s, p) => s + p.outBase, 0)
        }), [cashFlowSeries]);

        const overviewReport = useMemo(() => {
          const arRows = arDisplay.filter((r) => dateInRange(r.invoice_date, overviewFrom, overviewTo));
          const apRows = apDisplay.filter((b) => dateInRange(b.invoice_date, overviewFrom, overviewTo));
          const inTotal = arRows.reduce((s, r) => s + r.display_base_amount, 0);
          const outTotal = apRows.reduce((s, b) => s + b.display_base_amount, 0);
          return { arRows, apRows, inTotal, outTotal };
        }, [arDisplay, apDisplay, overviewFrom, overviewTo]);

        const overviewCollectedSeries = useMemo(() => {
          const year = new Date().getFullYear();
          const months = [];
          for (let m = 0; m < 12; m++) {
            const key = year + "-" + String(m + 1).padStart(2, "0");
            const label = monthChartLabel(year, m, lang);
            months.push({ key, label, inBase: 0, outBase: 0 });
          }
          const monthMap = Object.fromEntries(months.map((x) => [x.key, x]));
          arDisplay.filter(isArPaid).forEach((r) => {
            const payDate = arPaymentDate(r);
            if (!payDate) return;
            const key = dueDateMonthKey(payDate, year);
            if (key && monthMap[key]) monthMap[key].inBase += r.display_base_amount;
          });
          apDisplay.filter((b) => isApPaid(b) && b.pay_date).forEach((b) => {
            const key = dueDateMonthKey(b.pay_date, year);
            if (key && monthMap[key]) monthMap[key].outBase += b.display_base_amount;
          });
          return months;
        }, [arDisplay, apDisplay, lang]);

        const overviewCollectedFirstHalf = useMemo(() => overviewCollectedSeries.slice(0, 6), [overviewCollectedSeries]);
        const overviewCollectedSecondHalf = useMemo(() => overviewCollectedSeries.slice(6, 12), [overviewCollectedSeries]);
        const overviewCollectedMaxBar = useMemo(() => Math.max(1, ...overviewCollectedSeries.flatMap((p) => [p.inBase, p.outBase])), [overviewCollectedSeries]);
        const overviewCollectedTotals = useMemo(() => ({
          in: overviewCollectedSeries.reduce((s, p) => s + p.inBase, 0),
          out: overviewCollectedSeries.reduce((s, p) => s + p.outBase, 0)
        }), [overviewCollectedSeries]);

        const thisMonthKey = useMemo(() => {
          const d = new Date();
          return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0");
        }, []);

        const thisMonthIn = useMemo(() => cashFlowSeries.find((m) => m.key === thisMonthKey)?.inBase || 0, [cashFlowSeries, thisMonthKey]);
        const thisMonthOut = useMemo(() => cashFlowSeries.find((m) => m.key === thisMonthKey)?.outBase || 0, [cashFlowSeries, thisMonthKey]);

        const overdueArTotal = useMemo(() =>
          arDisplay.filter((r) => calcOverdueDays(r.due_date) > 0 && isArUnpaid(r)).reduce((s, r) => s + r.display_base_amount, 0),
          [arDisplay]
        );
        const overdueApTotal = useMemo(() =>
          apDisplay.filter((b) => calcOverdueDays(b.due_date) > 0 && !isApPaid(b)).reduce((s, b) => s + b.display_base_amount, 0),
          [apDisplay]
        );
        const overdueArCount = useMemo(() => arDisplay.filter((r) => calcOverdueDays(r.due_date) > 0 && isArUnpaid(r)).length, [arDisplay]);
        const overdueApCount = useMemo(() => apDisplay.filter((b) => calcOverdueDays(b.due_date) > 0 && !isApPaid(b)).length, [apDisplay]);

        const regionalMapStats = useMemo(() => {
          const allowed = getUserRegions();
          const ids = DATA_REGION_IDS.filter((id) => allowed.includes("ALL") || allowed.includes(id));
          return ids.map((regionId) => {
            const rAr = arInvoices.filter((r) => (r.region || "HK") === regionId);
            const rJobs = jobs.filter((j) => (j.region || "HK") === regionId);
            const rAp = apBills.filter((b) => (b.region || "HK") === regionId);
            const arUnpaid = rAr.filter(isArUnpaid);
            const arTotal = arUnpaid.reduce((s, r) => s + resolveArDisplayAmount(r, regionListCurrency), 0);
            const apTotal = rAp.filter((b) => !isApPaid(b)).reduce((s, b) => s + resolveApDisplayAmount(b, regionListCurrency), 0);
            const hasOverdueAr = arUnpaid.some((r) => calcOverdueDays(r.due_date) > 0);
            const activeJobs = rJobs.filter((j) => !isJobCompleted(j.status)).length;
            return { regionId, arTotal, apTotal, activeJobs, hasOverdueAr };
          });
        }, [arInvoices, jobs, apBills, regionListCurrency, users, sessionUserId]);

        const filteredJobs = useMemo(() => {
          let list = scopedJobs;
          if (jobClientFilter !== "all") list = list.filter((j) => j.company === jobClientFilter);
          if (jobStatusFilter !== "all") list = list.filter((j) => j.status === jobStatusFilter);
          if (jobTypeFilter !== "all") list = list.filter((j) => (j.job_type || "") === jobTypeFilter);
          if (jobSearch.trim()) list = list.filter((j) => matchesSearch(j, jobSearch, ["job_no", "company", "description", "customer_po", "quotation_no", "job_type", "status", "remarks", "currency"]));
          return applyTableSort(list, "job");
        }, [scopedJobs, jobClientFilter, jobStatusFilter, jobTypeFilter, jobSearch, tableSort]);

        const filteredClients = useMemo(() => {
          let list = scopedClients;
          if (clientsSearch.trim()) list = list.filter((c) => matchesSearch(c, clientsSearch, ["customer_no", "company", "invoice_title", "gst_no", "primary_contact", "company_phone", "mobile_phone", "email", "address", "postal_code", "account_dept_contact", "payment_terms", "bu_no"]));
          return applyTableSort(list, "clients");
        }, [scopedClients, clientsSearch, tableSort]);

        const filteredQuotations = useMemo(() => {
          let list = scopedQuotations;
          if (quotationSearch.trim()) list = list.filter((q) => matchesSearch(q, quotationSearch, ["quotation_no", "company", "description", "job_no", "status", "currency", "quotation_date", "valid_until"]));
          return applyTableSort(list, "quotation");
        }, [scopedQuotations, quotationSearch, tableSort]);

        const filteredVendors = useMemo(() => {
          let list = scopedVendors;
          if (vendorsSearch.trim()) list = list.filter((v) => matchesSearch(v, vendorsSearch, ["vendor_no", "name", "email", "bank", "swift_code", "charge", "contact", "phone", "address", "finance_contact", "finance_email"]));
          return applyTableSort(list, "vendors");
        }, [scopedVendors, vendorsSearch, tableSort]);

        const ongoingJobsList = useMemo(() => scopedJobs
          .filter((j) => !isJobCompleted(j.status))
          .map((j) => {
            const poAmt = Number(j.customer_po_amount || 0);
            const jobCurrency = j.currency || "USD";
            const billedAmt = Number(j.ongoing_billed_amt || 0);
            const poBalance = poAmt - billedAmt;
            const balanceUsd = convertCurrency(poBalance, jobCurrency, "USD");
            const poUsd = convertCurrency(poAmt, jobCurrency, "USD");
            return {
              ...j,
              nature: jobNatureCategory(j.job_type),
              arCurrency: jobCurrency,
              poAmt,
              poAmtLabel: "",
              billedAmt,
              poBalance,
              balanceUsd,
              poUsd
            };
          })
          .sort((a, b) => (a.job_no || "").localeCompare(b.job_no || "")),
        [scopedJobs]);

        const ongoingTotals = useMemo(() => {
          let totalPoUsd = 0;
          const byNature = { Maintenance: 0, Service: 0, Project: 0, Trade: 0 };
          ongoingJobsList.forEach((j) => {
            totalPoUsd += j.poUsd;
            if (byNature[j.nature] != null) byNature[j.nature] += j.poUsd;
            else byNature.Project += j.poUsd;
          });
          return { totalPoUsd, Maintenance: byNature.Maintenance, Service: byNature.Service, Project: byNature.Project, Trade: byNature.Trade };
        }, [ongoingJobsList]);

        function filterMonthlySheetLines(lines, year) {
          const y = Number(year);
          const allowed = getUserRegions();
          return lines.filter((l) => {
            if (l.year !== y) return false;
            const rg = l.region || "HK";
            if (activeRegion === "ALL") return allowed.includes("ALL") || allowed.includes(rg);
            return rg === activeRegion;
          });
        }

        const scopedMonthlyPoLines = useMemo(() => filterMonthlySheetLines(monthlyPoLines, monthlyReportYear), [monthlyPoLines, monthlyReportYear, activeRegion, sessionUserId, users]);
        const scopedMonthlyArLines = useMemo(() => filterMonthlySheetLines(monthlyArLines, monthlyReportYear), [monthlyArLines, monthlyReportYear, activeRegion, sessionUserId, users]);

        const poMonthTotals = useMemo(() => {
          const totals = Array(12).fill(0);
          scopedMonthlyPoLines.forEach((l) => {
            if (l.month >= 1 && l.month <= 12) totals[l.month - 1] += Number(l.base_amount || 0);
          });
          return totals;
        }, [scopedMonthlyPoLines]);

        const arInvoiceMonthTotals = useMemo(() => {
          const totals = Array(12).fill(0);
          scopedMonthlyArLines.forEach((l) => {
            if (l.month >= 1 && l.month <= 12) totals[l.month - 1] += Number(l.base_amount || 0);
          });
          return totals;
        }, [scopedMonthlyArLines]);

        const monthlyPoGrid = useMemo(() => buildMonthDetailGrid(scopedMonthlyPoLines, monthlyReportYear, lang), [scopedMonthlyPoLines, monthlyReportYear, lang]);
        const monthlyArGrid = useMemo(() => buildMonthDetailGrid(scopedMonthlyArLines, monthlyReportYear, lang), [scopedMonthlyArLines, monthlyReportYear, lang]);

        const monthlyReportSeries = useMemo(() => {
          const today = new Date().toISOString().slice(0, 10);
          return buildMonthlyReportSeries(Number(monthlyReportYear), regionListCurrency, lang, poMonthTotals, arInvoiceMonthTotals, scopedArInvoices, monthlyArExpectedSnapshots, activeRegion, today);
        }, [monthlyReportYear, regionListCurrency, lang, poMonthTotals, arInvoiceMonthTotals, scopedArInvoices, monthlyArExpectedSnapshots, activeRegion]);

        useEffect(() => {
          const today = new Date().toISOString().slice(0, 10);
          const y = Number(monthlyReportYear);
          if (!scopedArInvoices.length) {
            setMonthlyArExpectedSnapshots((prev) => {
              const next = { ...prev };
              let changed = false;
              for (let m = 0; m < 12; m++) {
                const key = arExpectedSnapshotKey(y, m, activeRegion);
                if (next[key] != null) {
                  delete next[key];
                  changed = true;
                }
              }
              return changed ? next : prev;
            });
            return;
          }
          const pending = {};
          for (let m = 0; m < 12; m++) {
            const key = arExpectedSnapshotKey(y, m, activeRegion);
            const ended = reportMonthHasEnded(y, m, today);
            const asOf = ended ? monthEndIso(y, m) : today;
            if (ended && monthlyArExpectedSnapshots[key] != null) continue;
            pending[key] = computeArExpectedForReportMonth(y, m, scopedArInvoices, regionListCurrency, asOf);
          }
          if (Object.keys(pending).length) {
            setMonthlyArExpectedSnapshots((prev) => {
              const changed = Object.keys(pending).some((k) => prev[k] !== pending[k]);
              if (!changed) return prev;
              return { ...prev, ...pending };
            });
          }
        }, [monthlyReportYear, activeRegion, scopedArInvoices, regionListCurrency, monthlyArExpectedSnapshots]);

        const monthlyReportMaxBar = useMemo(() => Math.max(1, ...monthlyReportSeries.flatMap((m) => [m.poReceived, m.invoiceIssued, m.arExpected])), [monthlyReportSeries]);

        const filteredAr = useMemo(() => {
          let list = arDisplay;
          if (arClientFilter !== "all") {
            const byCompany = clients.find((c) => c.company === arClientFilter);
            const byTitle = clients.find((c) => c.invoice_title === arClientFilter);
            const possible = new Set([arClientFilter]);
            if (byCompany?.invoice_title) possible.add(byCompany.invoice_title);
            if (byTitle?.company) possible.add(byTitle.company); // backward-compat if older AR.customer stored company
            list = list.filter((r) => possible.has(r.customer));
          }
          if (arSearch.trim()) list = list.filter((r) => matchesSearch(r, arSearch, ["invoice_no", "customer", "customer_po", "job_no"]));
          return applyTableSort(list, "ar");
        }, [arDisplay, arClientFilter, arSearch, tableSort]);
        const filteredAp = useMemo(() => {
          let list = apDisplay;
          if (apClientFilter !== "all") {
            list = list.filter((b) => {
              const job = scopedJobs.find((j) => j.id === b.job_id || j.job_no === b.job_no);
              return job ? job.company === apClientFilter : false;
            });
          }
          if (apJobFilter !== "all") {
            list = list.filter((b) => String(b.job_id) === apJobFilter || scopedJobs.find((j) => j.id === Number(apJobFilter))?.job_no === b.job_no);
          }
          if (apSearch.trim()) list = list.filter((b) => matchesSearch(b, apSearch, ["invoice_no", "company_name", "job_no", "remarks", "customer_po"]));
          return applyTableSort(list, "ap");
        }, [apDisplay, apClientFilter, apJobFilter, apSearch, scopedJobs, tableSort]);

        function getJobForAr(r) {
          if (r.job_id) return jobs.find((j) => j.id === r.job_id) || null;
          if (r.job_no) return jobs.find((j) => j.job_no === r.job_no) || null;
          return null;
        }

        function syncArWithJobStatus(jobId, jobNo, status) {
          const completed = isJobCompleted(status);
          setArInvoices((prev) => prev.map((r) => (
            r.job_id === jobId || r.job_no === jobNo ? { ...r, job_completed: completed } : r
          )));
        }

        function groupRows(rows, keyFn) {
          const map = {};
          rows.forEach((r) => {
            const k = keyFn(r) || "-";
            if (!map[k]) map[k] = [];
            map[k].push(r);
          });
          return Object.entries(map).sort((a, b) => a[0].localeCompare(b[0]));
        }

        function getJobsByClient(company) {
          return jobs.filter((j) => j.company === company);
        }

        function getArByClient(company) {
          const client = clients.find((c) => c.company === company);
          const invoiceTitle = client?.invoice_title;
          return arInvoices.filter((r) => r.customer === company || (invoiceTitle && r.customer === invoiceTitle));
        }

        function getArByJob(jobNo) {
          return arInvoices.filter((r) => r.job_no === jobNo);
        }

        function getApByJob(jobNo) {
          return apBills.filter((b) => b.job_no === jobNo);
        }

        function getApByVendor(name) {
          return apBills.filter((r) => r.company_name === name);
        }

        function syncCompanyName(oldName, newName) {
          if (!oldName || oldName === newName) return;
          setJobs((prev) => prev.map((j) => (j.company === oldName ? { ...j, company: newName } : j)));
          setArInvoices((prev) => prev.map((r) => (r.customer === oldName ? { ...r, customer: newName } : r)));
        }

        function calcBaseEditable(patch, form, setForm) {
          const next = { ...form, ...patch };
          const amtRaw = next.amount !== undefined && next.amount !== "" ? next.amount : next.invoice_amt;
          const a = Number(amtRaw);
          const r = Number(next.exchange_rate);
          if (amtRaw !== "" && amtRaw != null && !Number.isNaN(a) && !Number.isNaN(r) && r !== 0) {
            next.base_amount = (a * r).toFixed(2);
          }
          setForm(next);
        }

        function saveClient(e) {
          e.preventDefault();
          const data = {
            ...clientModal.data,
            is_bu: !!clientModal.data.is_bu,
            bu_no: clientModal.data.is_bu ? String(clientModal.data.bu_no || "").trim() : ""
          };
          if (clientModal.mode === "add") {
            if (!guardPermission("clients", "add")) return;
            setClients([{ id: nextErpRecordId(clients), region: regionForNewRecord(), ...data }, ...clients]);
            logAudit("clients", "create", data.company, `Created client ${data.customer_no} · ${data.company}`);
          } else {
            if (!guardPermission("clients", "edit")) return;
            const old = clients.find((c) => c.id === clientModal.id);
            setClients(clients.map((c) => (c.id === clientModal.id ? { ...c, ...data } : c)));
            if (old) syncCompanyName(old.company, data.company);
            logAudit("clients", "update", data.company, `Updated client ${data.customer_no} · ${data.company}`);
          }
          setClientModal(null);
        }

        function deleteClient(id) {
          if (!guardPermission("clients", "delete")) return;
          const c = clients.find((x) => x.id === id);
          if (!window.confirm(t("confirmDelete"))) return;
          setClients(clients.filter((x) => x.id !== id));
          logAudit("clients", "delete", c ? c.company : String(id), `Deleted client ${c ? c.company : id}`);
        }

        function saveJob(e) {
          e.preventDefault();
          const po_lines = (jobModal.data.po_lines || []).map((l) => ({
            ...l,
            id: l.id || nextErpRecordId(jobModal.data.po_lines, jobs),
            amount: Number(l.amount || 0)
          }));
          const data = { ...jobModal.data, po_lines, client_id: Number(jobModal.data.client_id), amount: Number(jobModal.data.amount || 0), customer_po_amount: Number(jobModal.data.customer_po_amount || 0) };
          const client = clients.find((c) => c.id === data.client_id);
          const payload = normalizeJobQuotations(normalizeJobPoFields({ ...data, company: client ? client.company : data.company }));
          const oldJob = jobModal.mode === "edit" ? jobs.find((j) => j.id === jobModal.id) : null;
          const oldQuotationNos = oldJob ? jobQuotationNos(oldJob) : [];
          const newQuotationNos = jobQuotationNos(payload);
          for (const qNo of newQuotationNos) {
            const q = findQuotationByNo(qNo, quotationsForJobRegion);
            const check = canLinkQuotationToJob(q, jobModal, jobs);
            if (!check.ok) {
              if (check.key === "quotationNotAccepted") alert(t("quotationNotAccepted"));
              else alert(t("quotationNotFound"));
              return;
            }
          }
          if (jobModal.mode === "add") {
            if (!guardPermission("job", "add")) return;
            if (jobs.some((j) => String(j.job_no || "").trim().toLowerCase() === String(payload.job_no || "").trim().toLowerCase())) {
              alert(t("duplicateJobNo").replace("{no}", payload.job_no));
              return;
            }
            const newId = nextErpRecordId(jobs, quotations);
            setJobs([{ id: newId, region: payload.region || regionForNewRecord(), ...payload }, ...jobs]);
            syncQuotationsForJob(newId, payload.job_no, newQuotationNos, []);
            logAudit("job", "create", payload.job_no, `Created job ${payload.job_no} for ${payload.company}`);
          } else {
            if (!guardPermission("job", "edit")) return;
            const old = jobs.find((j) => j.id === jobModal.id);
            setJobs(jobs.map((j) => (j.id === jobModal.id ? { ...j, ...payload } : j)));
            if (old && old.job_no !== payload.job_no) {
              setArInvoices((prev) => prev.map((r) => (r.job_no === old.job_no ? { ...r, job_no: payload.job_no, job_id: jobModal.id } : r)));
            }
            if (old && old.company !== payload.company) {
              const client = clients.find((c) => c.id === payload.client_id);
              const customerTitle = client?.invoice_title || payload.company;
              setArInvoices((prev) => prev.map((r) => (r.job_no === payload.job_no ? { ...r, customer: customerTitle } : r)));
            }
            syncArWithJobStatus(jobModal.id, payload.job_no, payload.status);
            syncQuotationsForJob(jobModal.id, payload.job_no, newQuotationNos, oldQuotationNos);
            logAudit("job", "update", payload.job_no, `Updated job ${payload.job_no}`);
          }
          setJobModal(null);
        }

        function syncQuotationsForJob(jobId, jobNo, newNos, oldNos) {
          const jNo = String(jobNo || "").trim();
          if (!jobId || !jNo) return;
          const newSet = new Set((newNos || []).map((n) => String(n).trim().toLowerCase()).filter(Boolean));
          const oldSet = new Set((oldNos || []).map((n) => String(n).trim().toLowerCase()).filter(Boolean));
          const added = [...newSet].filter((k) => !oldSet.has(k));
          const removed = [...oldSet].filter((k) => !newSet.has(k));
          if (!added.length && !removed.length) return;
          setQuotations((prev) => prev.map((q) => {
            const qKey = String(q.quotation_no || "").toLowerCase();
            let next = { ...q };
            if (removed.includes(qKey)) next = removeJobFromQuotation(next, jobId, jNo);
            if (added.includes(qKey)) next = addJobToQuotation(next, jobId, jNo);
            return next;
          }));
        }

        function addQuotationDraftToJob() {
          if (!jobModal) return;
          const trimmed = String(jobModal.data._quotationDraft || "").trim();
          if (!trimmed) return;
          const q = findQuotationByNo(trimmed, quotationsForJobRegion) || findQuotationByNo(trimmed, accessibleQuotations);
          const check = canLinkQuotationToJob(q, jobModal, jobs);
          if (!check.ok) {
            if (check.key === "quotationNotAccepted") alert(t("quotationNotAccepted"));
            else alert(t("quotationNotFound"));
            return;
          }
          const existing = jobQuotationNos(jobModal.data);
          if (existing.some((n) => n.toLowerCase() === trimmed.toLowerCase())) {
            setJobModal({ ...jobModal, data: { ...jobModal.data, _quotationDraft: "" } });
            return;
          }
          const withQ = addQuotationToJob(jobModal.data, trimmed);
          const patch = { ...withQ, _quotationDraft: "" };
          if (!existing.length) Object.assign(patch, { region: q?.region || jobModal.data.region || newRecordRegion }, fieldsFromQuotation(q));
          setJobModal({ ...jobModal, data: { ...jobModal.data, ...patch } });
        }

        function removeQuotationFromJobModal(quotationNo) {
          if (!jobModal) return;
          setJobModal({ ...jobModal, data: removeQuotationFromJob(jobModal.data, quotationNo) });
        }

        function applyQuotationNoToJob(quotationNo) {
          if (!jobModal) return;
          const trimmed = String(quotationNo || "").trim();
          if (!trimmed) return;
          const q = findQuotationByNo(trimmed, quotationsForJobRegion) || findQuotationByNo(trimmed, accessibleQuotations);
          const check = canLinkQuotationToJob(q, jobModal, jobs);
          if (!check.ok) {
            if (check.key === "quotationNotAccepted") alert(t("quotationNotAccepted"));
            else alert(t("quotationNotFound"));
            return;
          }
          const existing = jobQuotationNos(jobModal.data);
          if (existing.some((n) => n.toLowerCase() === trimmed.toLowerCase())) return;
          const withQ = addQuotationToJob(jobModal.data, trimmed);
          const patch = { ...withQ, _quotationDraft: "" };
          if (!existing.length) Object.assign(patch, { region: q?.region || jobModal.data.region || newRecordRegion }, fieldsFromQuotation(q));
          setJobModal({ ...jobModal, data: { ...jobModal.data, ...patch } });
        }

        function deleteJob(id) {
          if (!guardPermission("job", "delete")) return;
          const j = jobs.find((x) => x.id === id);
          if (!window.confirm(t("confirmDelete"))) return;
          if (j) {
            jobQuotationNos(j).forEach((qNo) => {
              setQuotations((prev) => prev.map((q) => {
                if (String(q.quotation_no || "").toLowerCase() !== qNo.toLowerCase()) return q;
                return removeJobFromQuotation(q, j.id, j.job_no);
              }));
            });
          }
          setJobs(jobs.filter((x) => x.id !== id));
          logAudit("job", "delete", j ? j.job_no : String(id), `Deleted job ${j ? j.job_no : id}`);
        }

        function startOngoingEdit(job) {
          if (!can("job", "edit")) return;
          setOngoingEditId(job.id);
          setOngoingDraft({
            billedAmt: job.ongoing_billed_amt != null && job.ongoing_billed_amt !== "" ? String(job.ongoing_billed_amt) : "",
            remarks: job.remarks || ""
          });
        }

        function saveOngoingRow(jobId) {
          if (!guardPermission("job", "edit")) return;
          const job = jobs.find((j) => j.id === jobId);
          if (!job) return;
          const billed = ongoingDraft.billedAmt === "" ? null : Number(ongoingDraft.billedAmt);
          if (billed != null && Number.isNaN(billed)) return;
          setJobs(jobs.map((j) => (j.id === jobId ? { ...j, ongoing_billed_amt: billed, remarks: ongoingDraft.remarks || "" } : j)));
          setOngoingEditId(null);
          logAudit("ongoing", "update", job.job_no, `Updated ongoing billed & remarks for ${job.job_no}`);
        }

        function cancelOngoingEdit() {
          setOngoingEditId(null);
        }

        function syncMonthlyPoFromJobs() {
          if (!guardPermission("monthly_report", "edit")) return;
          setMonthlyPoLines((prev) => {
            const next = syncPoLinesFromJobs(scopedJobs, monthlyReportYear, regionListCurrency, prev);
            if (!saveErpStorage("erp_monthly_po_lines", next)) alert(t("storageSaveFailed"));
            else localStorage.setItem("erp_monthly_po_saved_at", new Date().toISOString());
            return next;
          });
          logAudit("monthly_report", "update", "PO Received", `Synced PO received lines for ${monthlyReportYear}`);
          scheduleLiveSyncPush(true);
          alert(t("monthlySyncSaved"));
        }

        function syncMonthlyArFromInvoices() {
          if (!guardPermission("monthly_report", "edit")) return;
          setMonthlyArLines((prev) => {
            const next = syncArLinesFromInvoices(scopedArInvoices, monthlyReportYear, regionListCurrency, prev);
            if (!saveErpStorage("erp_monthly_ar_lines", next)) alert(t("storageSaveFailed"));
            else localStorage.setItem("erp_monthly_ar_saved_at", new Date().toISOString());
            return next;
          });
          logAudit("monthly_report", "update", "AR Invoice", `Synced AR invoice lines for ${monthlyReportYear}`);
          scheduleLiveSyncPush(true);
          alert(t("monthlySyncSaved"));
        }

        function refreshMonthlyArExpectedSnapshots() {
          const today = new Date().toISOString().slice(0, 10);
          const y = Number(monthlyReportYear);
          setMonthlyArExpectedSnapshots((prev) => {
            const updates = { ...prev };
            let changed = false;
            for (let m = 0; m < 12; m++) {
              const key = arExpectedSnapshotKey(y, m, activeRegion);
              const ended = reportMonthHasEnded(y, m, today);
              if (ended && prev[key] != null) continue;
              const asOf = ended ? monthEndIso(y, m) : today;
              const val = computeArExpectedForReportMonth(y, m, scopedArInvoices, regionListCurrency, asOf);
              if (updates[key] !== val) {
                updates[key] = val;
                changed = true;
              }
            }
            return changed ? updates : prev;
          });
        }

        function syncMonthlyReportAll() {
          if (!guardPermission("monthly_report", "edit")) return;
          setMonthlyPoLines((prev) => {
            const next = syncPoLinesFromJobs(scopedJobs, monthlyReportYear, regionListCurrency, prev);
            if (!saveErpStorage("erp_monthly_po_lines", next)) alert(t("storageSaveFailed"));
            else localStorage.setItem("erp_monthly_po_saved_at", new Date().toISOString());
            return next;
          });
          setMonthlyArLines((prev) => {
            const next = syncArLinesFromInvoices(scopedArInvoices, monthlyReportYear, regionListCurrency, prev);
            if (!saveErpStorage("erp_monthly_ar_lines", next)) alert(t("storageSaveFailed"));
            else localStorage.setItem("erp_monthly_ar_saved_at", new Date().toISOString());
            return next;
          });
          refreshMonthlyArExpectedSnapshots();
          logAudit("monthly_report", "update", "Report", `Synced full monthly report for ${monthlyReportYear}`);
          scheduleLiveSyncPush(true);
          alert(t("monthlySyncReportDone"));
        }

        function clearAllMonthlyReportData() {
          if (!isRoot()) { alert(t("rootOnly")); return; }
          if (!window.confirm(t("monthlyClearAllConfirm"))) return;
          setMonthlyPoLines([]);
          setMonthlyArLines([]);
          setMonthlyArExpectedSnapshots({});
          saveErpStorage("erp_monthly_po_lines", []);
          saveErpStorage("erp_monthly_ar_lines", []);
          saveErpStorage("erp_monthly_ar_expected", {});
          logAudit("monthly_report", "delete", "all", "Cleared all monthly report data");
          scheduleLiveSyncPush(true);
          setImportStatus(t("monthlyClearAllDone"));
          alert(t("monthlyClearAllDone"));
        }

        function updateMonthlyPoOtherCurrency(id, value) {
          if (!can("monthly_report", "edit")) return;
          setMonthlyPoLines((prev) => {
            const next = prev.map((l) => (l.id === id ? { ...l, other_currency: value } : l));
            saveErpStorage("erp_monthly_po_lines", next);
            return next;
          });
        }

        function updateMonthlyArOtherCurrency(id, value) {
          if (!can("monthly_report", "edit")) return;
          setMonthlyArLines((prev) => {
            const next = prev.map((l) => (l.id === id ? { ...l, other_currency: value } : l));
            saveErpStorage("erp_monthly_ar_lines", next);
            return next;
          });
        }

        function exportMonthlyReportPdfClick() {
          if (!guardPermission("monthly_report", "export")) return;
          const title = t("monthlyReportChartTitle").replace("{region}", activeRegion === "ALL" ? t("allRegions") : regionLabel(activeRegion, lang)).replace("{year}", monthlyReportYear);
          ensurePdfLibs().then(function (ok) {
            if (!ok) { alert("PDF library not loaded"); return; }
            exportMonthlyReportPdf({ series: monthlyReportSeries, title, t, moneyFn: money });
            logAudit("monthly_report", "export", title, "Exported monthly report PDF");
          });
        }

        function exportMonthlyDetailExcel(kind) {
          if (!guardPermission("monthly_report", "export")) return;
          if (typeof XLSX === "undefined") { alert("Excel library not loaded"); return; }
          const lines = kind === "po" ? scopedMonthlyPoLines : scopedMonthlyArLines;
          const rows = lines.slice().sort((a, b) => (a.month - b.month) || (a.job_no || "").localeCompare(b.job_no || "")).map((l) => ({
            month: l.month,
            job_no: l.job_no,
            other_currency: l.other_currency,
            source_amount: l.source_amount ?? "",
            amount: l.base_amount,
            currency: l.base_currency || regionListCurrency
          }));
          const schema = {
            sheetName: kind === "po" ? "PO Received" : "AR Invoice",
            filePrefix: (kind === "po" ? "po_received_" : "ar_invoice_") + monthlyReportYear,
            columns: [
              { header: "Month", field: "month" },
              { header: "Job#", field: "job_no" },
              { header: "Currency", field: "other_currency" },
              { header: "Source Amount", field: "source_amount" },
              { header: "Amount", field: "amount" },
              { header: "Currency", field: "currency" }
            ]
          };
          downloadWorkbook(buildWorkbook(schema, rows, false), schema.filePrefix + ".xlsx");
          logAudit("monthly_report", "export", schema.sheetName, `Exported ${rows.length} rows`);
        }

        function saveQuotation(e) {
          e.preventDefault();
          const data = { ...quotationModal.data, client_id: Number(quotationModal.data.client_id), amount: Number(quotationModal.data.amount || 0) };
          const client = clients.find((c) => c.id === data.client_id);
          const job = data.job_no ? jobs.find((j) => j.job_no === data.job_no) : null;
          const payload = {
            ...data,
            company: client ? client.company : data.company,
            job_id: job ? job.id : null,
            job_no: job ? job.job_no : (data.job_no || "")
          };
          const currentQuotation = quotationModal.mode === "edit" ? quotations.find((q) => q.id === quotationModal.id) : null;
          const oldStatus = currentQuotation ? currentQuotation.status : null;
          const oldJobNo = currentQuotation ? currentQuotation.job_no || "" : "";
          const oldQuotationNo = currentQuotation ? currentQuotation.quotation_no || "" : "";
          const dupQuotationNo = quotations.find((q) => q.id !== quotationModal.id && String(q.quotation_no || "").trim().toLowerCase() === String(payload.quotation_no || "").trim().toLowerCase());
          if (dupQuotationNo) {
            alert(t("duplicateQuotationNo").replace("{no}", payload.quotation_no || ""));
            return;
          }
          if (quotationModal.mode === "add") {
            if (!guardPermission("quotation", "add")) return;
            const newId = nextErpRecordId(jobs, quotations);
            let qRec = { id: newId, region: regionForNewRecord(), ...payload };
            if (payload.job_id && payload.job_no) qRec = addJobToQuotation(qRec, payload.job_id, payload.job_no);
            setQuotations([qRec, ...quotations]);
            setJobs((prev) => syncJobsOnQuotationSave(prev, payload.job_id, payload.job_no, payload.quotation_no, "", ""));
            logAudit("quotation", "create", payload.quotation_no, `Created quotation ${payload.quotation_no}`);
          } else {
            if (!guardPermission("quotation", "edit")) return;
            let qRec = { ...quotations.find((q) => q.id === quotationModal.id), ...payload };
            if (payload.job_id && payload.job_no) qRec = addJobToQuotation(qRec, payload.job_id, payload.job_no);
            setQuotations(quotations.map((q) => (q.id === quotationModal.id ? qRec : q)));
            setJobs((prev) => syncJobsOnQuotationSave(prev, payload.job_id, payload.job_no, payload.quotation_no, oldJobNo, oldQuotationNo));
            logAudit("quotation", "update", payload.quotation_no, `Updated quotation ${payload.quotation_no}`);
          }
          const shouldAutoOpenJob = payload.status === "Accepted" && oldStatus !== "Accepted" && !String(payload.job_no || "").trim();
          if (shouldAutoOpenJob && can("job", "add")) {
            const modalJobRegion = currentQuotation?.region || client?.region || newRecordRegion;
            setJobModal({
              mode: "add",
              data: {
                ...emptyJob(),
                po_lines: [emptyPoLine()],
                region: modalJobRegion,
                ...fieldsFromQuotation({ ...payload, company: payload.company })
              }
            });
          }
          setQuotationModal(null);
        }

        function deleteQuotation(id) {
          if (!guardPermission("quotation", "delete")) return;
          const q = quotations.find((x) => x.id === id);
          if (!window.confirm(t("confirmDelete"))) return;
          if (q) {
            const qNo = String(q.quotation_no || "").trim();
            quotationJobNos(q).forEach((jno) => {
              setJobs((prev) => prev.map((j) => (j.job_no === jno ? removeQuotationFromJob(j, qNo) : j)));
            });
          }
          setQuotations(quotations.filter((x) => x.id !== id));
          logAudit("quotation", "delete", q ? q.quotation_no : String(id), `Deleted quotation ${q ? q.quotation_no : id}`);
        }

        function saveVendor(e) {
          e.preventDefault();
          const data = vendorModal.data;
          if (vendorModal.mode === "add") {
            if (!guardPermission("vendors", "add")) return;
            setVendors([{ id: nextErpRecordId(vendors), region: regionForNewRecord(), ...data }, ...vendors]);
            logAudit("vendors", "create", data.vendor_no, `Created vendor ${data.vendor_no} · ${data.name}`);
          } else {
            if (!guardPermission("vendors", "edit")) return;
            const old = vendors.find((v) => v.id === vendorModal.id);
            setVendors(vendors.map((v) => (v.id === vendorModal.id ? { ...v, ...data } : v)));
            if (old && old.name !== data.name) {
              setApBills((prev) => prev.map((b) => (b.company_name === old.name && b.payee_type === "Vendor" ? { ...b, company_name: data.name } : b)));
            }
            logAudit("vendors", "update", data.vendor_no, `Updated vendor ${data.vendor_no} · ${data.name}`);
          }
          setVendorModal(null);
        }

        function deleteVendor(id) {
          if (!guardPermission("vendors", "delete")) return;
          const v = vendors.find((x) => x.id === id);
          if (!window.confirm(t("confirmDelete"))) return;
          setVendors(vendors.filter((x) => x.id !== id));
          logAudit("vendors", "delete", v ? v.name : String(id), `Deleted vendor ${v ? v.name : id}`);
        }

        function saveAR(e) {
          e.preventDefault();
          const f = arModal.data;
          const calcValue = Number(f.invoice_amt || 0) * Number(f.exchange_rate || 0);
          const baseInput = Number(f.base_amount || 0);
          const selectedJob = jobs.find((j) => j.id === Number(f.job_id));
          const rec = {
            ...f,
            job_id: Number(f.job_id) || null,
            job_no: selectedJob ? selectedJob.job_no : f.job_no,
            customer: selectedJob ? (((clients.find((c) => c.id === selectedJob.client_id) || clients.find((c) => c.company === selectedJob.company))?.invoice_title) || selectedJob.company) : f.customer,
            customer_po: f.customer_po || (selectedJob ? selectedJob.customer_po : ""),
            invoice_amt: Number(f.invoice_amt || 0),
            exchange_rate_locked: Number(f.exchange_rate || 0),
            base_amount_locked: baseInput,
            base_amount: baseInput,
            lock_timestamp: nowIso(),
            rate_source: "MANUAL",
            manual_override: Math.abs(baseInput - calcValue) > 0.01,
            job_completed: selectedJob ? isJobCompleted(selectedJob.status) : false,
            payment_received_date: f.payment_received_date || "",
            payment_status: deriveArPaymentStatus({ ...f, due_date: f.due_date || computeArDueDate(f, clients, jobs) }),
            payment_stage: f.payment_stage || "",
            due_date: f.due_date || computeArDueDate(f, clients, jobs)
          };
          const arDup = findDuplicateInvoiceNo(arInvoices, rec.invoice_no, arModal.mode === "edit" ? arModal.id : null);
          if (arDup) {
            alert(t("duplicateInvoiceNo").replace("{no}", rec.invoice_no).replace("{module}", "AR"));
            return;
          }
          if (arModal.mode === "add") {
            if (!guardPermission("ar", "add")) return;
            setArInvoices([{ id: nextErpRecordId(arInvoices), region: regionForNewRecord(), ...rec }, ...arInvoices]);
            logAudit("ar", "create", rec.invoice_no, `Created AR ${rec.invoice_no} for ${rec.customer}`);
          } else {
            if (!guardPermission("ar", "edit")) return;
            setArInvoices(arInvoices.map((r) => (r.id === arModal.id ? { ...r, ...rec } : r)));
            logAudit("ar", "update", rec.invoice_no, `Updated AR ${rec.invoice_no}`);
          }
          setArModal(null);
        }

        function deleteAr(id) {
          if (!guardPermission("ar", "delete")) return;
          const r = arInvoices.find((x) => x.id === id);
          if (!window.confirm(t("confirmDelete"))) return;
          setArInvoices(arInvoices.filter((x) => x.id !== id));
          logAudit("ar", "delete", r ? r.invoice_no : String(id), `Deleted AR ${r ? r.invoice_no : id}`);
        }

        function saveAP(e) {
          e.preventDefault();
          const f = apModal.data;
          const calcValue = Number(f.amount || 0) * Number(f.exchange_rate || 0);
          const baseInput = Number(f.base_amount || 0);
          const selectedJob = jobs.find((j) => j.id === Number(f.job_id));
          const rec = {
            ...f,
            job_id: Number(f.job_id) || null,
            job_no: selectedJob ? selectedJob.job_no : f.job_no,
            payee_type: "Vendor",
            si_id: null,
            amount: Number(f.amount || 0),
            exchange_rate_locked: Number(f.exchange_rate || 0),
            base_amount_locked: baseInput,
            base_amount: baseInput,
            lock_timestamp: nowIso(),
            rate_source: "MANUAL",
            manual_override: Math.abs(baseInput - calcValue) > 0.01,
            payment_stage: f.payment_stage || "",
            pay_date: f.pay_date || "",
            payment_status: deriveApPaymentStatus(f)
          };
          const apDup = findDuplicateInvoiceNo(apBills, rec.invoice_no, apModal.mode === "edit" ? apModal.id : null);
          if (apDup) {
            alert(t("duplicateInvoiceNo").replace("{no}", rec.invoice_no).replace("{module}", "AP"));
            return;
          }
          if (apModal.mode === "add") {
            if (!guardPermission("ap", "add")) return;
            setApBills([{ id: nextErpRecordId(apBills), region: regionForNewRecord(), ...rec }, ...apBills]);
            logAudit("ap", "create", rec.invoice_no, `Created AP ${rec.invoice_no} for ${rec.company_name}`);
          } else {
            if (!guardPermission("ap", "edit")) return;
            setApBills(apBills.map((b) => (b.id === apModal.id ? { ...b, ...rec } : b)));
            logAudit("ap", "update", rec.invoice_no, `Updated AP ${rec.invoice_no}`);
          }
          setApModal(null);
        }

        function deleteAp(id) {
          if (!guardPermission("ap", "delete")) return;
          const b = apBills.find((x) => x.id === id);
          if (!window.confirm(t("confirmDelete"))) return;
          setApBills(apBills.filter((x) => x.id !== id));
          logAudit("ap", "delete", b ? b.invoice_no : String(id), `Deleted AP ${b ? b.invoice_no : id}`);
        }

        function emptyUserForm() {
          return { name: "", email: "", password: "", role: "user", is_active: true, permissions: emptyPermissions(), allowed_regions: ["HK"] };
        }

        function toggleUserRegion(regionId) {
          if (!userModal || userModal.data.role === "root") return;
          const cur = userModal.data.allowed_regions || [];
          const next = cur.includes(regionId) ? cur.filter((x) => x !== regionId) : [...cur, regionId];
          setUserModal({ ...userModal, data: { ...userModal.data, allowed_regions: next.length ? next : ["HK"] } });
        }

        function saveUser(e) {
          e.preventDefault();
          if (!isRoot()) { alert(t("rootOnly")); return; }
          const data = userModal.data;
          if (data.role === "root") { data.permissions = fullPermissions(); data.allowed_regions = ["ALL"]; }
          const loginId = String(data.email || "").trim().toLowerCase();
          if (!loginId) { alert(t("loginFail")); return; }
          data.email = String(data.email || "").trim();
          if (userModal.mode === "add") {
            if (users.some((u) => String(u.email || "").trim().toLowerCase() === loginId)) { alert("Username already exists"); return; }
            const created = { id: nextErpRecordId(users), ...data };
            setUsers([...users, created]);
            pushUserRecordToCloud(created);
            logAudit("settings", "user_create", data.email, `Created user ${data.name} (${data.email})`);
          } else {
            const updated = { ...users.find((u) => u.id === userModal.id), ...data, id: userModal.id };
            setUsers(users.map((u) => (u.id === userModal.id ? updated : u)));
            pushUserRecordToCloud(updated);
            logAudit("settings", "user_update", data.email, `Updated user ${data.name} (${data.email})`);
          }
          setUserModal(null);
        }

        function deleteUser(userId) {
          if (!isRoot()) { alert(t("rootOnly")); return; }
          const target = users.find((u) => u.id === userId);
          if (!target) return;
          if (target.role === "root") { alert(t("cannotDeleteRoot")); return; }
          if (target.id === sessionUserId) { alert(t("cannotDeleteSelf")); return; }
          if (!window.confirm(t("confirmDeleteUser"))) return;
          setUsers((prev) => prev.filter((u) => u.id !== userId));
          pushUserDeleteToCloud(userId);
          logAudit("settings", "user_delete", target.email, `Deleted user ${target.name} (${target.email})`);
          if (userModal && userModal.id === userId) setUserModal(null);
        }

        function changeMyPassword(e) {
          e.preventDefault();
          const u = getCurrentUser();
          if (!u) return;
          const cur = String(pwForm.current || "");
          const next = String(pwForm.next || "");
          const confirm = String(pwForm.confirm || "");
          if (String(u.password || "") !== cur) { alert(t("passwordWrong")); return; }
          if (!next || next.length < 3) { alert(t("loginFail")); return; }
          if (next !== confirm) { alert(t("passwordMismatch")); return; }
          setUsers((prev) => {
            const updatedUsers = prev.map((x) => (x.id === u.id ? { ...x, password: next } : x));
            const updated = updatedUsers.find((x) => x.id === u.id);
            if (updated) pushUserRecordToCloud(updated);
            return updatedUsers;
          });
          setPwForm({ current: "", next: "", confirm: "" });
          alert(t("passwordChanged"));
        }

        function toggleUserPerm(module, perm) {
          if (!userModal || userModal.data.role === "root") return;
          const perms = { ...userModal.data.permissions };
          perms[module] = { ...perms[module], [perm]: !perms[module][perm] };
          if (perm !== "view" && perms[module][perm]) perms[module].view = true;
          if ((perm === "import" || perm === "export") && perms[module][perm]) perms[module].view = true;
          setUserModal({ ...userModal, data: { ...userModal.data, permissions: perms } });
        }

        function getExportRows(module) {
          if (module === "clients") return scopedClients.map((c) => {
            const row = rowFromSchema(EXCEL_SCHEMAS.clients, c);
            row.is_bu = c.is_bu ? "Y" : "";
            return row;
          });
          if (module === "job") return scopedJobs.map((j) => rowFromSchema(EXCEL_SCHEMAS.job, j));
          if (module === "ar") return scopedArInvoices.map((r) => {
            const due_date = r.due_date || computeArDueDate(r, clients, jobs);
            const linkedJob = jobs.find((j) => j.job_no === r.job_no);
            const status = deriveArPaymentStatus({ ...r, due_date });
            return {
              job_no: r.job_no || "",
              customer: r.customer || "",
              customer_po: r.customer_po || "",
              job_completed: r.job_completed || (linkedJob && isJobCompleted(linkedJob.status)) ? "Yes" : "No",
              invoice_no: r.invoice_no || "",
              invoice_currency: r.invoice_currency || "",
              invoice_amt: r.invoice_amt ?? "",
              base_amount_region: resolveArDisplayAmount(r, regionListCurrency),
              invoice_date: r.invoice_date || "",
              due_date: due_date || "",
              overdue_days: status === "Paid" ? 0 : calcOverdueDays(due_date),
              days: calcDaysBetween(r.invoice_date, due_date),
              payment_received_date: r.payment_received_date || ""
            };
          });
          if (module === "ap") return scopedApBills.map((b) => ({
            company_name: b.company_name || "",
            invoice_no: b.invoice_no || "",
            invoice_date: b.invoice_date || "",
            invoice_received_date: b.invoice_received_date || "",
            due_date: b.due_date || "",
            currency: b.currency || "",
            amount: b.amount ?? "",
            base_amount_region: resolveApDisplayAmount(b, regionListCurrency),
            pay_date: b.pay_date || "",
            remarks: b.remarks || "",
            bank_charge: [b.bank, b.charge].filter(Boolean).join(" / ") || "",
            swift_code: b.swift_code || "",
            payment_advice_email: b.payment_advice_email || ""
          }));
          if (module === "vendors") return scopedVendors.map((v) => rowFromSchema(EXCEL_SCHEMAS.vendors, v));
          if (module === "quotation") return scopedQuotations.map((q) => rowFromSchema(EXCEL_SCHEMAS.quotation, q));
          if (module === "monthly_report") return monthlyReportSeries.map((m) => ({
            month: m.header,
            currency: m.currency,
            po_received: m.poReceived,
            invoice_issued: m.invoiceIssued,
            ar_expected: m.arExpected
          }));
          return [];
        }

        function exportModuleExcel(module, mode) {
          if (!guardPermission(module, "export")) return;
          let schema = EXCEL_SCHEMAS[module];
          if (module === "ar") {
            schema = {
              sheetName: "AR",
              filePrefix: "ar",
              columns: [
                { header: "Job No", field: "job_no" },
                { header: "Customer", field: "customer" },
                { header: "Customer PO", field: "customer_po" },
                { header: "Job Complete", field: "job_completed" },
                { header: "Invoice No", field: "invoice_no" },
                { header: "Invoice Currency", field: "invoice_currency" },
                { header: "Invoice Amt", field: "invoice_amt" },
                { header: "Amount (" + regionListCurrency + ")", field: "base_amount_region" },
                { header: "Invoice Date", field: "invoice_date" },
                { header: "Due Date", field: "due_date" },
                { header: "Overdue Days", field: "overdue_days" },
                { header: "Days", field: "days" },
                { header: "Payment Received Date", field: "payment_received_date" }
              ]
            };
          } else if (module === "ap") {
            schema = {
              sheetName: "AP",
              filePrefix: "ap",
              columns: [
                { header: "Company Name", field: "company_name" },
                { header: "Invoice No.", field: "invoice_no" },
                { header: "Invoice Date", field: "invoice_date" },
                { header: "Invoice Received Date", field: "invoice_received_date" },
                { header: "Due Date", field: "due_date" },
                { header: "Currency", field: "currency" },
                { header: "Amount", field: "amount" },
                { header: "Amount in " + regionListCurrency, field: "base_amount_region" },
                { header: "Pay Date", field: "pay_date" },
                { header: "Remarks", field: "remarks" },
                { header: "Bank Charge", field: "bank_charge" },
                { header: "SWIFT CODE", field: "swift_code" },
                { header: "Payment Advice email", field: "payment_advice_email" }
              ]
            };
          }
          if (module === "monthly_report") {
            schema = { sheetName: "MonthlyReport", filePrefix: "monthly_report_" + monthlyReportYear, columns: [
              { header: "Month", field: "month" }, { header: "Currency", field: "currency" },
              { header: "PO Received (Current Month)", field: "po_received" },
              { header: "Invoice Issued (Current Month)", field: "invoice_issued" },
              { header: "AR Expected", field: "ar_expected" }
            ] };
          }
          if (!schema || typeof XLSX === "undefined") { alert("Excel library not loaded"); return; }
          const stamp = new Date().toISOString().slice(0, 10);
          const rows = mode === "data" ? getExportRows(module) : [];
          downloadWorkbook(buildWorkbook(schema, rows, true), schema.filePrefix + (mode === "data" ? "_data_" : "_template_") + stamp + ".xlsx");
          logAudit(module, "export", schema.sheetName, `Exported ${mode} (${rows.length} rows)`);
        }

        function findImportExisting(module, key) {
          const k = key.toLowerCase();
          if (module === "clients") return clients.find((c) => c.customer_no.toLowerCase() === k);
          if (module === "job") return jobs.find((j) => j.job_no.toLowerCase() === k);
          if (module === "ar") return arInvoices.find((r) => r.invoice_no.toLowerCase() === k);
          if (module === "ap") return apBills.find((b) => b.invoice_no.toLowerCase() === k);
          if (module === "vendors") return vendors.find((v) => v.vendor_no.toLowerCase() === k);
          if (module === "quotation") return quotations.find((q) => q.quotation_no.toLowerCase() === k);
          return null;
        }

        function validateImportRow(module, row, errors) {
          if (module === "job" || module === "quotation") {
            const company = String(row.company || "").trim();
            const client = clients.find((c) => c.company.toLowerCase() === company.toLowerCase());
            if (!client) errors.push("Client not found: " + (company || "(empty)"));
          }
          if (module === "job") {
            const jt = String(row.job_type || "").trim();
            if (jt && !JOB_TYPES.includes(jt)) errors.push("Invalid Type: " + jt);
            if (!String(row.status || "").trim()) row.status = "Open";
            const qNo = String(row.quotation_no || "").trim();
            if (qNo) {
              const q = quotations.find((x) => String(x.quotation_no || "").trim().toLowerCase() === qNo.toLowerCase());
              if (q && q.status !== "Accepted") errors.push("Quotation not Accepted: " + qNo);
            }
          }
          if (module === "quotation") {
            const st = String(row.status || "").trim();
            if (st && !QUOTATION_STATUSES.includes(st)) errors.push("Invalid status: " + st);
            if (!st) row.status = "Draft";
            if (row.job_no) {
              const job = jobs.find((j) => j.job_no.toLowerCase() === String(row.job_no).toLowerCase());
              if (!job) errors.push("Job not found: " + row.job_no);
            }
          }
          if (module === "ar" || module === "ap") {
            const jobNo = String(row.job_no || "").trim();
            const job = jobs.find((j) => j.job_no.toLowerCase() === jobNo.toLowerCase());
            if (!job) errors.push("Job not found: " + (jobNo || "(empty)"));
          }
          if (module === "ap") {
            const pt = String(row.payee_type || "Vendor").trim();
            if (pt && pt !== "Vendor") row.payee_type = "Vendor";
          }
        }

        function buildImportPreview(module, parsedRows) {
          const schema = EXCEL_SCHEMAS[module];
          const preview = [];
          const seen = new Set();
          parsedRows.forEach((row, idx) => {
            const lineNo = idx + 1;
            if (module === "vendors" && !String(row.vendor_no || "").trim() && row.name) {
              row.vendor_no = "VEND-" + String(row.name).trim().replace(/\s+/g, "-").slice(0, 24).toUpperCase();
            }
            if (module === "clients") {
              row.is_bu = parseBuFlag(row.is_bu);
              if (!row.is_bu) row.bu_no = "";
            }
            const key = String(row[schema.keyField] || "").trim();
            const errors = [];
            schema.columns.filter((c) => c.required).forEach((c) => {
              if (!String(row[c.field] || "").trim()) errors.push(c.header + " required");
            });
            if (!key) errors.push(schema.keyField + " required");
            if (key && seen.has(key.toLowerCase())) errors.push("Duplicate key in file: " + key);
            if (key) seen.add(key.toLowerCase());
            validateImportRow(module, row, errors);
            if (errors.length) {
              preview.push({ lineNo, action: "error", key, data: row, errors });
              return;
            }
            const existing = findImportExisting(module, key);
            preview.push({ lineNo, action: existing ? "update" : "create", key, data: row, errors: [] });
          });
          return preview;
        }

        function openImportPicker(module) {
          importModuleRef.current = module;
          const el = importFileInputRef.current;
          if (!el) {
            setImportStatus(t("importFileError"));
            return;
          }
          el.value = "";
          el.click();
          setImportPreview(null);
        }

        function onImportFileSelected(e) {
          const f = e.target.files && e.target.files[0];
          const module = importModuleRef.current;
          e.target.value = "";
          if (f && module) handleImportFile(module, f);
        }

        function handleImportFile(module, file) {
          if (!file) return;
          setImportStatus(t("importLoading") + " (" + file.name + ")");
          if (!can(module, "import")) {
            setImportPreview({ module, preview: [], fileName: file.name, parseErrors: [t("noPermission")] });
            setImportStatus(t("noPermission"));
            return;
          }
          if (typeof XLSX === "undefined") {
            setImportPreview({ module, preview: [], fileName: file.name, parseErrors: [t("importXlsxMissing")] });
            setImportStatus(t("importXlsxMissing"));
            return;
          }
          const schema = EXCEL_SCHEMAS[module];
          if (!schema) {
            setImportPreview({ module, preview: [], fileName: file.name, parseErrors: ["Unknown import module: " + module] });
            setImportStatus("Import error");
            return;
          }
          setImportLoading(true);
          const reader = new FileReader();
          reader.onerror = () => {
            setImportLoading(false);
            setImportPreview({ module, preview: [], fileName: file.name, parseErrors: [t("importFileError")] });
            setImportStatus(t("importFileError"));
          };
          reader.onload = (e) => {
            try {
              const { rows, errors } = parseExcelRows(e.target.result, schema);
              setImportLoading(false);
              if (errors.length) {
                setImportPreview({ module, preview: [], fileName: file.name, parseErrors: errors });
                setImportStatus(errors[0] || t("importFileError"));
                return;
              }
              if (!rows.length) {
                setImportPreview({ module, preview: [], fileName: file.name, parseErrors: [t("importNoRows")] });
                setImportStatus(t("importNoRows"));
                return;
              }
              const preview = buildImportPreview(module, rows);
              setImportPreview({ module, preview, fileName: file.name, parseErrors: [] });
              setImportStatus(t("importPreviewReady"));
            } catch (err) {
              setImportLoading(false);
              setImportPreview({
                module,
                preview: [],
                fileName: file.name,
                parseErrors: [t("importParseFailed"), String(err.message || err)]
              });
              setImportStatus(t("importParseFailed"));
            }
          };
          reader.readAsArrayBuffer(file);
        }

        function persistErpLists(snapshot) {
          if (cloudOnlyMode) return;
          if (snapshot.clients != null) { localStorage.setItem("erp_clients", JSON.stringify(snapshot.clients)); }
          if (snapshot.jobs != null) { localStorage.setItem("erp_jobs", JSON.stringify(snapshot.jobs)); }
          if (snapshot.quotations != null) { localStorage.setItem("erp_quotations", JSON.stringify(snapshot.quotations)); }
          if (snapshot.vendors != null) { localStorage.setItem("erp_vendors", JSON.stringify(snapshot.vendors)); }
          if (snapshot.arInvoices != null) { localStorage.setItem("erp_ar", JSON.stringify(snapshot.arInvoices)); }
          if (snapshot.apBills != null) { localStorage.setItem("erp_ap", JSON.stringify(snapshot.apBills)); }
        }

        function buildErpBackupPayload() {
          return {
            version: 1,
            exportedAt: new Date().toISOString(),
            app: "ERP System",
            data: {
              users, auditLogs, clients, jobs, quotations, vendors,
              arInvoices, apBills, monthlyPoLines, monthlyArLines, monthlyArExpectedSnapshots,
              settings: buildSharedSettings()
            }
          };
        }

        function applyErpBackupData(d) {
          if (!d || !Array.isArray(d.clients) || !Array.isArray(d.jobs)) return false;
          setUsers(d.users || users);
          setAuditLogs(d.auditLogs || []);
          setClients(ensureRegionOnRecords(d.clients));
          setJobs(ensureJobsQuotations(ensureRegionOnRecords(d.jobs)));
          setQuotations(ensureRegionOnRecords(d.quotations || []));
          setVendors(ensureVendorNumbers(ensureRegionOnRecords(d.vendors || [])));
          setArInvoices(ensureRegionOnRecords(d.arInvoices || []));
          setApBills(ensureRegionOnRecords(d.apBills || []));
          setMonthlyPoLines(d.monthlyPoLines || []);
          setMonthlyArLines(d.monthlyArLines || []);
          setMonthlyArExpectedSnapshots(d.monthlyArExpectedSnapshots || {});
          localStorage.setItem("erp_users", JSON.stringify(d.users || users));
          localStorage.setItem("erp_audit_logs", JSON.stringify(d.auditLogs || []));
          localStorage.setItem("erp_clients", JSON.stringify(ensureRegionOnRecords(d.clients)));
          localStorage.setItem("erp_jobs", JSON.stringify(ensureJobsQuotations(ensureRegionOnRecords(d.jobs))));
          localStorage.setItem("erp_quotations", JSON.stringify(ensureRegionOnRecords(d.quotations || [])));
          localStorage.setItem("erp_vendors", JSON.stringify(ensureVendorNumbers(ensureRegionOnRecords(d.vendors || []))));
          localStorage.setItem("erp_ar", JSON.stringify(ensureRegionOnRecords(d.arInvoices || [])));
          localStorage.setItem("erp_ap", JSON.stringify(ensureRegionOnRecords(d.apBills || [])));
          saveErpStorage("erp_monthly_po_lines", d.monthlyPoLines || []);
          saveErpStorage("erp_monthly_ar_lines", d.monthlyArLines || []);
          saveErpStorage("erp_monthly_ar_expected", d.monthlyArExpectedSnapshots || {});
          if (d.settings) applySharedSettings(d.settings);
          return true;
        }

        async function pushCloudBackup() {
          if (!isRoot()) { alert(t("rootOnly")); return; }
          if (!erpSyncKey.trim()) { alert(t("cloudSyncKey") + "?"); return; }
          try {
            const res = await fetch("/api/backup", {
              method: "POST",
              headers: { "Content-Type": "application/json", "X-ERP-Sync-Key": erpSyncKey.trim() },
              body: JSON.stringify(buildErpBackupPayload())
            });
            const body = await res.json().catch(() => ({}));
            if (!res.ok) throw new Error(body.detail || res.statusText);
            logAudit("settings", "export", "cloud-backup", "Pushed ERP backup to Supabase");
            alert(t("cloudPushDone"));
          } catch (err) {
            alert(t("cloudSyncFail") + ": " + (err.message || ""));
          }
        }

        async function pullCloudBackup() {
          if (!isRoot()) { alert(t("rootOnly")); return; }
          if (!erpSyncKey.trim()) { alert(t("cloudSyncKey") + "?"); return; }
          if (!window.confirm(t("restoreConfirm"))) return;
          try {
            const res = await fetch("/api/backup", {
              headers: { "X-ERP-Sync-Key": erpSyncKey.trim() }
            });
            const body = await res.json().catch(() => ({}));
            if (res.status === 404) { alert(t("cloudNoBackup")); return; }
            if (!res.ok) throw new Error(body.detail || res.statusText);
            const payload = body.payload || body;
            const d = payload.data;
            if (!applyErpBackupData(d)) { alert(t("restoreInvalid")); return; }
            logAudit("settings", "import", "cloud-backup", "Pulled ERP backup from Supabase");
            alert(t("cloudPullDone"));
            window.location.reload();
          } catch (err) {
            alert(t("cloudSyncFail") + ": " + (err.message || ""));
          }
        }

        function liveSyncHeaders() {
          return { "Content-Type": "application/json", "X-ERP-Sync-Key": erpSyncKey.trim() };
        }

        async function pushUserRecordToCloud(userRecord) {
          if (!userRecord || userRecord.id == null) return;
          let key = erpSyncKey.trim();
          if (!key) {
            try {
              const cfgRes = await fetch("/api/sync/client-config");
              const cfg = cfgRes.ok ? await cfgRes.json() : {};
              key = cfg.sync_key || "";
            } catch { return; }
          }
          if (!key) return;
          try {
            await fetch("/api/sync/push", {
              method: "POST",
              headers: { "Content-Type": "application/json", "X-ERP-Sync-Key": key },
              body: JSON.stringify({
                changes: [{
                  entity_type: "users",
                  entity_id: userRecord.id,
                  payload: userRecord,
                  is_deleted: false,
                  region: null
                }],
                updated_by: sessionUserId
              })
            });
          } catch { /* ignore */ }
        }

        async function pushUserDeleteToCloud(userId) {
          if (userId == null) return;
          let key = erpSyncKey.trim();
          if (!key) {
            try {
              const cfgRes = await fetch("/api/sync/client-config");
              const cfg = cfgRes.ok ? await cfgRes.json() : {};
              key = cfg.sync_key || "";
            } catch { return; }
          }
          if (!key) return;
          try {
            await fetch("/api/sync/push", {
              method: "POST",
              headers: { "Content-Type": "application/json", "X-ERP-Sync-Key": key },
              body: JSON.stringify({
                changes: [{
                  entity_type: "users",
                  entity_id: userId,
                  payload: {},
                  is_deleted: true,
                  region: null
                }],
                updated_by: sessionUserId
              })
            });
          } catch { /* ignore */ }
        }

        function stripSyncMeta(rec) {
          if (!rec || typeof rec !== "object") return rec;
          const copy = { ...rec };
          delete copy._syncAt;
          return copy;
        }

        function recordSnapshot(rec) {
          return JSON.stringify(stripSyncMeta(rec));
        }

        function computeChangesForArray(entityType, records) {
          const prev = liveSyncedSnapshotRef.current[entityType] || {};
          const nextIds = {};
          const changes = [];
          (records || []).forEach((r) => {
            if (!r || r.id == null) return;
            nextIds[r.id] = recordSnapshot(r);
            if (prev[r.id] !== nextIds[r.id]) {
              changes.push({ entity_type: entityType, entity_id: r.id, payload: stripSyncMeta(r), is_deleted: false, region: r.region || null });
            }
          });
          Object.keys(prev).forEach((id) => {
            if (!nextIds[id]) {
              changes.push({ entity_type: entityType, entity_id: Number(id), payload: {}, is_deleted: true, region: null });
            }
          });
          return changes;
        }

        function markSyncedArray(entityType, records) {
          const next = {};
          (records || []).forEach((r) => {
            if (!r || r.id == null) return;
            next[r.id] = recordSnapshot(r);
          });
          liveSyncedSnapshotRef.current[entityType] = next;
        }

        function markSyncedSingleton(entityType, payload) {
          liveSyncedSnapshotRef.current[entityType] = JSON.stringify(payload || {});
        }

        function markAllSyncedFromState() {
          markSyncedArray("clients", clients);
          markSyncedArray("jobs", jobs);
          markSyncedArray("quotations", quotations);
          markSyncedArray("vendors", vendors);
          markSyncedArray("ar_invoices", arInvoices);
          markSyncedArray("ap_bills", apBills);
          markSyncedArray("users", users);
          markSyncedArray("audit_logs", auditLogs);
          markSyncedArray("monthly_po_lines", monthlyPoLines);
          markSyncedArray("monthly_ar_lines", monthlyArLines);
          markSyncedSingleton("monthly_ar_expected", monthlyArExpectedSnapshots);
          markSyncedSingleton("settings", buildSharedSettings());
        }

        function markSyncedFromSnapshot(snap) {
          if (!snap) return;
          if (snap.clients) markSyncedArray("clients", snap.clients);
          if (snap.jobs) markSyncedArray("jobs", snap.jobs);
          if (snap.quotations) markSyncedArray("quotations", snap.quotations);
          if (snap.vendors) markSyncedArray("vendors", snap.vendors);
          if (snap.arInvoices) markSyncedArray("ar_invoices", snap.arInvoices);
          if (snap.apBills) markSyncedArray("ap_bills", snap.apBills);
        }

        function collectChangesFromSnapshot(snap) {
          if (!snap) return [];
          const changes = [];
          if (snap.clients) changes.push(...computeChangesForArray("clients", snap.clients));
          if (snap.jobs) changes.push(...computeChangesForArray("jobs", snap.jobs));
          if (snap.quotations) changes.push(...computeChangesForArray("quotations", snap.quotations));
          if (snap.vendors) changes.push(...computeChangesForArray("vendors", snap.vendors));
          if (snap.arInvoices) changes.push(...computeChangesForArray("ar_invoices", snap.arInvoices));
          if (snap.apBills) changes.push(...computeChangesForArray("ap_bills", snap.apBills));
          return changes;
        }

        function collectAllLiveChanges() {
          const changes = [];
          changes.push(...computeChangesForArray("clients", clients));
          changes.push(...computeChangesForArray("jobs", jobs));
          changes.push(...computeChangesForArray("quotations", quotations));
          changes.push(...computeChangesForArray("vendors", vendors));
          changes.push(...computeChangesForArray("ar_invoices", arInvoices));
          changes.push(...computeChangesForArray("ap_bills", apBills));
          changes.push(...computeChangesForArray("users", users));
          changes.push(...computeChangesForArray("audit_logs", auditLogs));
          changes.push(...computeChangesForArray("monthly_po_lines", monthlyPoLines));
          changes.push(...computeChangesForArray("monthly_ar_lines", monthlyArLines));
          const expectedSnap = JSON.stringify(monthlyArExpectedSnapshots || {});
          if (liveSyncedSnapshotRef.current.monthly_ar_expected !== expectedSnap) {
            changes.push({ entity_type: "monthly_ar_expected", entity_id: 0, payload: monthlyArExpectedSnapshots || {}, is_deleted: false, region: null });
          }
          const settingsPayload = buildSharedSettings();
          const settingsSnap = JSON.stringify(settingsPayload);
          if (liveSyncedSnapshotRef.current.settings !== settingsSnap) {
            changes.push({ entity_type: "settings", entity_id: 0, payload: settingsPayload, is_deleted: false, region: null });
          }
          return changes;
        }

        function buildSeedChanges() {
          const changes = [];
          const pushArr = (entityType, arr) => (arr || []).forEach((r) => {
            if (!r || r.id == null) return;
            changes.push({ entity_type: entityType, entity_id: r.id, payload: stripSyncMeta(r), is_deleted: false, region: r.region || null });
          });
          pushArr("clients", clients);
          pushArr("jobs", jobs);
          pushArr("quotations", quotations);
          pushArr("vendors", vendors);
          pushArr("ar_invoices", arInvoices);
          pushArr("ap_bills", apBills);
          pushArr("users", users);
          pushArr("audit_logs", auditLogs);
          pushArr("monthly_po_lines", monthlyPoLines);
          pushArr("monthly_ar_lines", monthlyArLines);
          changes.push({ entity_type: "monthly_ar_expected", entity_id: 0, payload: monthlyArExpectedSnapshots || {}, is_deleted: false, region: null });
          changes.push({ entity_type: "settings", entity_id: 0, payload: buildSharedSettings(), is_deleted: false, region: null });
          return changes;
        }

        async function pushLiveChangesInBatches(changes) {
          if (!changes.length) return liveServerVersionRef.current;
          const BATCH = 40;
          let version = liveServerVersionRef.current;
          for (let i = 0; i < changes.length; i += BATCH) {
            const res = await fetch("/api/sync/push", {
              method: "POST",
              headers: liveSyncHeaders(),
              body: JSON.stringify({ changes: changes.slice(i, i + BATCH), updated_by: sessionUserId })
            });
            const body = await res.json().catch(() => ({}));
            if (!res.ok) throw new Error(body.detail || res.statusText);
            version = body.server_version || version;
          }
          liveServerVersionRef.current = version;
          return version;
        }

        function mergeArrayByRemote(local, remoteList) {
          const map = new Map((local || []).map((r) => [r.id, r]));
          (remoteList || []).forEach((ch) => {
            if (ch.is_deleted) { map.delete(ch.entity_id); return; }
            const existing = map.get(ch.entity_id);
            const remoteAt = ch.updated_at || "";
            if (!existing) {
              map.set(ch.entity_id, { ...(ch.payload || {}), _syncAt: remoteAt });
              return;
            }
            if (!existing._syncAt) return;
            const localAt = existing._syncAt || "";
            if (remoteAt >= localAt) {
              map.set(ch.entity_id, { ...(ch.payload || {}), _syncAt: remoteAt });
            }
          });
          return Array.from(map.values());
        }

        function applyRemoteLiveChanges(changes) {
          if (!changes || !changes.length) return;
          const byType = {};
          changes.forEach((ch) => {
            if (!byType[ch.entity_type]) byType[ch.entity_type] = [];
            byType[ch.entity_type].push(ch);
          });
          liveApplyingRemoteRef.current = true;
          if (byType.clients) setClients((prev) => ensureRegionOnRecords(mergeArrayByRemote(prev, byType.clients)));
          if (byType.jobs) setJobs((prev) => ensureJobsQuotations(ensureRegionOnRecords(mergeArrayByRemote(prev, byType.jobs))));
          if (byType.quotations) setQuotations((prev) => ensureRegionOnRecords(mergeArrayByRemote(prev, byType.quotations)));
          if (byType.vendors) setVendors((prev) => ensureVendorNumbers(ensureRegionOnRecords(mergeArrayByRemote(prev, byType.vendors))));
          if (byType.ar_invoices) setArInvoices((prev) => ensureRegionOnRecords(mergeArrayByRemote(prev, byType.ar_invoices)));
          if (byType.ap_bills) setApBills((prev) => ensureRegionOnRecords(mergeArrayByRemote(prev, byType.ap_bills)));
          if (byType.users) setUsers((prev) => mergeArrayByRemote(prev, byType.users));
          if (byType.audit_logs) setAuditLogs((prev) => mergeArrayByRemote(prev, byType.audit_logs));
          if (byType.monthly_po_lines) setMonthlyPoLines((prev) => mergeArrayByRemote(prev, byType.monthly_po_lines));
          if (byType.monthly_ar_lines) setMonthlyArLines((prev) => mergeArrayByRemote(prev, byType.monthly_ar_lines));
          if (byType.monthly_ar_expected && byType.monthly_ar_expected.length) {
            const ch = byType.monthly_ar_expected[byType.monthly_ar_expected.length - 1];
            if (!ch.is_deleted) setMonthlyArExpectedSnapshots(ch.payload || {});
          }
          if (byType.settings && byType.settings.length) {
            applySharedSettings(byType.settings[byType.settings.length - 1].payload || {});
          }
          setTimeout(() => { liveApplyingRemoteRef.current = false; markAllSyncedFromState(); }, 0);
        }

        function remoteArrayToChanges(arr) {
          return (arr || []).map((r) => ({
            entity_id: r.id,
            payload: stripSyncMeta(r),
            is_deleted: false,
            updated_at: r.updated_at || r._syncAt || ""
          }));
        }

        function mergeFullWithLocal(local, remoteArr) {
          return mergeArrayByRemote(local || [], remoteArrayToChanges(remoteArr));
        }

        function applyFullLiveEntities(body) {
          const e = body.entities || {};
          const singletons = body.singletons || {};
          liveApplyingRemoteRef.current = true;
          setClients((prev) => ensureRegionOnRecords(mergeFullWithLocal(prev, e.clients)));
          setJobs((prev) => ensureJobsQuotations(ensureRegionOnRecords(mergeFullWithLocal(prev, e.jobs))));
          setQuotations((prev) => ensureRegionOnRecords(mergeFullWithLocal(prev, e.quotations)));
          setVendors((prev) => ensureVendorNumbers(ensureRegionOnRecords(mergeFullWithLocal(prev, e.vendors))));
          setArInvoices((prev) => ensureRegionOnRecords(mergeFullWithLocal(prev, e.ar_invoices)));
          setApBills((prev) => ensureRegionOnRecords(mergeFullWithLocal(prev, e.ap_bills)));
          if (e.users) setUsers((prev) => mergeFullWithLocal(prev, e.users));
          setAuditLogs((prev) => mergeFullWithLocal(prev, e.audit_logs));
          setMonthlyPoLines((prev) => mergeFullWithLocal(prev, e.monthly_po_lines));
          setMonthlyArLines((prev) => mergeFullWithLocal(prev, e.monthly_ar_lines));
          setMonthlyArExpectedSnapshots(singletons.monthly_ar_expected || {});
          if (singletons.settings) applySharedSettings(singletons.settings);
          liveServerVersionRef.current = body.server_version || 0;
          setTimeout(() => { liveApplyingRemoteRef.current = false; markAllSyncedFromState(); }, 0);
        }

        async function flushLiveSyncPush(snapshotOverride) {
          if (!liveSyncEnabled || !canLiveSyncPush() || liveApplyingRemoteRef.current) return false;
          if (!erpSyncKey.trim()) return false;
          const changes = snapshotOverride ? collectChangesFromSnapshot(snapshotOverride) : collectAllLiveChanges();
          if (!changes.length) return true;
          try {
            setLiveSyncStatus("syncing");
            await pushLiveChangesInBatches(changes);
            if (snapshotOverride) markSyncedFromSnapshot(snapshotOverride);
            else markAllSyncedFromState();
            setLiveSyncLastAt(new Date());
            setLiveSyncStatus("ok");
            setImportStatus((prev) => (isImportStatusSticky(prev) ? "" : prev));
            return true;
          } catch (err) {
            setLiveSyncStatus("error");
            console.warn("Live sync push failed", err);
            if (String(err.message || "").includes("401") || String(err.message || "").includes("403")) {
              setImportStatus(t("liveSyncNeedKey"));
            }
            throw err;
          }
        }

        async function pollLiveSync() {
          if (!liveSyncEnabled || liveApplyingRemoteRef.current) return;
          const unpushed = collectAllLiveChanges().length > 0;
          if (unpushed && canLiveSyncPush() && erpSyncKey.trim()) return;
          try {
            const res = await fetch("/api/sync/changes?since=" + liveServerVersionRef.current);
            const body = await res.json().catch(() => ({}));
            if (!res.ok) throw new Error(body.detail || res.statusText);
            if (body.server_version > liveServerVersionRef.current && body.changes && body.changes.length) {
              applyRemoteLiveChanges(body.changes);
              liveServerVersionRef.current = body.server_version;
              setLiveSyncLastAt(new Date());
            }
            setLiveSyncStatus("ok");
          } catch (err) {
            setLiveSyncStatus("error");
            console.warn("Live sync poll failed", err);
          }
        }

        async function bootstrapLiveSync() {
          try {
            setLiveSyncStatus("syncing");
            const statusRes = await fetch("/api/sync/status");
            const status = await statusRes.json().catch(() => ({}));
            if (!statusRes.ok) throw new Error(status.detail || statusRes.statusText);
            liveServerVersionRef.current = status.server_version || 0;
            if ((status.entity_count || 0) === 0) {
              const hasLocal = !cloudOnlyMode && (clients.length + jobs.length + quotations.length + vendors.length + arInvoices.length + apBills.length);
              if (hasLocal && canLiveSyncPush() && erpSyncKey.trim()) {
                setLiveSyncReady(true);
                await flushLiveSyncPush();
                setLiveSyncLastAt(new Date());
                setLiveSyncStatus("ok");
                return true;
              }
              if (!hasLocal) {
                applyFullLiveEntities({
                  server_version: status.server_version || 0,
                  entities: {},
                  singletons: {}
                });
              }
            } else {
              const fullRes = await fetch("/api/sync/full");
              const full = await fullRes.json().catch(() => ({}));
              if (!fullRes.ok) throw new Error(full.detail || fullRes.statusText);
              applyFullLiveEntities(full);
            }
            setLiveSyncLastAt(new Date());
            setLiveSyncStatus("ok");
            setLiveSyncReady(true);
            return true;
          } catch (err) {
            setLiveSyncStatus("error");
            console.warn("Live sync bootstrap failed", err);
            return false;
          }
        }

        function downloadErpBackup() {
          if (!can("settings", "view")) { alert(t("noPermission")); return; }
          const payload = buildErpBackupPayload();
          const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
          const a = document.createElement("a");
          a.href = URL.createObjectURL(blob);
          a.download = "erp_backup_" + new Date().toISOString().slice(0, 10) + ".json";
          a.click();
          URL.revokeObjectURL(a.href);
          logAudit("settings", "export", "backup", "Downloaded ERP backup");
          alert(t("backupDone"));
        }

        function handleRestoreBackup(file) {
          if (!file) return;
          if (!isRoot()) { alert(t("rootOnly")); return; }
          const reader = new FileReader();
          reader.onerror = () => alert(t("importFileError"));
          reader.onload = (e) => {
            try {
              const payload = JSON.parse(e.target.result);
              const d = payload && payload.data;
              if (!d || !Array.isArray(d.clients) || !Array.isArray(d.jobs)) { alert(t("restoreInvalid")); return; }
              if (!window.confirm(t("restoreConfirm"))) return;
              if (!applyErpBackupData(d)) { alert(t("restoreInvalid")); return; }
              logAudit("settings", "import", "backup", "Restored ERP backup from file");
              alert(t("restoreDone"));
              window.location.reload();
            } catch (err) {
              alert(t("restoreInvalid") + " " + (err.message || ""));
            }
          };
          reader.readAsText(file);
        }

        function applyImportPreview() {
          if (!importPreview) return;
          const { module, preview, fileName } = importPreview;
          const valid = preview.filter((p) => p.action !== "error");
          if (!valid.length) { alert(t("importNoRows")); return; }

          resetErpIdAllocator();
          const idSources = () => [clients, jobs, quotations, vendors, arInvoices, apBills, auditLogs, monthlyPoLines, monthlyArLines];

          let created = 0;
          let updated = 0;
          let nextClients = clients;
          let nextJobs = jobs;
          let nextQuotations = quotations;
          let nextVendors = vendors;
          let nextAr = arInvoices;
          let nextAp = apBills;
          let nextQuotationsFromJobs = quotations;
          if (module === "clients") {
            nextClients = [...clients];
            valid.forEach((p) => {
              const payload = { ...p.data };
              const existing = nextClients.find((c) => c.customer_no.toLowerCase() === p.key.toLowerCase());
              if (existing) {
                const oldCompany = existing.company;
                nextClients = nextClients.map((c) => (c.id === existing.id ? { ...c, ...payload } : c));
                if (oldCompany !== payload.company) syncCompanyName(oldCompany, payload.company);
                updated++;
              } else {
                nextClients = [{ id: nextErpRecordId(nextClients, ...idSources()), region: regionForNewRecord(), ...payload }, ...nextClients];
                created++;
              }
            });
            setClients(nextClients);
          } else if (module === "job") {
            nextJobs = [...jobs];
            valid.forEach((p) => {
              const client = clients.find((c) => c.company.toLowerCase() === String(p.data.company || "").trim().toLowerCase());
              const payload = normalizeJobQuotations({
                job_no: p.data.job_no,
                company: client ? client.company : p.data.company,
                client_id: client ? client.id : null,
                job_type: p.data.job_type || "Project",
                quotation_no: p.data.quotation_no || "",
                linked_quotation_nos: p.data.quotation_no ? String(p.data.quotation_no).split(",").map((s) => s.trim()).filter(Boolean) : [],
                amount: Number(p.data.amount || 0),
                customer_po_amount: Number(p.data.customer_po_amount || p.data.amount || 0),
                currency: p.data.currency || "USD",
                customer_po: p.data.customer_po || "",
                status: p.data.status || "Open",
                start_date: p.data.start_date || "",
                description: p.data.description || "",
                remarks: p.data.remarks || ""
              });
              const normalizedPayload = normalizeJobPoFields(payload);
              const existing = nextJobs.find((j) => j.job_no.toLowerCase() === p.key.toLowerCase());
              if (existing) {
                const oldNo = existing.job_no;
                const merged = p.data.quotation_no ? addQuotationToJob(existing, p.data.quotation_no) : existing;
                const updatePayload = { ...merged, ...normalizedPayload, linked_quotation_nos: merged.linked_quotation_nos, quotation_no: merged.quotation_no };
                nextJobs = nextJobs.map((j) => (j.id === existing.id ? { ...j, ...updatePayload } : j));
                if (oldNo !== payload.job_no) {
                  setArInvoices((prev) => prev.map((r) => (r.job_no === oldNo ? { ...r, job_no: payload.job_no, job_id: existing.id } : r)));
                }
                syncArWithJobStatus(existing.id, payload.job_no, payload.status);
                updated++;
              } else {
                const newId = nextErpRecordId(nextJobs, ...idSources());
                nextJobs = [{ id: newId, region: regionForNewRecord(), ...normalizedPayload }, ...nextJobs];
                created++;
              }
            });
            let nextQuotationsFromJobs = [...quotations];
            valid.forEach((p) => {
              const jobNo = String(p.data.job_no || "").trim();
              if (!jobNo) return;
              const job = nextJobs.find((j) => j.job_no.toLowerCase() === jobNo.toLowerCase());
              if (!job) return;
              jobQuotationNos(job).forEach((qNo) => {
                nextQuotationsFromJobs = nextQuotationsFromJobs.map((q) => {
                  if (String(q.quotation_no || "").trim().toLowerCase() !== qNo.toLowerCase()) return q;
                  return addJobToQuotation(q, job.id, job.job_no);
                });
              });
            });
            setJobs(nextJobs);
            setQuotations(nextQuotationsFromJobs);
          } else if (module === "ar") {
            nextAr = [...arInvoices];
            valid.forEach((p) => {
              const job = jobs.find((j) => j.job_no.toLowerCase() === String(p.data.job_no || "").toLowerCase());
              const invAmt = Number(p.data.invoice_amt || 0);
              const invCur = p.data.invoice_currency || "USD";
              const base = convertCurrency(invAmt, invCur, REGION_CURRENCY[job?.region || regionForNewRecord()] || "HKD");
              const payload = {
                job_id: job ? job.id : null,
                job_no: job ? job.job_no : p.data.job_no,
                customer: p.data.customer || (job ? job.company : ""),
                customer_po: p.data.customer_po || (job ? job.customer_po : ""),
                invoice_no: p.data.invoice_no,
                invoice_currency: invCur,
                invoice_amt: invAmt,
                exchange_rate_locked: fxUsdMap[invCur] ? (fxUsdMap[REGION_CURRENCY[job?.region || "HK"] || "HKD"] / fxUsdMap[invCur]) : 1,
                base_amount_locked: base,
                base_amount: base,
                lock_timestamp: nowIso(),
                rate_source: "IMPORT",
                manual_override: false,
                override_reason: "",
                invoice_date: p.data.invoice_date || "",
                due_date: p.data.due_date || "",
                payment_received_date: p.data.payment_received_date || "",
                payment_status: deriveArPaymentStatus({
                  payment_status: p.data.payment_status,
                  payment_received_date: p.data.payment_received_date || "",
                  due_date: p.data.due_date || ""
                }),
                job_completed: job ? isJobCompleted(job.status) : false
              };
              const existing = nextAr.find((r) => r.invoice_no.toLowerCase() === p.key.toLowerCase());
              if (existing) {
                nextAr = nextAr.map((r) => (r.id === existing.id ? { ...r, ...payload } : r));
                updated++;
              } else {
                nextAr = [{ id: nextErpRecordId(nextAr, ...idSources()), region: job?.region || regionForNewRecord(), ...payload }, ...nextAr];
                created++;
              }
            });
            setArInvoices(nextAr);
          } else if (module === "ap") {
            nextAp = [...apBills];
            valid.forEach((p) => {
              const job = jobs.find((j) => j.job_no.toLowerCase() === String(p.data.job_no || "").toLowerCase());
              const amt = Number(p.data.amount || 0);
              const cur = p.data.currency || "USD";
              const base = convertCurrency(amt, cur, REGION_CURRENCY[job?.region || regionForNewRecord()] || "HKD");
              const payload = {
                job_id: job ? job.id : null,
                job_no: job ? job.job_no : p.data.job_no,
                payee_type: p.data.payee_type || "Vendor",
                si_id: null,
                company_name: p.data.company_name,
                invoice_no: p.data.invoice_no,
                invoice_date: p.data.invoice_date || "",
                invoice_received_date: "",
                due_date: p.data.due_date || "",
                currency: cur,
                amount: amt,
                exchange_rate_locked: 1,
                base_amount_locked: base,
                base_amount: base,
                lock_timestamp: nowIso(),
                rate_source: "IMPORT",
                manual_override: false,
                override_reason: "",
                pay_date: p.data.pay_date || "",
                payment_status: deriveApPaymentStatus({
                  pay_date: p.data.pay_date || "",
                  due_date: p.data.due_date || "",
                  payment_status: p.data.payment_status
                }),
                remarks: p.data.remarks || "",
                bank: "", charge: "OUR", swift_code: "", payment_advice_email: ""
              };
              const existing = nextAp.find((b) => b.invoice_no.toLowerCase() === p.key.toLowerCase());
              if (existing) {
                nextAp = nextAp.map((b) => (b.id === existing.id ? { ...b, ...payload } : b));
                updated++;
              } else {
                nextAp = [{ id: nextErpRecordId(nextAp, ...idSources()), region: job?.region || regionForNewRecord(), ...payload }, ...nextAp];
                created++;
              }
            });
            setApBills(nextAp);
          } else if (module === "vendors") {
            nextVendors = [...vendors];
            valid.forEach((p) => {
              const payload = { vendor_no: p.data.vendor_no, name: p.data.name, email: p.data.email || "", phone: p.data.phone || "", address: p.data.address || "", bank: p.data.bank || "", swift_code: p.data.swift_code || "", charge: p.data.charge || "OUR", contact: p.data.contact || "", finance_contact: p.data.finance_contact || "", finance_email: p.data.finance_email || "" };
              const existing = nextVendors.find((v) => v.vendor_no.toLowerCase() === p.key.toLowerCase());
              if (existing) { nextVendors = nextVendors.map((v) => (v.id === existing.id ? { ...v, ...payload } : v)); updated++; }
              else { nextVendors = [{ id: nextErpRecordId(nextVendors, ...idSources()), region: regionForNewRecord(), ...payload }, ...nextVendors]; created++; }
            });
            setVendors(nextVendors);
          } else if (module === "quotation") {
            nextQuotations = [...quotations];
            valid.forEach((p) => {
              const client = clients.find((c) => c.company.toLowerCase() === String(p.data.company || "").trim().toLowerCase());
              const job = p.data.job_no ? jobs.find((j) => j.job_no.toLowerCase() === String(p.data.job_no).toLowerCase()) : null;
              const payload = {
                quotation_no: p.data.quotation_no,
                company: client ? client.company : p.data.company,
                client_id: client ? client.id : null,
                quotation_date: p.data.quotation_date || "",
                valid_until: p.data.valid_until || "",
                amount: Number(p.data.amount || 0),
                currency: p.data.currency || "USD",
                status: p.data.status || "Draft",
                description: p.data.description || "",
                job_no: job ? job.job_no : (p.data.job_no || ""),
                job_id: job ? job.id : null
              };
              const existing = nextQuotations.find((q) => q.quotation_no.toLowerCase() === p.key.toLowerCase());
              if (existing) { nextQuotations = nextQuotations.map((q) => (q.id === existing.id ? { ...q, ...payload } : q)); updated++; }
              else { nextQuotations = [{ id: nextErpRecordId(nextQuotations, ...idSources()), region: regionForNewRecord(), ...payload }, ...nextQuotations]; created++; }
            });
            setQuotations(nextQuotations);
          }

          const importSnap = {
            clients: nextClients,
            jobs: nextJobs,
            quotations: module === "job" ? nextQuotationsFromJobs : (module === "quotation" ? nextQuotations : quotations),
            vendors: nextVendors,
            arInvoices: nextAr,
            apBills: nextAp
          };
          persistErpLists(importSnap);

          logAudit(module, "import", fileName, `Imported ${created} new, ${updated} updated, ${preview.length - valid.length} errors`);
          setImportPreview(null);
          const regionNote = t("importRegionHint").replace("{region}", activeRegion === "ALL" ? t("allRegions") : regionLabel(activeRegion, lang));
          const summary = t("importSummary") + ": " + created + " " + t("importRowCreate") + ", " + updated + " " + t("importRowUpdate") + "\n" + regionNote;
          (async () => {
            let syncOk = true;
            if (liveSyncEnabled && canLiveSyncPush() && erpSyncKey.trim()) {
              try {
                await flushLiveSyncPush(importSnap);
              } catch (err) {
                syncOk = false;
                console.warn("Import cloud sync failed", err);
              }
            }
            if (syncOk && liveSyncEnabled && erpSyncKey.trim()) {
              setImportStatus(summary.replace("\n", " · ") + " · " + t("importSyncOk"));
              alert(summary + "\n" + t("importSyncOk"));
            } else if (!syncOk) {
              setImportStatus(t("importSyncFail"));
              alert(summary + "\n\n" + t("importSyncFail"));
            } else {
              setImportStatus(summary.replace("\n", " · "));
              alert(summary);
            }
          })();
        }

        function exportOngoingExcel() {
          if (!guardPermission("ongoing", "export")) return;
          if (typeof XLSX === "undefined") { alert("Excel library not loaded"); return; }
          const regionName = activeRegion === "ALL" ? t("allRegions") : regionLabel(activeRegion, lang);
          const rows = ongoingJobsList.map((j) => ({
            job_no: j.job_no,
            nature: j.nature,
            customer: j.company,
            project: j.description || "",
            customer_po: j.customer_po || "",
            start_date: j.start_date || "",
            currency: j.arCurrency || "",
            po_amt: j.poAmtLabel || j.poAmt,
            billed_amt: j.billedAmt,
            po_balance: j.poBalance,
            balance_usd: j.balanceUsd,
            remarks: j.remarks || ""
          }));
          const schema = {
            sheetName: "Ongoing",
            filePrefix: "ongoing_" + regionName.replace(/\s+/g, "_"),
            columns: [
              { header: "Job No.", field: "job_no" }, { header: "Job Nature", field: "nature" },
              { header: "Customer", field: "customer" }, { header: "Project Name", field: "project" },
              { header: "Customer PO", field: "customer_po" }, { header: "Start Date", field: "start_date" },
              { header: "Currency", field: "currency" }, { header: "PO Amt", field: "po_amt" },
              { header: "Billed Amt", field: "billed_amt" }, { header: "PO Balance", field: "po_balance" },
              { header: "Balance USD", field: "balance_usd" }, { header: "Remarks", field: "remarks" }
            ]
          };
          downloadWorkbook(buildWorkbook(schema, rows, false), schema.filePrefix + ".xlsx");
          logAudit("ongoing", "export", schema.filePrefix, `Exported ongoing list Excel (${rows.length} rows)`);
        }

        function exportOngoingPdf() {
          if (!guardPermission("ongoing", "export")) return;
          const regionName = activeRegion === "ALL" ? t("allRegions") : regionLabel(activeRegion, lang);
          const title = t("ongoingListTitle") + " - " + regionName;
          ensurePdfLibs().then(function (ok) {
            if (!ok) { alert("PDF library not loaded"); return; }
            exportOngoingPdfReport({ rows: ongoingJobsList, totals: ongoingTotals, title, t, moneyFn: money });
            logAudit("ongoing", "export", title, `Exported ongoing list PDF (${ongoingJobsList.length} rows)`);
          });
        }

        function ReportExportButtons({ onExcel, onPdf }) {
          return (
            <div className="flex items-center gap-2 shrink-0">
              <button type="button" onClick={onExcel} className="h-9 px-4 text-xs rounded-lg border border-emerald-400 bg-emerald-100 text-emerald-900 hover:bg-emerald-200 font-semibold whitespace-nowrap shadow-sm">{t("exportExcel")}</button>
              <button type="button" onClick={onPdf} className="h-9 px-4 text-xs rounded-lg border border-red-400 bg-red-100 text-red-900 hover:bg-red-200 font-semibold whitespace-nowrap shadow-sm">{t("exportPdf")}</button>
            </div>
          );
        }

        function ListImportExportBar({ module }) {
          const schema = EXCEL_SCHEMAS[module];
          const exportOnly = module === "monthly_report";
          if (!schema && !exportOnly) return null;
          const showExport = can(module, "export");
          const showImport = can(module, "import") && !!EXCEL_SCHEMAS[module];
          if (!showExport && !showImport) return null;
          return (
            <div className="flex items-center gap-2">
              {showExport && (
                <>
                  <button type="button" onClick={() => exportModuleExcel(module, "template")} className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 bg-white hover:bg-slate-50">{t("exportTemplate")}</button>
                  {can(module, "view") && <button type="button" onClick={() => exportModuleExcel(module, "data")} className="px-2.5 py-1.5 text-xs rounded-lg border border-slate-300 bg-white hover:bg-slate-50">{t("exportData")}</button>}
                </>
              )}
              {showImport && (
                <button type="button" onClick={() => openImportPicker(module)} className="px-2.5 py-1.5 text-xs rounded-lg border border-emerald-300 bg-emerald-50 text-emerald-800 hover:bg-emerald-100 cursor-pointer" title={t("importModeHint")}>{t("importExcel")}</button>
              )}
            </div>
          );
        }

        function SortableTh({ module, field, label, className = "p-3", align = "left" }) {
          const cur = tableSort[module] || LIST_DEFAULT_SORT[module] || { field: null, dir: "asc" };
          const active = cur.field === field;
          const arrow = active ? (cur.dir === "asc" ? " ▲" : " ▼") : "";
          return (
            <th
              className={className + " " + (align === "right" ? "text-right" : "text-left") + (active ? " text-blue-700 font-semibold" : "")}
              onClick={() => toggleListSort(module, field)}
              title={t("sortClickHint")}
              style={{ cursor: "pointer", userSelect: "none" }}
            >
              {label}{arrow}
            </th>
          );
        }

        function DashGradientCard({ title, value, hint, badge, subValue, gradient, icon }) {
          return (
            <article className="relative overflow-hidden rounded-xl text-white p-5 min-h-[8.5rem] shadow-xl" style={{ background: gradient }}>
              <div className="absolute -right-8 -bottom-10 w-36 h-36 rounded-full bg-white/10 blur-[1px] pointer-events-none" />
              <div className="absolute right-6 top-3 w-20 h-20 rounded-full bg-white/[0.07] pointer-events-none" />
              <div className="absolute inset-0 bg-gradient-to-t from-black/10 to-transparent pointer-events-none" />
              <p className="text-sm font-semibold text-white/95 relative z-10">{title}</p>
              <div className="flex items-end justify-between gap-3 mt-4 relative z-10">
                <div className="text-white/30 shrink-0">{icon}</div>
                <div className="text-right min-w-0">
                  <p className="text-2xl xl:text-3xl font-bold tabular-nums tracking-tight leading-none">{value}</p>
                  {subValue && <p className="text-base font-semibold text-white/95 tabular-nums mt-1.5">{subValue}</p>}
                  {hint && <p className="text-[11px] text-white/80 mt-1.5 leading-snug">{hint}</p>}
                </div>
              </div>
              {badge && <p className="text-[11px] mt-3 bg-white/20 backdrop-blur-sm rounded-md px-2 py-1 inline-block relative z-10">{badge}</p>}
            </article>
          );
        }

        const DASH_GRADIENTS = {
          blue: "linear-gradient(135deg, #0f2b6e 0%, #1d4ed8 45%, #38bdf8 100%)",
          purple: "linear-gradient(135deg, #5b21b6 0%, #7c3aed 40%, #d946ef 100%)",
          redOrange: "linear-gradient(135deg, #b91c1c 0%, #ea580c 55%, #fbbf24 100%)",
          indigo: "linear-gradient(135deg, #312e81 0%, #4f46e5 50%, #818cf8 100%)"
        };

        function DashCardIcon({ children }) {
          return (
            <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth={1.25} strokeLinecap="round" strokeLinejoin="round" className="w-14 h-14">
              {children}
            </svg>
          );
        }

        function DashDueRow({ item, type, nameKey, noKey }) {
          const name = item[nameKey];
          const no = item[noKey];
          const daysUntil = item.days_until;
          const overdueDays = item.overdue_days;
          return (
            <button
              type="button"
              onClick={() => setDetailPanel({ type, id: item.id })}
              className="w-full flex items-center gap-3 px-4 py-3 hover:bg-slate-50 text-left border-b border-slate-100 last:border-0"
            >
              <div className="min-w-0 flex-1">
                <p className="text-sm font-medium text-slate-800 truncate">{no}</p>
                <p className="text-xs text-slate-500 truncate">{name}</p>
              </div>
              <div className="text-right shrink-0">
                <p className="text-sm font-semibold tabular-nums">{money(item.display_base_amount)}</p>
                <span className={"inline-block mt-0.5 text-[10px] px-1.5 py-0.5 rounded-full font-medium " + dueBadgeClass(daysUntil, overdueDays)}>
                  {dueBadgeLabel(daysUntil, overdueDays, t)}
                </span>
              </div>
            </button>
          );
        }

        function SidebarItem({ id, label, module, collapsed }) {
          if (module && !can(module, "view")) return null;
          const active = page === id;
          return (
            <button
              type="button"
              onClick={() => setPage(id)}
              title={label}
              className={
                "group w-full rounded-lg text-xs flex items-center transition-all duration-300 " +
                (collapsed ? "justify-center py-2.5" : "justify-start gap-2.5 px-2.5 py-2") +
                " " + (active ? "bg-blue-600 text-white font-medium" : "text-slate-100 hover:bg-slate-800")
              }
            >
              <SidebarNavIcon id={id} className="w-5 h-5 shrink-0" />
              <span className={collapsed ? "hidden" : "truncate whitespace-nowrap"}>{label}</span>
            </button>
          );
        }

        function renderDetailPanel() {
          if (!detailPanel) return null;
          const { type, id } = detailPanel;

          if (type === "client") {
            const c = clients.find((x) => x.id === id);
            if (!c) return null;
            const relatedJobs = getJobsByClient(c.company);
            const relatedAr = getArByClient(c.company);
            const pos = [...new Set([...relatedJobs.map((j) => j.customer_po), ...relatedAr.map((r) => r.customer_po)].filter(Boolean))];
            return (
              <Modal title={t("clientDetail") + " - " + c.company} onClose={() => setDetailPanel(null)} wide>
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4 text-sm">
                  <div><p className="text-slate-500">{t("colCustomerNo")}</p><p className="font-medium">{c.customer_no}</p></div>
                  <div><p className="text-slate-500">{t("colCompany")}</p><p className="font-medium">{c.company}</p></div>
                  <div><p className="text-slate-500">Invoice Title</p><p>{c.invoice_title || "-"}</p></div>
                  {c.is_bu && <div><p className="text-slate-500">{t("colBuNo")}</p><p>{c.bu_no || "-"}</p></div>}
                  <div><p className="text-slate-500">{t("colPrimaryContact")}</p><p>{c.primary_contact}</p></div>
                  <div><p className="text-slate-500">E-mail</p><p>{c.email}</p></div>
                  <div><p className="text-slate-500">{t("colPaymentTerms")}</p><p>{c.payment_terms}</p></div>
                  <div><p className="text-slate-500">Address</p><p>{c.address}</p></div>
                </div>
                <div className="mt-6 space-y-4">
                  <div>
                    <h4 className="font-semibold mb-2">{t("relatedJobs")} ({relatedJobs.length})</h4>
                    {relatedJobs.length ? (
                      <table className="w-full text-sm border rounded-lg overflow-hidden">
                        <thead className="bg-slate-50"><tr><th className="p-2 text-left">{t("colJobNo")}</th><th className="p-2 text-left">{t("type")}</th><th className="p-2 text-left">{t("colCustomerPo")}</th><th className="p-2 text-left">{t("status")}</th><th className="p-2 text-left"></th></tr></thead>
                        <tbody>{relatedJobs.map((j) => (
                          <tr key={j.id} className="border-t">
                            <td className="p-2"><LinkBtn onClick={() => navigateTo("job", { client: c.company, highlightId: j.id })}>{j.job_no}</LinkBtn></td>
                            <td className="p-2">{j.job_type}</td>
                            <td className="p-2">{j.customer_po ? <LinkBtn onClick={() => navigateToJobByPo(c.company, j.customer_po)}>{j.customer_po}</LinkBtn> : "-"}</td>
                            <td className="p-2"><span className={"px-2 py-0.5 rounded text-xs " + jobStatusClass(j.status)}>{j.status}</span></td>
                            <td className="p-2"><button type="button" className="text-xs text-blue-600 hover:underline" onClick={() => navigateTo("job", { client: c.company, highlightId: j.id })}>{t("goToJob")}</button></td>
                          </tr>
                        ))}</tbody>
                      </table>
                    ) : <p className="text-slate-500">{t("noJobs")}</p>}
                  </div>
                  <div>
                    <h4 className="font-semibold mb-2">{t("relatedInvoices")} ({relatedAr.length})</h4>
                    {relatedAr.length ? (
                      <table className="w-full text-sm border rounded-lg overflow-hidden">
                        <thead className="bg-slate-50"><tr><th className="p-2 text-left">{t("colInvoiceNo")}</th><th className="p-2 text-left">{t("colJobNo")}</th><th className="p-2 text-right">{t("colAmt")}</th><th className="p-2 text-left">{t("paymentStatus")}</th><th className="p-2 text-left"></th></tr></thead>
                        <tbody>{relatedAr.map((r) => (
                          <tr key={r.id} className="border-t">
                            <td className="p-2"><LinkBtn onClick={() => navigateTo("ar", { client: c.company, highlightId: r.id })}>{r.invoice_no}</LinkBtn></td>
                            <td className="p-2">{r.job_no ? <LinkBtn onClick={() => { const job = jobs.find((j) => j.job_no === r.job_no); navigateTo("job", { client: c.company, highlightId: job ? job.id : null }); }}>{r.job_no}</LinkBtn> : "-"}</td>
                            <td className="p-2 text-right">{money(r.invoice_amt)} {r.invoice_currency}</td>
                            <td className="p-2"><span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(r.payment_status)}>{paymentStatusLabel(r.payment_status, t)}</span></td>
                            <td className="p-2"><button type="button" className="text-xs text-blue-600 hover:underline" onClick={() => navigateTo("ar", { client: c.company, highlightId: r.id })}>{t("goToAr")}</button></td>
                          </tr>
                        ))}</tbody>
                      </table>
                    ) : <p className="text-slate-500">{t("noInvoices")}</p>}
                  </div>
                  <div>
                    <h4 className="font-semibold mb-2">{t("relatedPo")} ({pos.length})</h4>
                    {pos.length ? (
                      <ul className="text-sm space-y-1">{pos.map((po) => (
                        <li key={po}><LinkBtn onClick={() => navigateToArByPo(c.company, po)}>{po}</LinkBtn> <span className="text-slate-400 text-xs">→ {t("goToAr")}</span></li>
                      ))}</ul>
                    ) : <p className="text-slate-500">{t("noPo")}</p>}
                  </div>
                </div>
                <DetailActions
                  editLabel={t("edit")}
                  deleteLabel={t("delete")}
                  canEdit={can("clients", "edit")}
                  canDelete={can("clients", "delete")}
                  onEdit={() => { setDetailPanel(null); setClientModal({ mode: "edit", id: c.id, data: { ...c } }); }}
                  onDelete={() => { deleteClient(c.id); setDetailPanel(null); }}
                />
              </Modal>
            );
          }

          if (type === "quotation") {
            const q = quotations.find((x) => x.id === id);
            if (!q) return null;
            return (
              <Modal title={t("quotationDetail") + " - " + q.quotation_no} onClose={() => setDetailPanel(null)} wide>
                <div className="grid grid-cols-2 gap-4 text-sm">
                  <div><p className="text-slate-500">{t("clientGroup")}</p><p className="font-medium">{q.company}</p></div>
                  <div><p className="text-slate-500">{t("quotationStatus")}</p><p>{q.status}</p></div>
                  <div><p className="text-slate-500">{t("quotationDate")}</p><p>{q.quotation_date || "-"}</p></div>
                  <div><p className="text-slate-500">{t("validUntil")}</p><p>{q.valid_until || "-"}</p></div>
                  <div><p className="text-slate-500">{t("amount")}</p><p>{money(q.amount)} {q.currency}</p></div>
                  <div><p className="text-slate-500">{t("colJobNo")}</p><p>{quotationJobNos(q).length ? quotationJobNos(q).map((jno, i) => (
                    <span key={jno}>{i > 0 ? ", " : ""}<LinkBtn onClick={() => { const j = jobs.find((x) => x.job_no === jno); if (j) setDetailPanel({ type: "job", id: j.id }); }}>{jno}</LinkBtn></span>
                  )) : "-"}</p></div>
                  <div className="col-span-2"><p className="text-slate-500">{t("description")}</p><p>{q.description || "-"}</p></div>
                </div>
                <DetailActions
                  editLabel={t("edit")}
                  deleteLabel={t("delete")}
                  canEdit={can("quotation", "edit")}
                  canDelete={can("quotation", "delete")}
                  onEdit={() => { setDetailPanel(null); setQuotationModal({ mode: "edit", id: q.id, data: { ...q, client_id: String(q.client_id || ""), amount: String(q.amount ?? ""), job_no: q.job_no || "" } }); }}
                  onDelete={() => { deleteQuotation(q.id); setDetailPanel(null); }}
                />
              </Modal>
            );
          }

          if (type === "job") {
            const j = normalizeJobPoFields(jobs.find((x) => x.id === id) || null);
            if (!j) return null;
            const relatedAr = getArByJob(j.job_no);
            const relatedAp = getApByJob(j.job_no);
            const client = clients.find((c) => c.id === j.client_id);
            const linkedQuotations = findLinkedQuotationsForJob(j, quotations);
            const jobCurrency = j.currency || "USD";
            const quotationTotal = sumLinkedQuotationAmounts(j, quotations, jobCurrency);
            const poTotal = sumJobPoAmount(j);
            const invoiceTotal = sumJobArInvoiceAmount(j.job_no, arInvoices, regionListCurrency);
            return (
              <Modal title={t("jobDetail") + " - " + j.job_no} onClose={() => setDetailPanel(null)} wide>
                <p className="text-xs text-slate-500 mb-4">{t("jobFlowHint")}</p>
                <div className="grid grid-cols-2 gap-6 text-sm mb-6">
                  <div className="space-y-4">
                    <div><p className="text-slate-500">Client</p><p className="font-medium">{j.company}</p></div>
                    <div><p className="text-slate-500">{t("colCustomerPo")}</p><p>{j.customer_po ? <LinkBtn onClick={() => navigateToArByPo(j.company, j.customer_po)}>{j.customer_po}</LinkBtn> : "-"}</p></div>
                    <div><p className="text-slate-500">{t("quotationNo")}</p><p>{linkedQuotations.length ? linkedQuotations.map((q, i) => (
                      <span key={q.id}>{i > 0 ? ", " : ""}<LinkBtn onClick={() => setDetailPanel({ type: "quotation", id: q.id })}>{q.quotation_no}</LinkBtn></span>
                    )) : (j.quotation_no || "-")}</p></div>
                    <div><p className="text-slate-500">{t("colInvoiceNo")}</p><p>{relatedAr.length ? relatedAr.map((r, i) => (
                      <span key={r.id}>{i > 0 ? ", " : ""}<LinkBtn onClick={() => navigateTo("ar", { client: j.company, highlightId: r.id })}>{r.invoice_no}</LinkBtn></span>
                    )) : "-"}</p></div>
                    <div><p className="text-slate-500">{t("status")}</p><p><span className={"inline-block px-2 py-0.5 rounded text-xs " + jobStatusClass(j.status)}>{j.status}</span></p></div>
                  </div>
                  <div className="space-y-4">
                    <div><p className="text-slate-500">{t("type")}</p><p>{j.job_type}</p></div>
                    <div><p className="text-slate-500">{t("customerPoAmountTotal")}</p><p>{poTotal ? money(poTotal) + " " + jobCurrency : "-"}</p></div>
                    <div><p className="text-slate-500">{t("quotationAmountTotal")}</p><p>{quotationTotal ? money(quotationTotal) + " " + jobCurrency : "-"}</p></div>
                    <div><p className="text-slate-500">{t("invoiceAmountTotal")}</p><p className="font-medium">{invoiceTotal ? money(invoiceTotal) + " " + regionListCurrency : "-"}</p></div>
                  </div>
                </div>
                <div className="grid grid-cols-2 gap-4 text-sm mb-6">
                  <div className="col-span-2">
                    <p className="text-slate-500 mb-2">{t("jobPoLinesTitle")}</p>
                    {j.po_lines && j.po_lines.length ? (
                      <table className="w-full text-sm border rounded-lg overflow-hidden">
                        <thead className="bg-slate-50"><tr><th className="p-2 text-left">{t("jobPoColNo")}</th><th className="p-2 text-right">{t("jobPoColAmount")}</th><th className="p-2 text-left">{t("jobPoColDate")}</th><th className="p-2 text-left">{t("jobPoColRemarks")}</th></tr></thead>
                        <tbody>{j.po_lines.map((pl) => (
                          <tr key={pl.id} className="border-t"><td className="p-2">{pl.po_no || "-"}</td><td className="p-2 text-right tabular-nums">{money(pl.amount)} {j.currency || ""}</td><td className="p-2">{pl.received_date || "-"}</td><td className="p-2">{pl.remarks || "-"}</td></tr>
                        ))}</tbody>
                      </table>
                    ) : <p>-</p>}
                  </div>
                  <div><p className="text-slate-500">{t("colStartDate")}</p><p>{j.start_date || "-"}</p></div>
                  <div><p className="text-slate-500">Description</p><p>{j.description || "-"}</p></div>
                  <div className="col-span-2"><p className="text-slate-500">{t("ongoingColRemarks")}</p><p>{j.remarks || "-"}</p></div>
                </div>
                {client && (
                  <div className="mb-4 p-3 bg-slate-50 rounded-lg text-sm">
                    <p className="font-medium">{t("ownedClient")}</p>
                    <p><LinkBtn onClick={() => { setPage("clients"); setDetailPanel({ type: "client", id: client.id }); }}>{client.customer_no} · {client.company}</LinkBtn> · {client.primary_contact}</p>
                  </div>
                )}
                <h4 className="font-semibold mb-2">{t("relatedInvoices")} ({relatedAr.length})</h4>
                {relatedAr.length ? (
                  <table className="w-full text-sm border rounded-lg overflow-hidden mb-6">
                    <thead className="bg-slate-50"><tr><th className="p-2 text-left">{t("colInvoiceNo")}</th><th className="p-2 text-left">{t("paymentStage")}</th><th className="p-2 text-right">{amtInLabel}</th><th className="p-2 text-left">{t("colDueDate")}</th><th className="p-2 text-left">{t("paymentStatus")}</th></tr></thead>
                    <tbody>{relatedAr.map((r) => (
                      <tr key={r.id} className="border-t">
                        <td className="p-2"><LinkBtn onClick={() => navigateTo("ar", { client: j.company, highlightId: r.id })}>{r.invoice_no}</LinkBtn></td>
                        <td className="p-2">{paymentStageLabel(r.payment_stage, t)}</td>
                        <td className="p-2 text-right">{money(convertCurrency(r.invoice_amt, r.invoice_currency, regionListCurrency))}</td>
                        <td className="p-2">{r.due_date}</td>
                        <td className="p-2"><span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(deriveArPaymentStatus(r))}>{paymentStatusLabel(deriveArPaymentStatus(r), t)}</span></td>
                      </tr>
                    ))}</tbody>
                  </table>
                ) : <p className="text-slate-500 mb-6">{t("noInvoices")}</p>}
                <h4 className="font-semibold mb-2">{t("relatedAp")} ({relatedAp.length})</h4>
                {relatedAp.length ? (
                  <table className="w-full text-sm border rounded-lg overflow-hidden">
                    <thead className="bg-slate-50"><tr><th className="p-2 text-left">{t("colInvoiceNo")}</th><th className="p-2 text-left">{t("payeeType")}</th><th className="p-2 text-left">{t("colPayee")}</th><th className="p-2 text-right">{t("amount")}</th><th className="p-2 text-left">{t("paymentStatus")}</th><th className="p-2 text-left"></th></tr></thead>
                    <tbody>{relatedAp.map((b) => (
                      <tr key={b.id} className={"border-t " + (isApPaid(b) ? "bg-green-50/50" : "")}>
                        <td className="p-2"><LinkBtn onClick={() => navigateTo("ap", { client: j.company, jobId: j.id, highlightId: b.id })}>{b.invoice_no}</LinkBtn></td>
                        <td className="p-2">{b.payee_type === "SI" ? t("payeeSi") : t("payeeVendor")}</td>
                        <td className="p-2">{b.company_name}</td>
                        <td className="p-2 text-right">{money(b.amount)} {b.currency}</td>
                        <td className="p-2"><span className={"px-2 py-0.5 rounded text-xs " + apPaymentClass(b)}>{apPaymentLabel(b, t)}</span></td>
                        <td className="p-2"><button type="button" className="text-xs text-blue-600 hover:underline" onClick={() => navigateTo("ap", { client: j.company, jobId: j.id, highlightId: b.id })}>{t("goToAp")}</button></td>
                      </tr>
                    ))}</tbody>
                  </table>
                ) : <p className="text-slate-500">{t("noAp")}</p>}
                <DetailActions
                  editLabel={t("edit")}
                  deleteLabel={t("delete")}
                  canEdit={can("job", "edit")}
                  canDelete={can("job", "delete")}
                  onEdit={() => {
                    setDetailPanel(null);
                    const norm = normalizeJobQuotations(normalizeJobPoFields(j));
                    setJobModal({ mode: "edit", id: j.id, data: { ...norm, client_id: String(j.client_id), amount: String(j.amount ?? ""), customer_po_amount: String(norm.customer_po_amount ?? ""), po_lines: (norm.po_lines || []).map((pl) => ({ ...pl, amount: String(pl.amount ?? "") })), _quotationDraft: "" } });
                  }}
                  onDelete={() => { deleteJob(j.id); setDetailPanel(null); }}
                  extra={isJobCompleted(j.status) && can("ar", "add") ? (
                    <button type="button" onClick={() => openArFromJob(j)} className="px-3 py-1.5 text-xs rounded-lg bg-green-600 text-white hover:bg-green-700 font-medium">{t("issueInvoiceToClient")}</button>
                  ) : null}
                />
              </Modal>
            );
          }

          if (type === "vendor") {
            const v = vendors.find((x) => x.id === id);
            if (!v) return null;
            const relatedAp = getApByVendor(v.name);
            return (
              <Modal title={t("vendorDetail") + " - " + (v.vendor_no ? v.vendor_no + " · " : "") + v.name} onClose={() => setDetailPanel(null)} wide>
                <div className="grid grid-cols-2 gap-4 text-sm mb-6">
                  <div><p className="text-slate-500">{t("vendorNo")}</p><p className="font-medium">{v.vendor_no || "-"}</p></div>
                  <div><p className="text-slate-500">Name</p><p className="font-medium">{v.name}</p></div>
                  <div><p className="text-slate-500">Email</p><p>{v.email}</p></div>
                  <div><p className="text-slate-500">{t("phoneNo")}</p><p>{v.phone || "-"}</p></div>
                  <div><p className="text-slate-500">{t("address")}</p><p>{v.address || "-"}</p></div>
                  <div><p className="text-slate-500">Bank</p><p>{v.bank}</p></div>
                  <div><p className="text-slate-500">Charge</p><p>{v.charge || "OUR"}</p></div>
                  <div><p className="text-slate-500">SWIFT</p><p>{v.swift_code || "-"}</p></div>
                  <div><p className="text-slate-500">Contact</p><p>{v.contact || "-"}</p></div>
                  <div><p className="text-slate-500">{t("colFinanceContact")}</p><p>{v.finance_contact || "-"}</p></div>
                  <div><p className="text-slate-500">{t("colFinanceEmail")}</p><p>{v.finance_email || "-"}</p></div>
                </div>
                <h4 className="font-semibold mb-2">{t("relatedAp")} ({relatedAp.length})</h4>
                {relatedAp.length ? (
                  <>
                    <p className="text-xs text-slate-500 mb-2">
                      {relatedAp.filter((b) => isApPaid(b)).length} {t("apPaid")} · {relatedAp.filter((b) => !isApPaid(b)).length} {t("apUnpaid")}
                    </p>
                  <table className="w-full text-sm border rounded-lg overflow-hidden">
                    <thead className="bg-slate-50"><tr><th className="p-2 text-left">{t("colInvoiceNo")}</th><th className="p-2 text-left">{t("colJobNo")}</th><th className="p-2 text-right">{amtInLabel}</th><th className="p-2 text-left">{t("colDueDate")}</th><th className="p-2 text-left">{t("paymentStatus")}</th><th className="p-2 text-left">{t("colPayDate")}</th><th className="p-2 text-left">{t("colRemarks")}</th><th className="p-2 text-left"></th></tr></thead>
                    <tbody>{relatedAp.map((b) => (
                      <tr key={b.id} className={"border-t " + (isApPaid(b) ? "bg-green-50/70" : "")}>
                        <td className="p-2"><LinkBtn onClick={() => navigateTo("ap", { highlightId: b.id, jobId: b.job_id })}>{b.invoice_no}</LinkBtn></td>
                        <td className="p-2">{b.job_no || "-"}</td>
                        <td className="p-2 text-right">{money(convertCurrency(b.amount, b.currency, regionListCurrency))}</td>
                        <td className="p-2">{b.due_date}</td>
                        <td className="p-2"><span className={"px-2 py-0.5 rounded text-xs " + apPaymentClass(b)}>{apPaymentLabel(b, t)}</span></td>
                        <td className="p-2">{b.pay_date || "-"}</td>
                        <td className="p-2">{b.remarks || "-"}</td>
                        <td className="p-2"><button type="button" className="text-xs text-blue-600 hover:underline" onClick={() => navigateTo("ap", { highlightId: b.id, jobId: b.job_id })}>{t("goToAp")}</button></td>
                      </tr>
                    ))}</tbody>
                  </table>
                  </>
                ) : <p className="text-slate-500">{t("noAp")}</p>}
                <DetailActions
                  editLabel={t("edit")}
                  deleteLabel={t("delete")}
                  canEdit={can("vendors", "edit")}
                  canDelete={can("vendors", "delete")}
                  onEdit={() => { setDetailPanel(null); setVendorModal({ mode: "edit", id: v.id, data: { ...v } }); }}
                  onDelete={() => { deleteVendor(v.id); setDetailPanel(null); }}
                />
              </Modal>
            );
          }

          if (type === "ar") {
            const r = arInvoices.find((x) => x.id === id);
            if (!r) return null;
            const job = jobs.find((j) => j.job_no === r.job_no);
            const arStatus = deriveArPaymentStatus(r);
            return (
              <Modal title={t("arDetail") + " - " + r.invoice_no} onClose={() => setDetailPanel(null)} wide>
                <div className="grid grid-cols-2 md:grid-cols-3 gap-4 text-sm">
                  <div><p className="text-slate-500">{t("colJobNo")}</p><p>{r.job_no ? <LinkBtn onClick={() => { const j = jobs.find((x) => x.job_no === r.job_no); navigateTo("job", { client: r.customer, highlightId: j ? j.id : null }); }}>{r.job_no}</LinkBtn> : "-"}</p></div>
                  <div><p className="text-slate-500">Customer</p><p>{r.customer}</p></div>
                  <div><p className="text-slate-500">{t("colCustomerPo")}</p><p>{r.customer_po ? <LinkBtn onClick={() => navigateToArByPo(r.customer, r.customer_po)}>{r.customer_po}</LinkBtn> : "-"}</p></div>
                  <div><p className="text-slate-500">{t("jobStatus")}</p><p>{getJobForAr(r) ? <span className={"px-2 py-0.5 rounded text-xs " + jobStatusClass(getJobForAr(r).status)}>{getJobForAr(r).status}</span> : "-"}</p></div>
                  <div><p className="text-slate-500">Invoice Amt</p><p>{money(r.invoice_amt)} {r.invoice_currency}</p></div>
                  <div><p className="text-slate-500">{amtInLabel}</p><p>{money(convertCurrency(r.invoice_amt, r.invoice_currency, regionListCurrency))}</p></div>
                  <div><p className="text-slate-500">{t("paymentStatus")}</p><p><span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(arStatus)}>{paymentStatusLabel(arStatus, t)}</span></p></div>
                </div>
                {job && <div className="mt-4 p-3 bg-slate-50 rounded-lg text-sm"><p className="font-medium">{t("linkedJob")}</p><p><LinkBtn onClick={() => navigateTo("job", { client: r.customer, highlightId: job.id })}>{job.job_no}</LinkBtn> · {job.job_type} · <span className={"px-1.5 py-0.5 rounded text-xs " + jobStatusClass(job.status)}>{job.status}</span></p></div>}
                <DetailActions
                  editLabel={t("edit")}
                  deleteLabel={t("delete")}
                  canEdit={can("ar", "edit")}
                  canDelete={can("ar", "delete")}
                  onEdit={() => { setDetailPanel(null); setArModal({ mode: "edit", id: r.id, data: { ...r, job_id: String(r.job_id || ""), invoice_amt: String(r.invoice_amt ?? ""), exchange_rate: String(r.exchange_rate_locked || ""), base_amount: String(r.base_amount_locked || r.base_amount || ""), payment_stage: r.payment_stage || "" } }); }}
                  onDelete={() => { deleteAr(r.id); setDetailPanel(null); }}
                  extra={arStatus !== "Paid" && can("ar", "edit") ? (
                    <button type="button" onClick={() => markArPaid(r.id)} className="px-3 py-1.5 text-xs rounded-lg bg-green-600 text-white hover:bg-green-700 font-medium">{t("markPaid")}</button>
                  ) : null}
                />
              </Modal>
            );
          }

          if (type === "ap") {
            const b = apBills.find((x) => x.id === id);
            if (!b) return null;
            const job = jobs.find((j) => j.id === b.job_id || j.job_no === b.job_no);
            const apStatus = deriveApPaymentStatus(b);
            return (
              <Modal title={t("apDetail") + " - " + b.invoice_no} onClose={() => setDetailPanel(null)} wide>
                <div className="grid grid-cols-2 md:grid-cols-3 gap-4 text-sm">
                  <div><p className="text-slate-500">{t("colJobNo")}</p><p>{b.job_no ? <LinkBtn onClick={() => navigateTo("job", { client: job ? job.company : "all", highlightId: job ? job.id : null })}>{b.job_no}</LinkBtn> : "-"}</p></div>
                  <div><p className="text-slate-500">{t("payeeType")}</p><p>{b.payee_type === "SI" ? t("payeeSi") : t("payeeVendor")}</p></div>
                  <div><p className="text-slate-500">Company</p><p>{b.company_name}</p></div>
                  <div><p className="text-slate-500">{t("amount")}</p><p>{money(b.amount)} {b.currency}</p></div>
                  <div><p className="text-slate-500">{amtInLabel}</p><p>{money(convertCurrency(b.amount, b.currency, regionListCurrency))}</p></div>
                  <div><p className="text-slate-500">{t("colDueDate")}</p><p>{b.due_date}</p></div>
                  <div><p className="text-slate-500">{t("paymentStatus")}</p><p><span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(apStatus)}>{paymentStatusLabel(apStatus, t)}</span></p></div>
                  <div><p className="text-slate-500">{t("colPayDate")}</p><p>{b.pay_date || "-"}</p></div>
                  <div><p className="text-slate-500">Bank</p><p>{b.bank}</p></div>
                  <div><p className="text-slate-500">SWIFT</p><p>{b.swift_code}</p></div>
                </div>
                <DetailActions
                  editLabel={t("edit")}
                  deleteLabel={t("delete")}
                  canEdit={can("ap", "edit")}
                  canDelete={can("ap", "delete")}
                  onEdit={() => { setDetailPanel(null); setApModal({ mode: "edit", id: b.id, data: { ...b, job_id: String(b.job_id || ""), si_id: String(b.si_id || ""), payee_type: b.payee_type || "Vendor", amount: String(b.amount ?? ""), exchange_rate: String(b.exchange_rate_locked || ""), base_amount: String(b.base_amount_locked || b.base_amount || ""), payment_stage: b.payment_stage || "" } }); }}
                  onDelete={() => { deleteAp(b.id); setDetailPanel(null); }}
                  extra={apStatus !== "Paid" && can("ap", "edit") ? (
                    <button type="button" onClick={() => markApPaid(b.id)} className="px-3 py-1.5 text-xs rounded-lg bg-green-600 text-white hover:bg-green-700 font-medium">{t("markApPaid")}</button>
                  ) : null}
                />
              </Modal>
            );
          }
          return null;
        }

        function renderJobRows(list) {
          return list.map((j) => (
            <tr key={j.id} className={"hover:bg-blue-50/50 cursor-pointer " + (isJobCompleted(j.status) ? "bg-green-50/70 " : "") + rowHighlightClass(j.id)} onClick={() => setDetailPanel({ type: "job", id: j.id })}>
              <td className="p-3 font-medium">{j.job_no}</td>
              <td className="p-3"><span className="px-2 py-0.5 rounded bg-blue-100 text-blue-800 text-xs">{j.company}</span></td>
              <td className="p-3 max-w-[12rem] erp-cell-truncate" title={j.description || ""}>{j.description || "-"}</td>
              <td className="p-3">{j.job_type}</td>
              <td className="p-3">{j.quotation_no || "-"}</td>
              <td className="p-3 text-right tabular-nums">{j.amount ? money(j.amount) : "-"}</td>
              <td className="p-3">{j.currency || "-"}</td>
              <td className="p-3">{j.customer_po || "-"}</td>
              <td className="p-3 text-right tabular-nums">{j.customer_po_amount ? money(j.customer_po_amount) : "-"}</td>
              <td className="p-3"><span className={"px-2 py-0.5 rounded text-xs " + jobStatusClass(j.status)}>{j.status}</span></td>
              <td className="p-3">{j.start_date || "-"}</td>
              <td className="p-3">{getArByJob(j.job_no).length}</td>
              <td className="p-3">{getApByJob(j.job_no).length}</td>
            </tr>
          ));
        }

        function renderArRows(list) {
          return list.map((r) => {
            const linkedJob = getJobForAr(r);
            const jobStatus = linkedJob ? linkedJob.status : "-";
            return (
            <tr key={r.id} className={"hover:bg-blue-50/50 cursor-pointer " + rowHighlightClass(r.id)} onClick={() => setDetailPanel({ type: "ar", id: r.id })}>
              <td className="p-3">{r.job_no}</td><td className="p-3">{r.customer}</td><td className="p-3">{r.customer_po || "-"}</td>
              <td className="p-3"><span className={"px-2 py-0.5 rounded text-xs " + (linkedJob ? jobStatusClass(linkedJob.status) : "bg-slate-100")}>{jobStatus}</span></td>
              <td className="p-3">{r.invoice_no}</td><td className="p-3">{paymentStageLabel(r.payment_stage, t)}</td><td className="p-3">{r.invoice_currency}</td><td className="p-3 text-right">{money(r.invoice_amt)}</td>
              <td className="p-3 text-right font-medium">{money(r.display_base_amount)}</td>
              <td className="p-3">{r.invoice_date}</td><td className="p-3">{r.due_date}</td>
              <td className="p-3 text-right">{r.payment_status === "Paid" ? 0 : calcOverdueDays(r.due_date)}</td><td className="p-3 text-right">{calcDaysBetween(r.invoice_date, r.due_date)}</td>
              <td className="p-3">{r.payment_received_date || "-"}</td>
              <td className="p-3"><span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(r.payment_status)}>{paymentStatusLabel(r.payment_status, t)}</span></td>
            </tr>
            );
          });
        }

        function renderApRows(list) {
          return list.map((r) => (
            <tr key={r.id} className={"hover:bg-blue-50/50 cursor-pointer " + (isApPaid(r) ? "bg-green-50/70 " : "") + rowHighlightClass(r.id)} onClick={() => setDetailPanel({ type: "ap", id: r.id })}>
              <td className="p-3">{r.job_no || "-"}</td>
              <td className="p-3"><span className="text-xs px-1.5 py-0.5 rounded bg-slate-100">{r.payee_type === "SI" ? t("payeeSi") : t("payeeVendor")}</span></td>
              <td className="p-3">{r.company_name}</td><td className="p-3">{r.invoice_no}</td><td className="p-3">{paymentStageLabel(r.payment_stage, t)}</td><td className="p-3">{r.invoice_date || "-"}</td><td className="p-3">{r.invoice_received_date || "-"}</td>
              <td className="p-3">{r.due_date || "-"}</td><td className="p-3">{r.currency}</td><td className="p-3 text-right">{money(r.amount)}</td>
              <td className="p-3 text-right font-medium">{money(r.display_base_amount)}</td>
              <td className="p-3 whitespace-nowrap"><span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(r.payment_status)}>{paymentStatusLabel(r.payment_status, t)}</span></td>
              <td className="p-3">{r.pay_date || "-"}</td><td className="p-3 erp-cell-truncate" title={r.remarks || ""}>{r.remarks || "-"}</td>
            </tr>
          ));
        }

        if (cloudOnlyMode && liveSyncEnabled && !liveSyncReady) {
          return (
            <div className="erp-login-shell">
              <div className="erp-login-card text-center space-y-3">
                <p className="text-slate-700 font-medium">{t("cloudOnlyLoading")}</p>
                <p className="text-sm text-slate-500">{t("cloudOnlyHint")}</p>
                {liveSyncStatus === "error" && (
                  <p className="text-sm text-red-600">{t("liveSyncError")} — {t("importSyncRetry")}</p>
                )}
              </div>
            </div>
          );
        }

        return (
          <div className="erp-app-shell flex min-h-screen w-full">
            <input
              ref={importFileInputRef}
              type="file"
              accept=".xlsx,.xls,.xlsm,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
              className="hidden"
              tabIndex={-1}
              aria-hidden="true"
              onChange={onImportFileSelected}
            />
            <aside className={(sidebarCollapsed ? "w-16" : "w-52") + " shrink-0 bg-slate-900 text-slate-100 py-3 transition-all duration-300 ease-in-out overflow-hidden flex flex-col min-h-screen " + (sidebarCollapsed ? "px-2" : "px-3")}>
              <div className={"mb-4 shrink-0 " + (sidebarCollapsed ? "flex flex-col items-center gap-2" : "px-0.5")}>
                <div className={"flex items-center w-full gap-2 " + (sidebarCollapsed ? "justify-center" : "")}>
                  <div className={sidebarCollapsed ? "w-10 rounded-md bg-white px-1 py-0.5 flex items-center justify-center overflow-hidden shrink-0" : "flex-1 min-w-0 rounded-md bg-white px-2 py-1 flex items-center overflow-hidden"}>
                    {sidebarLogo ? (
                      <img
                        src={sidebarLogo}
                        alt="AIRLINK"
                        className={sidebarCollapsed ? "max-h-8 max-w-full object-contain" : "h-8 w-full object-contain object-left"}
                        onError={(e) => {
                          const el = e.currentTarget;
                          if (el.dataset.fallback === "1") return;
                          el.dataset.fallback = "1";
                          el.src = "/AIRLINK/logo.png";
                        }}
                      />
                    ) : (
                      <svg className="w-5 h-5 text-slate-500" viewBox="0 0 24 24" fill="none" stroke="currentColor" strokeWidth="1.5" aria-hidden="true">
                        <rect x="3" y="5" width="18" height="14" rx="2" />
                        <circle cx="9" cy="10" r="1.5" />
                        <path d="M3 16l4.5-4.5 3 3L14 11l7 7" />
                      </svg>
                    )}
                  </div>
                  {!sidebarCollapsed && (
                    <button
                      type="button"
                      onClick={() => setSidebarCollapsed((v) => !v)}
                      className="shrink-0 text-slate-300 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 transition-colors duration-200"
                      title="Collapse sidebar"
                      aria-label="Collapse sidebar"
                    >
                      <svg {...NAV_ICON_SVG} className="w-4 h-4 transition-transform duration-300">
                        <rect width="18" height="18" x="3" y="3" rx="2" />
                        <path d="M9 3v18" />
                        <path d="m14 9-3 3 3 3" />
                      </svg>
                    </button>
                  )}
                </div>
                {sidebarCollapsed && (
                  <button
                    type="button"
                    onClick={() => setSidebarCollapsed((v) => !v)}
                    className="text-slate-300 hover:text-white p-1.5 rounded-lg hover:bg-slate-800 transition-colors duration-200"
                    title="Expand sidebar"
                    aria-label="Expand sidebar"
                  >
                    <svg {...NAV_ICON_SVG} className="w-4 h-4 rotate-180 transition-transform duration-300">
                      <rect width="18" height="18" x="3" y="3" rx="2" />
                      <path d="M9 3v18" />
                      <path d="m14 9-3 3 3 3" />
                    </svg>
                  </button>
                )}
              </div>
              <nav className="space-y-1 flex-1 min-h-0 overflow-y-auto">
                <SidebarItem id="dashboard" label={t("nav_dashboard")} module="dashboard" collapsed={sidebarCollapsed} />
                <SidebarItem id="management_report" label={t("nav_management_report")} module="management_report" collapsed={sidebarCollapsed} />
                <SidebarItem id="overview" label={t("nav_overview")} module="overview" collapsed={sidebarCollapsed} />
                <SidebarItem id="ongoing" label={t("nav_ongoing")} module="ongoing" collapsed={sidebarCollapsed} />
                <SidebarItem id="monthly_report" label={t("nav_monthly_report")} module="monthly_report" collapsed={sidebarCollapsed} />
                <SidebarItem id="clients" label={t("nav_clients")} module="clients" collapsed={sidebarCollapsed} />
                <SidebarItem id="quotation" label={t("nav_quotation")} module="quotation" collapsed={sidebarCollapsed} />
                <SidebarItem id="job" label={t("nav_job")} module="job" collapsed={sidebarCollapsed} />
                <SidebarItem id="ar" label={t("nav_ar")} module="ar" collapsed={sidebarCollapsed} />
                <SidebarItem id="ap" label={t("nav_ap")} module="ap" collapsed={sidebarCollapsed} />
                <SidebarItem id="vendors" label={t("nav_vendors")} module="vendors" collapsed={sidebarCollapsed} />
                <SidebarItem id="settings" label={t("nav_settings")} module="settings" collapsed={sidebarCollapsed} />
              </nav>
              <div className={"shrink-0 mt-3 pt-3 border-t border-slate-800 " + (sidebarCollapsed ? "px-0" : "px-0.5")}>
                {!sidebarCollapsed && getCurrentUser() && (
                  <p className="text-[10px] text-slate-500 truncate mb-2 px-1">{getCurrentUser().name} · {getCurrentUser().email}</p>
                )}
                <button
                  type="button"
                  onClick={onLogout}
                  className={"w-full flex items-center justify-center gap-1.5 rounded-lg text-xs font-semibold transition-colors " + (sidebarCollapsed ? "p-2 text-red-400 hover:bg-red-950/40 hover:text-red-300" : "px-3 py-2 bg-red-950/30 text-red-300 hover:bg-red-900/50 hover:text-white border border-red-900/50")}
                  title={t("logout")}
                >
                  <svg className="w-4 h-4 shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24" strokeWidth="2"><path strokeLinecap="round" strokeLinejoin="round" d="M17 16l4-4m0 0l-4-4m4 4H7m6 4v1a3 3 0 01-3 3H6a3 3 0 01-3-3V7a3 3 0 013-3h4a3 3 0 013 3v1" /></svg>
                  {!sidebarCollapsed && <span>{t("logout")}</span>}
                </button>
              </div>
            </aside>

            <main className="flex-1 min-w-0 p-5 space-y-5">
              {importStatus && (
                <div className={"rounded-lg border px-4 py-2 text-sm " + (isImportStatusSticky(importStatus) ? "border-amber-300 bg-amber-50 text-amber-900" : importPreview && importPreview.parseErrors && importPreview.parseErrors.length ? "border-amber-300 bg-amber-50 text-amber-900" : "border-blue-300 bg-blue-50 text-blue-900")}>
                  <div className="flex items-start justify-between gap-3">
                    <p className="flex-1 min-w-0">{importStatus}</p>
                    <div className="flex items-center gap-2 shrink-0">
                      {isImportStatusSticky(importStatus) && liveSyncEnabled && canLiveSyncPush() && erpSyncKey.trim() && (
                        <button type="button" onClick={() => { flushLiveSyncPush().catch(() => {}); }} className="text-xs font-semibold px-2 py-1 rounded border border-amber-400 bg-white hover:bg-amber-100">{t("importSyncRetry")}</button>
                      )}
                      <button type="button" onClick={() => setImportStatus("")} className="text-lg leading-none px-1 text-slate-500 hover:text-slate-800" title={t("importStatusDismiss")} aria-label={t("importStatusDismiss")}>×</button>
                    </div>
                  </div>
                  {importLoading && (
                    <div className="mt-2 h-1.5 rounded-full bg-blue-200 overflow-hidden">
                      <div className="h-full w-full bg-blue-600 rounded-full animate-pulse" />
                    </div>
                  )}
                </div>
              )}
              <header className="flex flex-wrap justify-between items-start gap-3">
                <div className="min-w-0">
                  <h2 className="text-2xl font-bold">
                    {page === "dashboard" && t("nav_dashboard")}
                    {page === "management_report" && t("nav_management_report")}
                    {page === "overview" && t("nav_overview")}
                    {page === "clients" && t("nav_clients")}
                    {page === "quotation" && t("nav_quotation")}
                    {page === "job" && t("nav_job")}
                    {page === "ongoing" && t("nav_ongoing")}
                    {page === "ar" && t("nav_ar")}
                    {page === "ap" && t("nav_ap")}
                    {page === "monthly_report" && t("nav_monthly_report")}
                    {page === "vendors" && t("nav_vendors")}
                    {page === "settings" && t("nav_settings")}
                  </h2>
                  <p className="text-sm text-slate-500">
                    {page === "dashboard" ? t("dashSubtitle") : page === "management_report" ? t("mgmtSubtitle") : page === "overview" ? t("overviewHint") : page === "ongoing" ? t("ongoingHint") : page === "monthly_report" ? t("monthlyReportHint") : <>{t("baseCurrency")}: {regionListCurrency} · {t("region")}: {regionLabel(activeRegion, lang)} · {t("hintClickRow")}</>}
                  </p>
                </div>
                <div className="flex items-center gap-2 shrink-0 flex-nowrap">
                  <div className="flex items-center gap-2 h-9 px-2.5 rounded-lg border border-slate-200 bg-white" title={t("worldTimeHint")}>
                    <span className="text-[10px] uppercase tracking-wide text-slate-400 whitespace-nowrap">{t("worldTime")}</span>
                    <span className="text-xs font-semibold text-slate-600 whitespace-nowrap">{timezoneLabel(worldTimeZone, lang)}</span>
                    <span className="text-slate-300">|</span>
                    <span className="text-xs font-mono font-semibold text-slate-700 tabular-nums whitespace-nowrap">{formatWorldTime(worldClock, worldTimeZone, lang)}</span>
                  </div>
                  <div className="flex items-center gap-1.5 h-9 px-2.5 rounded-lg border border-slate-200 bg-white">
                    <span className="text-[10px] text-slate-400 whitespace-nowrap">{t("currentUser")}</span>
                    <span className="text-xs font-semibold text-slate-700 whitespace-nowrap max-w-[8rem] truncate">{getCurrentUser() ? getCurrentUser().name : "—"}</span>
                    <span className="text-slate-300">|</span>
                    <button type="button" onClick={onLogout} className="h-9 px-3 text-xs font-semibold rounded-lg border border-red-200 bg-red-50 text-red-700 hover:bg-red-100 whitespace-nowrap">{t("logout")}</button>
                  </div>
                  <div className="flex items-center gap-1.5 h-9">
                    <span className="text-[10px] text-slate-400 whitespace-nowrap">{t("region")}</span>
                    <select value={activeRegion} onChange={(e) => setActiveRegion(e.target.value)} className="h-9 min-w-[4.5rem] rounded-lg border border-slate-300 bg-white text-xs px-2 font-semibold">
                      {getUserRegions().map((rid) => <option key={rid} value={rid}>{regionLabel(rid, lang)}</option>)}
                    </select>
                  </div>
                  <select value={baseCurrency} onChange={(e) => setBaseCurrency(e.target.value)} className="h-9 w-[4.5rem] shrink-0 rounded-lg border border-slate-300 bg-white text-xs px-2">
                    <option value="HKD">HKD</option><option value="USD">USD</option><option value="RMB">RMB</option>
                    <option value="TWD">TWD</option><option value="SGD">SGD</option><option value="EUR">EUR</option><option value="OTHER">Other</option>
                  </select>
                  {baseCurrency === "OTHER" && <Input placeholder="e.g. MYD" className="w-16 h-9 text-xs px-2 bg-white shrink-0" value={customCurrency} onChange={(e) => setCustomCurrency(e.target.value)} />}
                  {(can("ar", "add") || can("ap", "add")) && <button type="button" onClick={() => setNewTxOpen(true)} className="h-9 shrink-0 px-3 text-xs rounded-lg bg-blue-600 text-white hover:bg-blue-700 whitespace-nowrap">{t("newTransaction")}</button>}
                </div>
              </header>

              {page === "dashboard" && (
                <div className="space-y-5">
                  {dashRegionMismatch && (
                    <div className="rounded-xl border border-blue-200 bg-blue-50 px-4 py-3 text-sm text-blue-900">
                      {t("dashRegionHint")}
                    </div>
                  )}
                  {dashShowsAllPaid && (
                    <div className="rounded-xl border border-amber-200 bg-amber-50 px-4 py-3 text-sm text-amber-900">
                      {t("dashAllPaidHint")}
                    </div>
                  )}
                  {/* KPI — Boss sees cash position at a glance */}
                  <section className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-4 gap-4">
                    <DashGradientCard
                      title={t("dashCashIn")}
                      value={money(arTotal)}
                      hint={regionListCurrency + " · " + t("dashCashInHint")}
                      badge={overdueArCount > 0 ? overdueArCount + " " + t("overdue") + " · " + money(overdueArTotal) : null}
                      gradient={DASH_GRADIENTS.blue}
                      icon={<DashCardIcon><path d="M3 3v18h18" /><path d="M7 16l4-8 4 5 4-9" /></DashCardIcon>}
                    />
                    <DashGradientCard
                      title={t("dashCashOut")}
                      value={money(apTotal)}
                      hint={regionListCurrency + " · " + t("dashCashOutHint")}
                      badge={overdueApCount > 0 ? overdueApCount + " " + t("overdue") + " · " + money(overdueApTotal) : null}
                      gradient={DASH_GRADIENTS.redOrange}
                      icon={<DashCardIcon><rect x="2" y="5" width="20" height="14" rx="2" /><path d="M2 10h20" /><path d="M6 15h.01" /><path d="M10 15h4" /></DashCardIcon>}
                    />
                    <DashGradientCard
                      title={t("netPosition")}
                      value={money(netPosition)}
                      hint={regionListCurrency + " · " + t("dashNetHint")}
                      badge={t("dashThisMonthIn") + ": " + money(thisMonthIn) + "  |  " + t("dashThisMonthOut") + ": " + money(thisMonthOut)}
                      gradient={netPosition >= 0 ? DASH_GRADIENTS.purple : DASH_GRADIENTS.redOrange}
                      icon={<DashCardIcon><circle cx="12" cy="12" r="9" /><path d="M8 12h8" /><path d="M12 8v8" /></DashCardIcon>}
                    />
                    <DashGradientCard
                      title={t("dashActiveJobs")}
                      value={String(activeJobCount)}
                      hint={t("dashActiveJobsHint")}
                      badge={t("dashArDueSoon") + ": " + arDueSoonList.length + "  |  " + t("dashApDueSoon") + ": " + apDueSoonList.length}
                      gradient={DASH_GRADIENTS.indigo}
                      icon={<DashCardIcon><path d="M9 5H7a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2V7a2 2 0 00-2-2h-2" /><rect x="9" y="3" width="6" height="4" rx="1" /><path d="M9 12h6" /><path d="M9 16h6" /></DashCardIcon>}
                    />
                  </section>

                  {/* Income vs Out + Global Map (when ALL) */}
                  <section className={activeRegion === "ALL" ? "grid grid-cols-1 xl:grid-cols-2 gap-4" : ""}>
                    <article className={"bg-white border border-slate-200 rounded-2xl p-4 shadow-sm w-full " + (activeRegion === "ALL" ? "flex flex-col h-full min-h-0" : "")}>
                      <div className="flex flex-wrap items-start justify-between gap-2 mb-3 shrink-0">
                        <div>
                          <h3 className="font-semibold text-lg text-slate-800">{t("dashIncomeOut")}</h3>
                          <p className="text-xs text-slate-500 mt-0.5">{t("dashIncomeOutHint")} · {regionListCurrency}</p>
                        </div>
                        <div className="flex items-center gap-3 text-xs shrink-0">
                          <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-sm" style={{ background: "linear-gradient(135deg, #1d4ed8, #38bdf8)" }}></span>{t("dashCashInLegend")}</span>
                          <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-sm" style={{ background: "linear-gradient(135deg, #ea580c, #fbbf24)" }}></span>{t("dashCashOutLegend")}</span>
                        </div>
                      </div>
                      <div className={"flex flex-col gap-3 " + (activeRegion === "ALL" ? "flex-1 min-h-0 justify-end" : "")}>
                        <MonthlyBarChart
                          firstHalf={cashFlowFirstHalf}
                          secondHalf={cashFlowSecondHalf}
                          maxBar={maxBar}
                          t={t}
                          inLegend={t("dashCashInLegend")}
                          outLegend={t("dashCashOutLegend")}
                          h1Label={t("dashIncomeOutH1")}
                          h2Label={t("dashIncomeOutH2")}
                          totals={cashFlowTotals}
                          netLabel={t("dashNetHint")}
                        />
                      </div>
                    </article>
                    {activeRegion === "ALL" && (
                      <GlobalFinancialMap
                        stats={regionalMapStats}
                        lang={lang}
                        currency={regionListCurrency}
                        t={t}
                        onRegionClick={(rid) => setActiveRegion(rid)}
                      />
                    )}
                  </section>

                  {/* Due soon — AR & AP side by side */}
                  <section className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                    <article className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
                      <div className="px-4 py-3 border-b bg-emerald-50/60 flex justify-between items-center">
                        <h3 className="font-semibold text-emerald-900">{t("dashArDueSoon")}</h3>
                        {can("ar", "view") && <button type="button" onClick={() => setPage("ar")} className="text-xs text-emerald-700 hover:underline">{t("dashViewAll")}</button>}
                      </div>
                      {arDueSoonList.length ? arDueSoonList.map((r) => <DashDueRow key={r.id} item={r} type="ar" nameKey="customer" noKey="invoice_no" />) : (
                        <p className="px-4 py-8 text-sm text-slate-400 text-center">{t("dashNoDueSoon")}</p>
                      )}
                    </article>
                    <article className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
                      <div className="px-4 py-3 border-b bg-rose-50/60 flex justify-between items-center">
                        <h3 className="font-semibold text-rose-900">{t("dashApDueSoon")}</h3>
                        {can("ap", "view") && <button type="button" onClick={() => setPage("ap")} className="text-xs text-rose-700 hover:underline">{t("dashViewAll")}</button>}
                      </div>
                      {apDueSoonList.length ? apDueSoonList.map((b) => <DashDueRow key={b.id} item={b} type="ap" nameKey="company_name" noKey="invoice_no" />) : (
                        <p className="px-4 py-8 text-sm text-slate-400 text-center">{t("dashNoDueSoon")}</p>
                      )}
                    </article>
                  </section>

                  {/* Overdue + Aging */}
                  <section className="grid grid-cols-1 xl:grid-cols-3 gap-4">
                    <article className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden xl:col-span-1">
                      <div className="px-4 py-3 border-b bg-red-50 flex justify-between items-center">
                        <h3 className="font-semibold text-red-900">{t("dashOverdueAr")}</h3>
                        <span className="text-xs font-semibold text-red-700">{money(overdueArTotal)}</span>
                      </div>
                      {arOverdueList.length ? arOverdueList.map((r) => <DashDueRow key={r.id} item={r} type="ar" nameKey="customer" noKey="invoice_no" />) : (
                        <p className="px-4 py-6 text-sm text-slate-400 text-center">{t("dashNoOverdue")}</p>
                      )}
                    </article>
                    <article className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden xl:col-span-1">
                      <div className="px-4 py-3 border-b bg-red-50 flex justify-between items-center">
                        <h3 className="font-semibold text-red-900">{t("dashOverdueAp")}</h3>
                        <span className="text-xs font-semibold text-red-700">{money(overdueApTotal)}</span>
                      </div>
                      {apOverdueList.length ? apOverdueList.map((b) => <DashDueRow key={b.id} item={b} type="ap" nameKey="company_name" noKey="invoice_no" />) : (
                        <p className="px-4 py-6 text-sm text-slate-400 text-center">{t("dashNoOverdue")}</p>
                      )}
                    </article>
                    <article className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm xl:col-span-1">
                      <h3 className="font-semibold text-slate-800 mb-4">{t("arAging")} <span className="text-xs font-normal text-slate-500">({regionListCurrency})</span></h3>
                      <ul className="space-y-3.5">
                        {[
                          { key: "current", label: t("dashAgingCurrent"), color: "bg-emerald-500", val: aging.current },
                          { key: "d1_30", label: t("dashAging1_30"), color: "bg-amber-400", val: aging.d1_30 },
                          { key: "d31_60", label: t("dashAging31_60"), color: "bg-orange-400", val: aging.d31_60 },
                          { key: "d61_90", label: t("dashAging61_90"), color: "bg-orange-500", val: aging.d61_90 },
                          { key: "d90_plus", label: t("dashAging90plus"), color: "bg-red-500", val: aging.d90_plus }
                        ].map((row) => (
                          <li key={row.key}>
                            <div className="flex justify-between text-xs mb-1">
                              <span className="text-slate-600">{row.label}</span>
                              <span className="font-semibold tabular-nums">{money(row.val)}</span>
                            </div>
                            <div className="h-2 rounded-full bg-slate-100 overflow-hidden">
                              <div className={"h-full rounded-full transition-all duration-500 " + row.color} style={{ width: Math.max(row.val ? 4 : 0, (row.val / agingMax) * 100) + "%" }} />
                            </div>
                          </li>
                        ))}
                      </ul>
                    </article>
                  </section>
                </div>
              )}

              {page === "management_report" && can("management_report", "view") && (
                <div className="space-y-5">
                  <section className="flex flex-wrap items-end gap-3 p-4 bg-white border border-slate-200 rounded-2xl shadow-sm">
                    <div className="min-w-[10rem]">
                      <label className="block text-[10px] uppercase tracking-wide text-slate-400 mb-1">{t("mgmtFilterClient")}</label>
                      <div className="w-full min-w-[18rem]">
                        <SearchableSelect
                          value={mgmtClientFilter}
                          options={mgmtClientSearchOptions}
                          placeholder={t("selectClientHint")}
                          noResultsText={t("noMatchFound")}
                          onChange={(nextValue, option) => setMgmtClientFilter(nextValue || (option ? option.value : "all") || "all")}
                        />
                      </div>
                    </div>
                    {activeRegion === "ALL" && (
                      <>
                        <div>
                          <label className="block text-[10px] uppercase tracking-wide text-slate-400 mb-1">{t("mgmtGroupCurrency")}</label>
                          <select value={mgmtGroupCurrency} onChange={(e) => setMgmtGroupCurrency(e.target.value)} className="h-9 w-[5.5rem] rounded-lg border border-slate-300 bg-white text-sm px-2 font-semibold">
                            {MGMT_GROUP_CURRENCIES.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                        <div>
                          <label className="block text-[10px] uppercase tracking-wide text-slate-400 mb-1">{t("mgmtCompareCurrency")}</label>
                          <select value={mgmtCompareCurrency} onChange={(e) => setMgmtCompareCurrency(e.target.value)} className="h-9 w-[5.5rem] rounded-lg border border-slate-300 bg-white text-sm px-2 font-semibold">
                            {MGMT_GROUP_CURRENCIES.map((c) => <option key={c} value={c}>{c}</option>)}
                          </select>
                        </div>
                      </>
                    )}
                    {mgmtClientFilter !== "all" && (
                      <button type="button" onClick={() => setMgmtClientFilter("all")} className="h-9 px-3 text-xs rounded-lg border border-slate-300 bg-slate-50 hover:bg-slate-100 text-slate-700">
                        {t("mgmtClearClient")}
                      </button>
                    )}
                  </section>

                  {mgmtClientFilter !== "all" && (
                    <div className="rounded-xl border border-indigo-200 bg-indigo-50 px-4 py-3 text-sm text-indigo-900">
                      {t("mgmtShowingClient").replace("{name}", mgmtClientFilter)}
                    </div>
                  )}

                  {!mgmtReport.hasData ? (
                    <div className="rounded-xl border border-slate-200 bg-slate-50 px-6 py-12 text-center text-slate-500">
                      {t("mgmtNoQualifyingData")}
                    </div>
                  ) : (
                    <>
                      <section className="grid grid-cols-1 md:grid-cols-3 gap-4">
                        <DashGradientCard
                          title={t("mgmtGrossRevenue")}
                          value={money(mgmtReport.grossRevenue) + " " + mgmtReport.displayCurrency}
                          hint={t("mgmtRevenueHint")}
                          badge={activeRegion === "ALL" ? t("mgmtAggregatedRegions").replace("{n}", String(mgmtReport.regionCount)).replace("{currency}", mgmtReport.displayCurrency) : mgmtReport.displayCurrency}
                          gradient={MGMT_CARD_GRADIENTS.revenue}
                          icon={<DashCardIcon><path d="M12 2v20M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6" /></DashCardIcon>}
                        />
                        <DashGradientCard
                          title={t("mgmtTotalCogs")}
                          value={money(mgmtReport.totalCogs) + " " + mgmtReport.displayCurrency}
                          hint={t("mgmtCogsHint")}
                          gradient={MGMT_CARD_GRADIENTS.cogs}
                          icon={<DashCardIcon><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z" /></DashCardIcon>}
                        />
                        <DashGradientCard
                          title={t("mgmtNetProfit")}
                          value={money(mgmtReport.netProfit) + " " + mgmtReport.displayCurrency}
                          subValue={mgmtReport.grossRevenue > 0 ? t("mgmtGrossMargin") + " " + ((mgmtReport.netProfit / mgmtReport.grossRevenue) * 100).toFixed(1) + "%" : null}
                          hint={mgmtReport.displayCurrency}
                          badge={mgmtReport.clientProfitSharePct != null ? t("mgmtClientProfitShare").replace("{pct}", mgmtReport.clientProfitSharePct.toFixed(1)) : null}
                          gradient={mgmtReport.netProfit >= 0 ? MGMT_CARD_GRADIENTS.profit : MGMT_CARD_GRADIENTS.loss}
                          icon={<DashCardIcon><path d="M22 12h-4l-3 9L9 3l-3 9H2" /></DashCardIcon>}
                        />
                      </section>

                      {activeRegion === "ALL" && mgmtClientFilter === "all" && mgmtReport.regionBreakdown.length > 0 && (
                        <section>
                          <h3 className="font-semibold text-lg text-slate-800 mb-3">{t("mgmtRegionBreakdownTitle")}</h3>
                          <div className="grid grid-cols-1 sm:grid-cols-2 xl:grid-cols-3 gap-4">
                            {mgmtReport.regionBreakdown.map((reg) => (
                              <article key={reg.regionId} className="bg-white border border-slate-200 rounded-2xl p-4 shadow-sm">
                                <div className="flex items-center gap-2 mb-3">
                                  <span className="text-xl">{REGION_FLAGS[reg.regionId]}</span>
                                  <h4 className="font-semibold text-slate-800">{regionLabel(reg.regionId, lang)}</h4>
                                </div>
                                <div className="space-y-2 mb-3">
                                  <div className="rounded-lg bg-slate-50 px-3 py-2">
                                    <p className="text-[10px] uppercase text-slate-400">{t("mgmtLocalAmount")}</p>
                                    <p className="text-lg font-bold tabular-nums text-slate-800">{money(reg.localProfit)} {reg.localCurrency}</p>
                                    <p className="text-xs text-slate-500">{t("mgmtNetProfit")}</p>
                                  </div>
                                  <div className="rounded-lg bg-indigo-50 px-3 py-2">
                                    <p className="text-[10px] uppercase text-indigo-500">{t("mgmtConvertedAmount")}</p>
                                    <p className="text-lg font-bold tabular-nums text-indigo-900">{money(reg.convertedProfit)} {mgmtReport.compareCurrency}</p>
                                  </div>
                                </div>
                                {reg.clients.length > 0 && (
                                  <ul className="space-y-1.5 border-t border-slate-100 pt-2 max-h-36 overflow-y-auto">
                                    {reg.clients.map((c) => (
                                      <li key={c.company} className="text-xs flex justify-between gap-2">
                                        <button type="button" onClick={() => setMgmtClientFilter(c.company)} className="text-left text-blue-700 hover:underline truncate">{c.company}</button>
                                        <span className="shrink-0 tabular-nums text-slate-600">{money(c.profit)} · {c.marginPct.toFixed(0)}%</span>
                                      </li>
                                    ))}
                                  </ul>
                                )}
                              </article>
                            ))}
                          </div>
                        </section>
                      )}

                      <section className="grid grid-cols-1 xl:grid-cols-2 gap-4">
                        {mgmtClientFilter === "all" && (
                          <article className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm">
                            <h3 className="font-semibold text-lg text-slate-800">{t("mgmtTopClientsTitle")}</h3>
                            <p className="text-xs text-slate-500 mt-0.5 mb-4">{t("mgmtTopClientsHint")}</p>
                            {mgmtReport.clientMargins.length ? (
                              <div className="divide-y divide-slate-100">
                                {mgmtReport.clientMargins.map((row, i) => (
                                  <div key={row.company} className="flex items-start gap-2">
                                    <span className="text-xs font-bold text-slate-400 w-5 pt-3 shrink-0">{i + 1}</span>
                                    <div className="flex-1 min-w-0">
                                      <MgmtMarginBar row={row} maxProfit={Math.max(...mgmtReport.clientMargins.map((c) => Math.max(0, c.profit)), 1)} onSelect={setMgmtClientFilter} t={t} />
                                    </div>
                                  </div>
                                ))}
                              </div>
                            ) : (
                              <p className="text-sm text-slate-400 text-center py-6">{t("mgmtNoQualifyingData")}</p>
                            )}
                          </article>
                        )}

                        <article className={"bg-white border border-slate-200 rounded-2xl p-5 shadow-sm " + (mgmtClientFilter !== "all" ? "xl:col-span-2" : "")}>
                          <h3 className="font-semibold text-lg text-slate-800">{t("mgmtJobTypeTitle")}</h3>
                          <p className="text-xs text-slate-500 mt-0.5 mb-4">{t("mgmtJobTypeHint")}</p>
                          {mgmtReport.jobTypeBreakdown.length ? (
                            <div className="flex flex-col sm:flex-row items-center gap-6">
                              <MgmtDonutChart
                                breakdown={mgmtReport.jobTypeBreakdown}
                                size={176}
                                centerLabel={t("mgmtNetProfit")}
                                centerSub={mgmtReport.displayCurrency}
                              />
                              <ul className="flex-1 w-full">
                                <li className="grid grid-cols-[minmax(0,1fr)_5.5rem_2.75rem] gap-x-3 text-[10px] uppercase tracking-wide text-slate-400 pb-1 mb-1 border-b border-slate-100">
                                  <span />
                                  <span className="text-right">{mgmtReport.displayCurrency}</span>
                                  <span className="text-right">%</span>
                                </li>
                                {mgmtReport.jobTypeBreakdown.map((row) => {
                                  const palette = MGMT_JOB_TYPE_COLORS[row.type] || { to: "#94a3b8" };
                                  return (
                                    <li key={row.type} className="grid grid-cols-[minmax(0,1fr)_5.5rem_2.75rem] gap-x-3 items-center text-sm py-1.5">
                                      <span className="flex items-center gap-2 min-w-0">
                                        <span className="w-2.5 h-2.5 rounded-full shrink-0 ring-2 ring-white shadow-sm" style={{ background: "linear-gradient(135deg, " + (MGMT_JOB_TYPE_COLORS[row.type]?.from || "#475569") + ", " + palette.to + ")" }} />
                                        <span className="text-slate-700 truncate">{mgmtNatureLabel(row.type, t)}</span>
                                      </span>
                                      <span className="text-right tabular-nums font-semibold text-slate-800">{money(row.profit)}</span>
                                      <span className="text-right tabular-nums text-xs text-slate-500">{row.sharePct.toFixed(0)}%</span>
                                    </li>
                                  );
                                })}
                              </ul>
                            </div>
                          ) : (
                            <p className="text-sm text-slate-400 text-center py-6">{t("mgmtNoQualifyingData")}</p>
                          )}
                        </article>
                      </section>

                      <section>
                        <article className="bg-white border border-slate-200 rounded-2xl p-5 shadow-sm">
                          <h3 className="font-semibold text-lg text-slate-800">
                            {mgmtClientFilter === "all" ? t("mgmtCashForecastTitle") : t("mgmtClientRiskTitle")}
                          </h3>
                          <p className="text-xs text-slate-500 mt-0.5 mb-4">
                            {mgmtClientFilter === "all" ? t("mgmtCashForecastHint") : t("mgmtClientRiskHint")}
                          </p>
                          {mgmtClientFilter !== "all" && mgmtReport.overdueAr > 0 && (
                            <div className="mb-4 rounded-lg border border-red-200 bg-red-50 px-4 py-3 flex justify-between items-center">
                              <span className="text-sm font-medium text-red-800">{t("mgmtOverdueAr")}</span>
                              <span className="text-lg font-bold tabular-nums text-red-700">{money(mgmtReport.overdueAr)} {mgmtReport.displayCurrency}</span>
                            </div>
                          )}
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                            {mgmtReport.cashForecast.map((mo) => {
                              const cashAlert = mo.netCash < 0;
                              return (
                              <div key={mo.key} className={"rounded-xl border p-4 " + (cashAlert ? "border-red-500/50 bg-red-950/20" : "border-slate-200 bg-slate-50/50")}>
                                <p className="font-semibold text-slate-800 mb-3">{mo.label}</p>
                                {mgmtClientFilter === "all" ? (
                                  <>
                                    <div className="space-y-2 text-sm">
                                      <div className="flex justify-between"><span className="text-slate-600">{t("mgmtCashIn")}</span><span className="font-semibold tabular-nums text-emerald-700">{money(mo.cashIn)}</span></div>
                                      <div className="flex justify-between"><span className="text-slate-600">{t("mgmtCashOut")}</span><span className="font-semibold tabular-nums text-orange-700">{money(mo.cashOut)}</span></div>
                                      <div className="flex justify-between pt-2 border-t border-slate-200"><span className="font-medium">{t("mgmtNetCashFlow")}</span><span className={"font-bold tabular-nums " + (mo.netCash >= 0 ? "text-emerald-700" : "text-red-600")}>{mo.netCash >= 0 ? "+" : ""}{money(mo.netCash)}</span></div>
                                    </div>
                                    {cashAlert && (
                                      <p className="mt-3 text-xs font-medium text-red-700 flex items-start gap-1">
                                        <span>⚠️</span> {t("mgmtDryRunAlert")}
                                      </p>
                                    )}
                                  </>
                                ) : (
                                  <div className="space-y-2 text-sm">
                                    <div className="flex justify-between"><span className="text-slate-600">{t("mgmtExpectedInMonth")}</span><span className="font-semibold tabular-nums text-emerald-700">{money(mo.cashIn)} {mgmtReport.displayCurrency}</span></div>
                                    <div className="flex justify-between pt-2 border-t border-slate-200"><span className="font-medium">{t("mgmtNetCashFlow")}</span><span className="font-bold tabular-nums text-slate-800">{money(mo.cashIn)}</span></div>
                                  </div>
                                )}
                                <p className="text-[10px] text-slate-400 mt-2">{mgmtReport.displayCurrency}</p>
                              </div>
                            );})}
                          </div>
                        </article>
                      </section>
                    </>
                  )}
                </div>
              )}

              {page === "overview" && can("overview", "view") && (
                <div className="space-y-4">
                  <section className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4">
                    <div className="flex flex-wrap items-start justify-between gap-2 mb-3">
                      <div>
                        <h3 className="font-semibold text-lg text-slate-800">{t("overviewPaidTitle")}</h3>
                        <p className="text-xs text-slate-500 mt-0.5">{t("overviewPaidHint")} · {regionListCurrency}</p>
                      </div>
                      <div className="flex items-center gap-3 text-xs shrink-0">
                        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-sm" style={{ background: "linear-gradient(135deg, #1d4ed8, #38bdf8)" }}></span>{t("overviewArCollected")}</span>
                        <span className="flex items-center gap-1.5"><span className="w-3 h-3 rounded-sm" style={{ background: "linear-gradient(135deg, #ea580c, #fbbf24)" }}></span>{t("overviewApPaid")}</span>
                      </div>
                    </div>
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-4">
                      <div className="rounded-xl bg-gradient-to-br from-blue-600 to-sky-400 px-4 py-3 text-white shadow-sm">
                        <p className="text-[10px] text-blue-100 uppercase tracking-wide">{t("overviewArCollectedTotal")}</p>
                        <p className="text-2xl font-bold tabular-nums mt-1">{money(overviewCollectedTotals.in)}</p>
                        <p className="text-[10px] text-blue-100 mt-0.5">{regionListCurrency}</p>
                      </div>
                      <div className="rounded-xl bg-gradient-to-br from-orange-600 to-amber-400 px-4 py-3 text-white shadow-sm">
                        <p className="text-[10px] text-orange-100 uppercase tracking-wide">{t("overviewApPaidTotal")}</p>
                        <p className="text-2xl font-bold tabular-nums mt-1">{money(overviewCollectedTotals.out)}</p>
                        <p className="text-[10px] text-orange-100 mt-0.5">{regionListCurrency}</p>
                      </div>
                      <div className={"rounded-xl px-4 py-3 text-white shadow-sm " + (overviewCollectedTotals.in >= overviewCollectedTotals.out ? "bg-gradient-to-br from-emerald-600 to-teal-400" : "bg-gradient-to-br from-amber-600 to-orange-400")}>
                        <p className="text-[10px] text-white/80 uppercase tracking-wide">{t("overviewNet")}</p>
                        <p className="text-2xl font-bold tabular-nums mt-1">{money(overviewCollectedTotals.in - overviewCollectedTotals.out)}</p>
                        <p className="text-[10px] text-white/80 mt-0.5">{regionListCurrency}</p>
                      </div>
                    </div>
                    <MonthlyBarChart
                      firstHalf={overviewCollectedFirstHalf}
                      secondHalf={overviewCollectedSecondHalf}
                      maxBar={overviewCollectedMaxBar}
                      t={t}
                      inLegend={t("overviewArCollected")}
                      outLegend={t("overviewApPaid")}
                      h1Label={t("dashIncomeOutH1")}
                      h2Label={t("dashIncomeOutH2")}
                      totals={overviewCollectedTotals}
                      netLabel={t("overviewNet")}
                    />
                  </section>

                  <section className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4">
                    <div className="flex flex-wrap items-end gap-3 mb-4">
                      <Field label={t("overviewStartDate")}>
                        <Input type="date" value={overviewFrom} onChange={(e) => setOverviewFrom(e.target.value)} className="h-9" />
                      </Field>
                      <Field label={t("overviewEndDate")}>
                        <Input type="date" value={overviewTo} onChange={(e) => setOverviewTo(e.target.value)} className="h-9" />
                      </Field>
                    </div>
                    <p className="text-xs text-slate-500 mb-2">{t("overviewByInvoiceDate")} · {regionListCurrency}</p>
                    <OverviewDonutChart inTotal={overviewReport.inTotal} outTotal={overviewReport.outTotal} currency={regionListCurrency} t={t} />
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mt-4 pt-4 border-t border-slate-100">
                      <div className="rounded-lg bg-blue-50/80 px-3 py-2 text-center">
                        <p className="text-[10px] text-slate-500">{t("overviewArCount")}</p>
                        <p className="text-xl font-bold text-blue-700">{overviewReport.arRows.length}</p>
                      </div>
                      <div className="rounded-lg bg-orange-50/80 px-3 py-2 text-center">
                        <p className="text-[10px] text-slate-500">{t("overviewApCount")}</p>
                        <p className="text-xl font-bold text-orange-600">{overviewReport.apRows.length}</p>
                      </div>
                      <div className="rounded-lg bg-slate-50 px-3 py-2 text-center">
                        <p className="text-[10px] text-slate-500">{t("overviewNet")}</p>
                        <p className={"text-xl font-bold tabular-nums " + (overviewReport.inTotal >= overviewReport.outTotal ? "text-emerald-700" : "text-amber-700")}>{money(overviewReport.inTotal - overviewReport.outTotal)}</p>
                      </div>
                    </div>
                  </section>

                  <section className="grid grid-cols-1 lg:grid-cols-2 gap-4">
                    <article className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                      <div className="px-4 py-3 border-b bg-blue-50/60">
                        <h3 className="font-semibold text-blue-900">{t("nav_ar")} ({overviewReport.arRows.length})</h3>
                      </div>
                      <div className="erp-list-scroll erp-list-scroll--dash">
                        {overviewReport.arRows.length ? (
                          <table className="w-full text-sm">
                            <thead className="bg-slate-50 text-slate-500">
                              <tr>
                                <th className="p-2 text-left">{t("colInvoiceNo")}</th><th className="p-2 text-left">{t("colCustomer")}</th><th className="p-2 text-left">{t("colDate")}</th><th className="p-2 text-right">{t("amount")}</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y">
                              {overviewReport.arRows.map((r) => (
                                <tr key={r.id} className="hover:bg-blue-50/40 cursor-pointer" onClick={() => setDetailPanel({ type: "ar", id: r.id })}>
                                  <td className="p-2 font-medium">{r.invoice_no}</td>
                                  <td className="p-2">{r.customer}</td>
                                  <td className="p-2">{r.invoice_date}</td>
                                  <td className="p-2 text-right tabular-nums">{money(r.display_base_amount)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        ) : (
                          <p className="px-4 py-8 text-sm text-slate-400 text-center">{t("overviewNoData")}</p>
                        )}
                      </div>
                    </article>
                    <article className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                      <div className="px-4 py-3 border-b bg-orange-50/60">
                        <h3 className="font-semibold text-orange-900">{t("nav_ap")} ({overviewReport.apRows.length})</h3>
                      </div>
                      <div className="erp-list-scroll erp-list-scroll--dash">
                        {overviewReport.apRows.length ? (
                          <table className="w-full text-sm">
                            <thead className="bg-slate-50 text-slate-500">
                              <tr>
                                <th className="p-2 text-left">{t("colInvoiceNo")}</th><th className="p-2 text-left">{t("colPayee")}</th><th className="p-2 text-left">{t("colDate")}</th><th className="p-2 text-right">{t("amount")}</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y">
                              {overviewReport.apRows.map((b) => (
                                <tr key={b.id} className="hover:bg-orange-50/40 cursor-pointer" onClick={() => setDetailPanel({ type: "ap", id: b.id })}>
                                  <td className="p-2 font-medium">{b.invoice_no}</td>
                                  <td className="p-2">{b.company_name}</td>
                                  <td className="p-2">{b.invoice_date}</td>
                                  <td className="p-2 text-right tabular-nums">{money(b.display_base_amount)}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        ) : (
                          <p className="px-4 py-8 text-sm text-slate-400 text-center">{t("overviewNoData")}</p>
                        )}
                      </div>
                    </article>
                  </section>
                </div>
              )}

              {page === "clients" && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  <div className="p-4 border-b flex justify-between items-center gap-2 flex-wrap">
                    <h3 className="font-semibold">{t("nav_clients")}</h3>
                    <div className="flex items-center gap-2 flex-wrap">
                      <ListImportExportBar module="clients" />
                      {can("clients", "add") && <button onClick={() => setClientModal({ mode: "add", data: { ...emptyClient(), customer_no: nextCustomerNo(scopedClients) } })} className="px-3 py-2 text-sm rounded-lg bg-blue-600 text-white">{t("addClient")}</button>}
                    </div>
                  </div>
                  <ListToolbar t={t} mode="clients" searchQ={clientsSearch} setSearchQ={setClientsSearch} lang={lang} />
                  <div className="erp-list-scroll erp-list-scroll--compact">
                    <table className="w-full text-sm">
                      <thead className="bg-slate-50 text-slate-500">
                        <tr>
                          <SortableTh module="clients" field="customer_no" label={t("colCustomerNo")} />
                          <th className="p-3 text-left">{t("colBuNo")}</th>
                          <SortableTh module="clients" field="company" label={t("colCompany")} />
                          <th className="p-3 text-left">{t("colGstNo")}</th>
                          <th className="p-3 text-left">{t("colPrimaryContact")}</th><th className="p-3 text-left">{t("colCompanyPhone")}</th><th className="p-3 text-left">{t("colMobilePhone")}</th>
                          <th className="p-3 text-left">{t("colEmail")}</th><th className="p-3 text-left">{t("address")}</th><th className="p-3 text-left">{t("colPostalCode")}</th>
                          <th className="p-3 text-left">{t("colAccountDeptContact")}</th><th className="p-3 text-left">{t("colPaymentTerms")}</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y">
                        {filteredClients.map((c) => (
                          <tr key={c.id} className="hover:bg-blue-50/50 cursor-pointer" onClick={() => setDetailPanel({ type: "client", id: c.id })}>
                            <td className="p-3">{c.customer_no}</td>
                            <td className="p-3">{c.is_bu ? (c.bu_no || <span className="text-xs px-1.5 py-0.5 rounded bg-amber-100 text-amber-800">{t("colBu")}</span>) : "-"}</td>
                            <td className="p-3 font-medium">{c.company}</td><td className="p-3">{c.gst_no}</td>
                            <td className="p-3">{c.primary_contact}</td><td className="p-3">{c.company_phone}</td><td className="p-3">{c.mobile_phone}</td>
                            <td className="p-3 erp-cell-truncate" title={c.email}>{c.email}</td><td className="p-3 erp-cell-truncate" title={c.address}>{c.address}</td><td className="p-3">{c.postal_code}</td>
                            <td className="p-3">{c.account_dept_contact}</td><td className="p-3">{c.payment_terms}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}

              {page === "quotation" && can("quotation", "view") && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  <div className="p-4 border-b flex justify-between items-center gap-2 flex-wrap">
                    <h3 className="font-semibold">{t("nav_quotation")}</h3>
                    <div className="flex items-center gap-2 flex-wrap">
                      <ListImportExportBar module="quotation" />
                      {can("quotation", "add") && <button onClick={() => setQuotationModal({ mode: "add", data: emptyQuotation() })} className="px-3 py-2 text-sm rounded-lg bg-blue-600 text-white">{t("addQuotation")}</button>}
                    </div>
                  </div>
                  <ListToolbar t={t} mode="quotation" searchQ={quotationSearch} setSearchQ={setQuotationSearch} lang={lang} />
                  <div className="erp-list-scroll erp-list-scroll--compact">
                    <table className="w-full text-sm">
                      <thead className="bg-slate-50 text-slate-500">
                        <tr>
                          <SortableTh module="quotation" field="quotation_no" label={t("quotationNo")} />
                          <SortableTh module="quotation" field="company" label={t("clientGroup")} />
                          <th className="p-3 text-left">{t("quotationDate")}</th>
                          <th className="p-3 text-left">{t("validUntil")}</th><th className="p-3 text-right">{t("amount")}</th><th className="p-3 text-left">{t("currency")}</th>
                          <th className="p-3 text-left">{t("quotationStatus")}</th><th className="p-3 text-left">{t("colJobNo")}</th><th className="p-3 text-left">{t("description")}</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y">
                        {filteredQuotations.map((q) => (
                          <tr key={q.id} className="hover:bg-blue-50/50 cursor-pointer" onClick={() => setDetailPanel({ type: "quotation", id: q.id })}>
                            <td className="p-3 font-medium">{q.quotation_no}</td>
                            <td className="p-3">{q.company}</td>
                            <td className="p-3">{q.quotation_date || "-"}</td>
                            <td className="p-3">{q.valid_until || "-"}</td>
                            <td className="p-3 text-right tabular-nums">{money(q.amount)}</td>
                            <td className="p-3">{q.currency || "-"}</td>
                            <td className="p-3"><span className="px-2 py-0.5 rounded text-xs bg-slate-100">{q.status}</span></td>
                            <td className="p-3">{q.job_no || "-"}</td>
                            <td className="p-3 erp-cell-truncate" title={q.description || ""}>{q.description || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}

              {page === "job" && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  <div className="p-4 border-b">
                    <div className="flex justify-between items-center gap-2 flex-wrap">
                      <h3 className="font-semibold">{t("nav_job")}</h3>
                      <div className="flex items-center gap-2 flex-wrap">
                        <ListImportExportBar module="job" />
                        {can("job", "add") && <button onClick={() => setJobModal({ mode: "add", data: { ...emptyJob(), po_lines: [emptyPoLine()] } })} className="px-3 py-2 text-sm rounded-lg bg-blue-600 text-white">{t("addJob")}</button>}
                      </div>
                    </div>
                    <p className="text-xs text-slate-500 mt-2">{t("jobFlowHint")}</p>
                  </div>
                  <ListToolbar t={t} mode="job" clients={scopedClients} jobs={scopedJobs} viewMode={jobViewMode} setViewMode={setJobViewMode} clientFilter={jobClientFilter} setClientFilter={setJobClientFilter} statusFilter={jobStatusFilter} setStatusFilter={setJobStatusFilter} typeFilter={jobTypeFilter} setTypeFilter={setJobTypeFilter} searchQ={jobSearch} setSearchQ={setJobSearch} lang={lang} />
                  <div className="erp-list-scroll erp-list-scroll--toolbar erp-list-scroll--compact">
                    {jobViewMode === "grouped" ? (
                      groupRows(filteredJobs, (j) => j.company).map(([group, rows]) => (
                        <div key={group} className="mb-4">
                          <div className="px-4 py-2 bg-blue-50 border-y font-semibold text-blue-900">{group}</div>
                          <table className="w-full text-sm">
                            <thead className="bg-slate-50 text-slate-500">
                              <tr>
                                <SortableTh module="job" field="job_no" label={t("colJobNo")} />
                                <SortableTh module="job" field="company" label={t("clientGroup")} />
                                <th className="p-3 text-left">{t("description")}</th><th className="p-3 text-left">{t("type")}</th>
                                <th className="p-3 text-left">{t("quotationNo")}</th><th className="p-3 text-right">{t("amount")}</th><th className="p-3 text-left">{t("currency")}</th>
                                <th className="p-3 text-left">{t("colCustomerPo")}</th><th className="p-3 text-right">{t("colCustomerPoAmount")}</th><th className="p-3 text-left">{t("status")}</th><th className="p-3 text-left">{t("colStartDate")}</th>
                                <th className="p-3 text-left">{t("invoices")}</th><th className="p-3 text-left">{t("apBills")}</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y">{renderJobRows(rows)}</tbody>
                          </table>
                        </div>
                      ))
                    ) : (
                      <table className="w-full text-sm">
                        <thead className="bg-slate-50 text-slate-500">
                          <tr>
                                <SortableTh module="job" field="job_no" label={t("colJobNo")} />
                                <SortableTh module="job" field="company" label={t("clientGroup")} />
                                <th className="p-3 text-left">{t("description")}</th><th className="p-3 text-left">{t("type")}</th>
                                <th className="p-3 text-left">{t("quotationNo")}</th><th className="p-3 text-right">{t("amount")}</th><th className="p-3 text-left">{t("currency")}</th>
                                <th className="p-3 text-left">{t("colCustomerPo")}</th><th className="p-3 text-right">{t("colCustomerPoAmount")}</th><th className="p-3 text-left">{t("status")}</th><th className="p-3 text-left">{t("colStartDate")}</th>
                                <th className="p-3 text-left">{t("invoices")}</th><th className="p-3 text-left">{t("apBills")}</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y">{renderJobRows(filteredJobs)}</tbody>
                      </table>
                    )}
                  </div>
                </section>
              )}

              {page === "ongoing" && can("ongoing", "view") && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                    <div className="p-4 border-b flex flex-wrap items-start justify-between gap-3">
                      <div>
                        <h3 className="font-semibold text-lg">{t("ongoingListTitle")} - {activeRegion === "ALL" ? t("allRegions") : regionLabel(activeRegion, lang)}</h3>
                        <p className="text-xs text-slate-500 mt-1">{t("ongoingHint")}</p>
                      </div>
                      {can("ongoing", "export") && (
                        <ReportExportButtons onExcel={exportOngoingExcel} onPdf={exportOngoingPdf} />
                      )}
                    </div>
                  <div className="erp-list-scroll erp-list-scroll--ongoing" id="ongoing-report-area">
                    <table className="w-full text-xs min-w-[1100px]">
                      <thead>
                        <tr className="bg-[#dae3f3] text-slate-800">
                          <th className="p-2 text-left border border-slate-300 font-semibold whitespace-nowrap">{t("ongoingColJobNo")}</th>
                          <th className="p-2 text-left border border-slate-300 font-semibold whitespace-nowrap">{t("ongoingColNature")}<br /><span className="font-normal text-[10px]">{t("ongoingNatureHint")}</span></th>
                          <th className="p-2 text-left border border-slate-300 font-semibold">{t("ongoingColCustomer")}</th>
                          <th className="p-2 text-left border border-slate-300 font-semibold">{t("ongoingColProject")}</th>
                          <th className="p-2 text-left border border-slate-300 font-semibold">{t("ongoingColPo")}</th>
                          <th className="p-2 text-left border border-slate-300 font-semibold whitespace-nowrap">{t("ongoingColStart")}</th>
                          <th className="p-2 text-left border border-slate-300 font-semibold">{t("ongoingColCurrency")}</th>
                          <th className="p-2 text-right border border-slate-300 font-semibold">{t("ongoingColPoAmt")}</th>
                          <th className="p-2 text-right border border-slate-300 font-semibold">{t("ongoingColBilled")}</th>
                          <th className="p-2 text-right border border-slate-300 font-semibold">{t("ongoingColBalance")}</th>
                          <th className="p-2 text-right border border-slate-300 font-semibold whitespace-nowrap">{t("ongoingColBalanceUsd")}</th>
                          <th className="p-2 text-left border border-slate-300 font-semibold">{t("ongoingColRemarks")}</th>
                          <th className="p-2 text-center border border-slate-300 font-semibold whitespace-nowrap w-28">{t("edit")}</th>
                        </tr>
                      </thead>
                      <tbody>
                        {ongoingJobsList.length ? ongoingJobsList.map((j) => {
                          const isEditing = ongoingEditId === j.id;
                          const liveBilled = isEditing
                            ? (ongoingDraft.billedAmt === "" ? 0 : Number(ongoingDraft.billedAmt) || 0)
                            : j.billedAmt;
                          const livePoBalance = j.poAmt - liveBilled;
                          const liveBalanceUsd = convertCurrency(livePoBalance, (j.arCurrency || "USD").split("/")[0], "USD");
                          return (
                          <tr key={j.id} className={"hover:bg-blue-50/40 " + (isEditing ? "" : "cursor-pointer")} onClick={() => { if (!isEditing) setDetailPanel({ type: "job", id: j.id }); }}>
                            <td className="p-2 border border-slate-200 font-medium whitespace-nowrap">{j.job_no}</td>
                            <td className="p-2 border border-slate-200">{j.nature}</td>
                            <td className="p-2 border border-slate-200">{j.company}</td>
                            <td className="p-2 border border-slate-200">{j.description || ""}</td>
                            <td className="p-2 border border-slate-200 whitespace-nowrap">{j.customer_po || ""}</td>
                            <td className="p-2 border border-slate-200 whitespace-nowrap">{j.start_date || ""}</td>
                            <td className="p-2 border border-slate-200">{j.arCurrency || ""}</td>
                            <td className="p-2 border border-slate-200 text-right tabular-nums">{j.poAmtLabel || (j.poAmt ? money(j.poAmt) : "")}</td>
                            <td className="p-2 border border-slate-200 text-right tabular-nums" onClick={(e) => e.stopPropagation()}>
                              {isEditing ? (
                                <input
                                  type="number"
                                  step="0.01"
                                  min="0"
                                  className="w-full min-w-[5.5rem] text-right rounded border border-blue-300 px-2 py-1 text-xs tabular-nums focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                                  value={ongoingDraft.billedAmt}
                                  placeholder="0"
                                  onChange={(e) => setOngoingDraft((d) => ({ ...d, billedAmt: e.target.value }))}
                                />
                              ) : (
                                j.billedAmt ? money(j.billedAmt) : ""
                              )}
                            </td>
                            <td className="p-2 border border-slate-200 text-right tabular-nums">{money(livePoBalance)}</td>
                            <td className="p-2 border border-slate-200 text-right tabular-nums">{money(liveBalanceUsd)}</td>
                            <td className="p-2 border border-slate-200" onClick={(e) => e.stopPropagation()}>
                              {isEditing ? (
                                <input
                                  type="text"
                                  className="w-full min-w-[8rem] rounded border border-blue-300 px-2 py-1 text-xs focus:ring-2 focus:ring-blue-500 focus:border-blue-500"
                                  value={ongoingDraft.remarks}
                                  placeholder=""
                                  onChange={(e) => setOngoingDraft((d) => ({ ...d, remarks: e.target.value }))}
                                />
                              ) : (
                                j.remarks || ""
                              )}
                            </td>
                            <td className="p-2 border border-slate-200 text-center whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
                              {can("job", "edit") && (
                                isEditing ? (
                                  <div className="flex flex-col gap-1 items-center">
                                    <button type="button" onClick={() => saveOngoingRow(j.id)} className="px-2 py-1 text-[10px] rounded bg-blue-600 text-white hover:bg-blue-700 w-full">{t("save")}</button>
                                    <button type="button" onClick={cancelOngoingEdit} className="px-2 py-1 text-[10px] rounded border border-slate-300 hover:bg-slate-50 w-full">{t("cancel")}</button>
                                  </div>
                                ) : (
                                  <button
                                    type="button"
                                    disabled={ongoingEditId != null && ongoingEditId !== j.id}
                                    onClick={() => startOngoingEdit(j)}
                                    className={"px-2 py-1 text-[10px] rounded " + (ongoingEditId != null && ongoingEditId !== j.id ? "bg-slate-100 text-slate-400 cursor-not-allowed" : "bg-amber-100 text-amber-800 hover:bg-amber-200")}
                                  >
                                    {t("edit")}
                                  </button>
                                )
                              )}
                            </td>
                          </tr>
                          );
                        }) : (
                          <tr><td colSpan={13} className="p-12 text-center text-slate-400 border border-slate-200">{t("noJobs")}</td></tr>
                        )}
                        {ongoingJobsList.length > 0 && (
                          <>
                            <tr><td colSpan={13} className="h-10"></td></tr>
                            <tr>
                              <td colSpan={7}></td>
                              <td colSpan={5} className="p-0 align-top">
                                <table className="w-full text-xs border-collapse">
                                  <tbody>
                                    <tr>
                                      <td className="p-2 border border-slate-300 font-semibold bg-slate-50">{t("ongoingTotalPoUsd")}</td>
                                      <td className="p-2 border border-slate-300 text-right tabular-nums font-semibold bg-slate-50">{money(ongoingTotals.totalPoUsd)}</td>
                                    </tr>
                                    <tr>
                                      <td className="p-2 border border-slate-300">{t("ongoingTotalMaintUsd")}</td>
                                      <td className="p-2 border border-slate-300 text-right tabular-nums">{money(ongoingTotals.Maintenance)}</td>
                                    </tr>
                                    <tr>
                                      <td className="p-2 border border-slate-300">{t("ongoingTotalServiceUsd")}</td>
                                      <td className="p-2 border border-slate-300 text-right tabular-nums">{money(ongoingTotals.Service)}</td>
                                    </tr>
                                    <tr>
                                      <td className="p-2 border border-slate-300">{t("ongoingTotalProjectUsd")}</td>
                                      <td className="p-2 border border-slate-300 text-right tabular-nums">{money(ongoingTotals.Project)}</td>
                                    </tr>
                                    <tr>
                                      <td className="p-2 border border-slate-300">{t("ongoingTotalTradeUsd")}</td>
                                      <td className="p-2 border border-slate-300 text-right tabular-nums">{money(ongoingTotals.Trade)}</td>
                                    </tr>
                                  </tbody>
                                </table>
                              </td>
                            </tr>
                          </>
                        )}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}

              {page === "vendors" && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  <div className="p-4 border-b flex justify-between items-center gap-2 flex-wrap">
                    <h3 className="font-semibold">{t("nav_vendors")}</h3>
                    <div className="flex items-center gap-2 flex-wrap">
                      <ListImportExportBar module="vendors" />
                      {can("vendors", "add") && <button onClick={() => setVendorModal({ mode: "add", data: emptyVendor() })} className="px-3 py-2 text-sm rounded-lg bg-blue-600 text-white">{t("addVendor")}</button>}
                    </div>
                  </div>
                  <ListToolbar t={t} mode="vendors" searchQ={vendorsSearch} setSearchQ={setVendorsSearch} lang={lang} />
                  <div className="erp-list-scroll erp-list-scroll--compact">
                    <table className="w-full text-sm">
                      <thead className="bg-slate-50 text-slate-500">
                        <tr><SortableTh module="vendors" field="vendor_no" label={t("vendorNo")} /><SortableTh module="vendors" field="name" label={t("colName")} /><th className="p-3 text-left">{t("phoneNo")}</th><th className="p-3 text-left">{t("address")}</th><th className="p-3 text-left">{t("colEmail")}</th><th className="p-3 text-left">{t("colFinanceContact")}</th><th className="p-3 text-left">{t("colFinanceEmail")}</th><th className="p-3 text-left">{t("colBank")}</th><th className="p-3 text-left">{t("colSwift")}</th><th className="p-3 text-left">{t("colContact")}</th></tr>
                      </thead>
                      <tbody className="divide-y">
                        {filteredVendors.map((v) => (
                          <tr key={v.id} className="hover:bg-blue-50/50 cursor-pointer" onClick={() => setDetailPanel({ type: "vendor", id: v.id })}>
                            <td className="p-3 whitespace-nowrap">{v.vendor_no || "-"}</td><td className="p-3 font-medium">{v.name}</td><td className="p-3 whitespace-nowrap">{v.phone || "-"}</td><td className="p-3 erp-cell-truncate" title={v.address || ""}>{v.address || "-"}</td><td className="p-3 erp-cell-truncate" title={v.email}>{v.email}</td><td className="p-3">{v.finance_contact || "-"}</td><td className="p-3 erp-cell-truncate" title={v.finance_email || ""}>{v.finance_email || "-"}</td><td className="p-3 erp-cell-truncate" title={v.bank}>{v.bank}</td><td className="p-3">{v.swift_code || "-"}</td><td className="p-3">{v.contact || "-"}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              )}

              {page === "ar" && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  <div className="p-4 border-b flex justify-between items-center gap-2 flex-wrap">
                    <h3 className="font-semibold">{t("nav_ar")}</h3>
                    <div className="flex items-center gap-2 flex-wrap">
                      <ListImportExportBar module="ar" />
                      {can("ar", "add") && <button onClick={() => setArModal({ mode: "add", data: { job_id: "", job_no: "", customer: "", customer_po: "", job_completed: false, invoice_no: "", invoice_currency: "USD", invoice_amt: "", exchange_rate: "", base_amount: "", override_reason: "", invoice_date: "", due_date: "", payment_received_date: "", payment_status: "Awaiting Payment", payment_stage: "" } })} className="px-3 py-2 text-sm rounded-lg bg-blue-600 text-white">{t("createAr")}</button>}
                    </div>
                  </div>
                  <ListToolbar t={t} mode="ar" clients={scopedClients} jobs={scopedJobs} viewMode={arViewMode} setViewMode={setArViewMode} clientFilter={arClientFilter} setClientFilter={setArClientFilter} searchQ={arSearch} setSearchQ={setArSearch} lang={lang} />
                  <div className="erp-list-scroll erp-list-scroll--toolbar erp-list-scroll--compact">
                    {arViewMode === "grouped" ? (
                      groupRows(filteredAr, (r) => r.customer).map(([group, rows]) => (
                        <div key={group} className="mb-4">
                          <div className="px-4 py-2 bg-blue-50 border-y font-semibold text-blue-900">{group}</div>
                          <table className="w-full text-sm">
                            <thead className="bg-slate-50 text-slate-500">
                              <tr>
                                <th className="p-3 text-left">{t("colJobNo")}</th><th className="p-3 text-left">{t("colCustomer")}</th><th className="p-3 text-left">{t("colCustomerPo")}</th><th className="p-3 text-left">{t("jobStatus")}</th>
                                <SortableTh module="ar" field="invoice_no" label={t("colInvoiceNo")} /><th className="p-3 text-left">{t("paymentStage")}</th><th className="p-3 text-left">{t("colInvoiceCurrency")}</th><th className="p-3 text-right">{t("colInvoiceAmt")}</th>
                                <th className="p-3 text-right">{amtInLabel}</th><th className="p-3 text-left">{t("colInvoiceDate")}</th><th className="p-3 text-left">{t("colDueDate")}</th>
                                <th className="p-3 text-right">{t("colOverdueDays")}</th><th className="p-3 text-right">{t("colDays")}</th><th className="p-3 text-left">{t("paymentReceived")}</th><th className="p-3 text-left">{t("paymentStatus")}</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y">{renderArRows(rows)}</tbody>
                          </table>
                        </div>
                      ))
                    ) : (
                      <table className="w-full text-sm">
                        <thead className="bg-slate-50 text-slate-500">
                          <tr>
                            <th className="p-3 text-left">{t("colJobNo")}</th><th className="p-3 text-left">{t("colCustomer")}</th><th className="p-3 text-left">{t("colCustomerPo")}</th><th className="p-3 text-left">{t("jobStatus")}</th>
                            <SortableTh module="ar" field="invoice_no" label={t("colInvoiceNo")} /><th className="p-3 text-left">{t("paymentStage")}</th><th className="p-3 text-left">{t("colInvoiceCurrency")}</th><th className="p-3 text-right">{t("colInvoiceAmt")}</th>
                            <th className="p-3 text-right">{amtInLabel}</th><th className="p-3 text-left">{t("colInvoiceDate")}</th><th className="p-3 text-left">{t("colDueDate")}</th>
                            <th className="p-3 text-right">{t("colOverdueDays")}</th><th className="p-3 text-right">{t("colDays")}</th><th className="p-3 text-left">{t("paymentReceived")}</th><th className="p-3 text-left">{t("paymentStatus")}</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y">{renderArRows(filteredAr)}</tbody>
                      </table>
                    )}
                  </div>
                </section>
              )}

              {page === "ap" && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  <div className="p-4 border-b flex justify-between items-center gap-2 flex-wrap">
                    <h3 className="font-semibold">{t("nav_ap")}</h3>
                    <div className="flex items-center gap-2 flex-wrap">
                      <ListImportExportBar module="ap" />
                      {can("ap", "add") && <button onClick={() => setApModal({ mode: "add", data: emptyApForm() })} className="px-3 py-2 text-sm rounded-lg bg-blue-600 text-white">{t("createAp")}</button>}
                    </div>
                  </div>
                  <ListToolbar t={t} mode="ap" clients={scopedClients} jobs={scopedJobs} viewMode={apViewMode} setViewMode={setApViewMode} clientFilter={apClientFilter} setClientFilter={setApClientFilter} jobFilter={apJobFilter} setJobFilter={setApJobFilter} searchQ={apSearch} setSearchQ={setApSearch} lang={lang} />
                  <div className="erp-list-scroll erp-list-scroll--toolbar erp-list-scroll--compact">
                    {apViewMode === "grouped" ? (
                      groupRows(filteredAp, (r) => r.company_name).map(([group, rows]) => (
                        <div key={group} className="mb-4">
                          <div className="px-4 py-2 bg-violet-50 border-y font-semibold text-violet-900">{group}</div>
                          <table className="w-full text-sm">
                            <thead className="bg-slate-50 text-slate-500">
                              <tr>
                                <th className="p-3 text-left">{t("colJobNo")}</th><th className="p-3 text-left">{t("payeeType")}</th><th className="p-3 text-left">{t("colCompanyName")}</th><SortableTh module="ap" field="invoice_no" label={t("colInvoiceNo")} /><th className="p-3 text-left">{t("paymentStage")}</th><th className="p-3 text-left">{t("colInvoiceDate")}</th><th className="p-3 text-left">{t("colInvoiceReceivedDate")}</th>
                                <th className="p-3 text-left">{t("colDueDate")}</th><th className="p-3 text-left">{t("currency")}</th><th className="p-3 text-right">{t("amount")}</th><th className="p-3 text-right">{amtInLabel}</th>
                            <th className="p-3 text-left whitespace-nowrap">{t("paymentStatus")}</th><th className="p-3 text-left">{t("colPayDate")}</th><th className="p-3 text-left">{t("colRemarks")}</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y">{renderApRows(rows)}</tbody>
                          </table>
                        </div>
                      ))
                    ) : (
                      <table className="w-full text-sm">
                        <thead className="bg-slate-50 text-slate-500">
                          <tr>
                            <th className="p-3 text-left">{t("colJobNo")}</th><th className="p-3 text-left">{t("payeeType")}</th><th className="p-3 text-left">{t("colCompanyName")}</th><SortableTh module="ap" field="invoice_no" label={t("colInvoiceNo")} /><th className="p-3 text-left">{t("paymentStage")}</th><th className="p-3 text-left">{t("colInvoiceDate")}</th><th className="p-3 text-left">{t("colInvoiceReceivedDate")}</th>
                            <th className="p-3 text-left">{t("colDueDate")}</th><th className="p-3 text-left">{t("currency")}</th><th className="p-3 text-right">{t("amount")}</th><th className="p-3 text-right">{amtInLabel}</th>
                            <th className="p-3 text-left whitespace-nowrap">{t("paymentStatus")}</th><th className="p-3 text-left">{t("colPayDate")}</th><th className="p-3 text-left">{t("colRemarks")}</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y">{renderApRows(filteredAp)}</tbody>
                      </table>
                    )}
                  </div>
                </section>
              )}

              {page === "monthly_report" && can("monthly_report", "view") && (
                <div className="space-y-4">
                  <section className="bg-white border border-slate-200 rounded-2xl shadow-sm overflow-hidden">
                    <div className="p-4 border-b flex flex-wrap items-start justify-between gap-3">
                      <div className="min-w-0 flex-1">
                        <h3 className="font-semibold text-lg">{t("monthlyReportChartTitle").replace("{region}", activeRegion === "ALL" ? t("allRegions") : regionLabel(activeRegion, lang)).replace("{year}", monthlyReportYear)}</h3>
                        <p className="text-xs text-slate-500 mt-0.5">{t("monthlyReportHint")} · {regionListCurrency}</p>
                        <p className="text-[10px] text-slate-400 mt-1">{t("monthlyStorageHint")}</p>
                      </div>
                      <div className="flex flex-nowrap items-center gap-2 shrink-0">
                        <span className="text-xs font-medium text-slate-600 whitespace-nowrap">{t("monthlyReportYear")}</span>
                        <Input type="number" min="2020" max="2099" value={monthlyReportYear} onChange={(e) => setMonthlyReportYear(e.target.value)} className="h-9 w-24 shrink-0" />
                        {can("monthly_report", "edit") && (
                          <button type="button" onClick={syncMonthlyReportAll} className="px-3 py-1.5 text-xs rounded-lg bg-blue-600 text-white hover:bg-blue-700 whitespace-nowrap font-medium">{t("monthlySyncReport")}</button>
                        )}
                        {isRoot() && (
                          <button type="button" onClick={clearAllMonthlyReportData} className="px-3 py-1.5 text-xs rounded-lg border border-red-300 bg-red-50 text-red-700 hover:bg-red-100 whitespace-nowrap font-medium">{t("monthlyClearAll")}</button>
                        )}
                        {can("monthly_report", "export") && (
                          <ReportExportButtons
                            onExcel={monthlyReportTab === "report" ? () => exportModuleExcel("monthly_report", "data") : monthlyReportTab === "po_received" ? () => exportMonthlyDetailExcel("po") : () => exportMonthlyDetailExcel("ar")}
                            onPdf={exportMonthlyReportPdfClick}
                          />
                        )}
                      </div>
                    </div>
                    <div className="px-4 pt-3 border-b flex flex-wrap gap-1">
                      <button type="button" onClick={() => setMonthlyReportTab("report")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (monthlyReportTab === "report" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500 hover:text-slate-700")}>{t("monthlyTabReport")}</button>
                      <button type="button" onClick={() => setMonthlyReportTab("po_received")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (monthlyReportTab === "po_received" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500 hover:text-slate-700")}>{t("monthlyTabPoReceived")} Y{monthlyReportYear.slice(-2)}</button>
                      <button type="button" onClick={() => setMonthlyReportTab("ar_invoice")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (monthlyReportTab === "ar_invoice" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500 hover:text-slate-700")}>{t("monthlyTabArInvoice")} Y{monthlyReportYear.slice(-2)}</button>
                    </div>

                    {monthlyReportTab === "report" && (
                      <div className="erp-list-scroll erp-list-scroll--ongoing p-4">
                        <p className="text-xs text-slate-500 mb-3">{t("monthlyReportSyncHint")}</p>
                        <table className="w-full text-xs min-w-[960px]">
                          <thead>
                            <tr className="bg-[#dae3f3] text-slate-800">
                              <th className="p-2 text-left border border-slate-300 font-semibold w-48"></th>
                              {monthlyReportSeries.map((m) => (
                                <th key={m.key} className="p-2 text-center border border-slate-300 font-semibold whitespace-nowrap min-w-[5.5rem]">{m.header}</th>
                              ))}
                            </tr>
                          </thead>
                          <tbody>
                            <tr>
                              <td className="p-2 border border-slate-200"></td>
                              {monthlyReportSeries.map((m) => (
                                <td key={m.key} className="p-2 border border-slate-200 text-center font-medium">{m.currency}</td>
                              ))}
                            </tr>
                            <tr>
                              <td className="p-2 border border-slate-200 font-medium whitespace-nowrap">{t("monthlyPoReceived")}</td>
                              {monthlyReportSeries.map((m) => (
                                <td key={m.key} className="p-2 border border-slate-200 text-right tabular-nums">{m.poReceived ? money(m.poReceived) : ""}</td>
                              ))}
                            </tr>
                            <tr>
                              <td className="p-2 border border-slate-200 font-medium whitespace-nowrap">{t("monthlyInvoiceIssued")}</td>
                              {monthlyReportSeries.map((m) => (
                                <td key={m.key} className="p-2 border border-slate-200 text-right tabular-nums">{m.invoiceIssued ? money(m.invoiceIssued) : ""}</td>
                              ))}
                            </tr>
                            <tr>
                              <td className="p-2 border border-slate-200 font-medium whitespace-nowrap text-red-600">{t("monthlyArExpected")}</td>
                              {monthlyReportSeries.map((m) => (
                                <td key={m.key} className="p-2 border border-slate-200 text-right tabular-nums text-red-600 font-medium">{m.arExpected ? money(m.arExpected) : ""}</td>
                              ))}
                            </tr>
                          </tbody>
                        </table>
                        <p className="text-xs text-red-600 mt-3">{t("monthlyArExpectedHint")}</p>
                      </div>
                    )}

                    {monthlyReportTab === "po_received" && (
                      <div className="p-4 space-y-3">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <p className="text-xs text-slate-500">{t("monthlyPoDetailHint").replace("{currency}", regionListCurrency)}</p>
                          {can("monthly_report", "edit") && (
                            <button type="button" onClick={syncMonthlyPoFromJobs} className="px-3 py-1.5 text-xs rounded-lg bg-blue-600 text-white hover:bg-blue-700">{t("monthlySyncFromJobs")}</button>
                          )}
                        </div>
                        <MonthlyDetailGrid grid={monthlyPoGrid} baseCurrency={regionListCurrency} canEdit={can("monthly_report", "edit")} moneyFn={money} t={t} onOtherCurrencyChange={updateMonthlyPoOtherCurrency} />
                      </div>
                    )}

                    {monthlyReportTab === "ar_invoice" && (
                      <div className="p-4 space-y-3">
                        <div className="flex flex-wrap items-center justify-between gap-2">
                          <p className="text-xs text-slate-500">{t("monthlyArDetailHint").replace("{currency}", regionListCurrency)}</p>
                          {can("monthly_report", "edit") && (
                            <button type="button" onClick={syncMonthlyArFromInvoices} className="px-3 py-1.5 text-xs rounded-lg bg-blue-600 text-white hover:bg-blue-700">{t("monthlySyncFromAr")}</button>
                          )}
                        </div>
                        <MonthlyDetailGrid grid={monthlyArGrid} baseCurrency={regionListCurrency} canEdit={can("monthly_report", "edit")} moneyFn={money} t={t} onOtherCurrencyChange={updateMonthlyArOtherCurrency} />
                      </div>
                    )}
                  </section>

                  {monthlyReportTab === "report" && (
                    <section className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4">
                      <h4 className="font-semibold text-center text-slate-700 mb-4">{t("monthlyReportChartTitle").replace("{region}", activeRegion === "ALL" ? t("allRegions") : regionLabel(activeRegion, lang)).replace("{year}", monthlyReportYear)}</h4>
                      <MonthlyReportBarChart series={monthlyReportSeries} maxBar={monthlyReportMaxBar} t={t} moneyFn={money} />
                    </section>
                  )}
                </div>
              )}

              {page === "settings" && (
                <section className="bg-white border border-slate-200 rounded-xl shadow-sm overflow-hidden">
                  <div className="px-6 pt-5 border-b flex flex-wrap gap-2">
                    <button type="button" onClick={() => setSettingsTab("general")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (settingsTab === "general" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500")}>{t("settingsGeneral")}</button>
                    <button type="button" onClick={() => setSettingsTab("guide")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (settingsTab === "guide" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500")}>{t("settingsGuide")}</button>
                    <button type="button" onClick={() => setSettingsTab("backup")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (settingsTab === "backup" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500")}>{t("settingsBackup")}</button>
                    {isRoot() && <button type="button" onClick={() => setSettingsTab("management")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (settingsTab === "management" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500")}>{t("settingsManagement")}</button>}
                    <button type="button" onClick={() => setSettingsTab("audit")} className={"px-3 py-2 text-sm rounded-t-lg border-b-2 " + (settingsTab === "audit" ? "border-blue-600 text-blue-700 font-semibold" : "border-transparent text-slate-500")}>{t("settingsAudit")}</button>
                  </div>

                  {settingsTab === "general" && (
                    <div className="p-6 max-w-xl space-y-4">
                      <h3 className="font-semibold text-lg">{t("settingsTitle")}</h3>
                      <Field label={t("language")}>
                        <Select value={lang} onChange={(e) => changeLang(e.target.value)}>
                          <option value="en">{t("langEn")}</option>
                          <option value="zh_TW">{t("langZhTw")}</option>
                        </Select>
                      </Field>
                      <p className="text-sm text-slate-500">{t("settingsLangHint")}</p>
                      <Field label={t("timezone")}>
                        <Select value={worldTimeZone} onChange={(e) => setWorldTimeZone(e.target.value)}>
                          {WORLD_TIMEZONES.map((tz) => <option key={tz.id} value={tz.id}>{i18nLang(lang) === "zh_TW" ? tz.labelZh : tz.labelEn}</option>)}
                        </Select>
                      </Field>
                      <p className="text-sm text-slate-500">{t("worldTimeHint")}</p>
                      <p className="text-sm font-mono bg-slate-50 border rounded-lg px-3 py-2">{timezoneLabel(worldTimeZone, lang)} · {formatWorldTime(worldClock, worldTimeZone, lang)}</p>
                      <p className="text-sm text-slate-500">{t("regionDataHint")}</p>
                      <p className="text-sm text-amber-700 bg-amber-50 border border-amber-200 rounded-lg px-3 py-2">{t("noDatabaseNote")}</p>
                      <div className="pt-4 mt-4 border-t border-slate-200 space-y-3">
                        <h4 className="font-semibold text-slate-800">{t("myAccountTitle")}</h4>
                        <p className="text-sm text-slate-500">{getCurrentUser() ? (getCurrentUser().name + " · " + getCurrentUser().email) : "—"}</p>
                        <form onSubmit={changeMyPassword} className="space-y-3 max-w-md">
                          <Field label={t("currentPassword")}><Input required type="password" autoComplete="current-password" value={pwForm.current} onChange={(e) => setPwForm({ ...pwForm, current: e.target.value })} /></Field>
                          <Field label={t("newPassword")}><Input required type="password" autoComplete="new-password" value={pwForm.next} onChange={(e) => setPwForm({ ...pwForm, next: e.target.value })} /></Field>
                          <Field label={t("confirmPassword")}><Input required type="password" autoComplete="new-password" value={pwForm.confirm} onChange={(e) => setPwForm({ ...pwForm, confirm: e.target.value })} /></Field>
                          <button type="submit" className="px-4 py-2 text-sm rounded-lg bg-blue-600 text-white hover:bg-blue-700">{t("changePassword")}</button>
                        </form>
                        {isRoot() && <p className="text-xs text-slate-500">{t("mgmtHint")}</p>}
                      </div>
                    </div>
                  )}

                  {settingsTab === "guide" && (
                    <div className="p-6 max-w-3xl space-y-4">
                      <h3 className="font-semibold text-lg">{t("settingsGuide")}</h3>
                      <p className="text-sm text-slate-500">{t("userGuideHint")}</p>
                      <button type="button" onClick={() => window.open(airlinkAsset("user-guide.html"), "_blank")} className="px-4 py-2 text-sm rounded-lg bg-blue-600 text-white hover:bg-blue-700">{t("openUserGuide")}</button>
                      <p className="text-xs text-slate-400">{i18nLang(lang) === "zh_TW" ? "若新分頁無法開啟，請向管理員索取 user-guide.html 檔案。" : "If the new tab fails, ask your admin for the user-guide.html file."}</p>
                    </div>
                  )}

                  {settingsTab === "backup" && (
                    <div className="p-6 max-w-xl space-y-4">
                      <h3 className="font-semibold text-lg">{t("backupTitle")}</h3>
                      <p className="text-sm text-slate-500">{t("backupHint")}</p>
                      <div className="flex flex-wrap gap-2">
                        <button type="button" onClick={downloadErpBackup} className="px-4 py-2 text-sm rounded-lg bg-emerald-600 text-white hover:bg-emerald-700">{t("backupDownload")}</button>
                        {isRoot() && (
                          <>
                            <input id="erp-restore-backup" type="file" accept=".json,application/json" className="hidden" onChange={(e) => { handleRestoreBackup(e.target.files && e.target.files[0]); e.target.value = ""; }} />
                            <button type="button" onClick={() => { const el = document.getElementById("erp-restore-backup"); if (el) el.click(); }} className="px-4 py-2 text-sm rounded-lg border border-amber-400 bg-amber-50 text-amber-900 hover:bg-amber-100">{t("restoreUpload")}</button>
                          </>
                        )}
                      </div>
                      {!isRoot() && <p className="text-xs text-slate-500">{t("rootOnly")}</p>}

                      <div className="mt-6 pt-6 border-t border-slate-200 space-y-3">
                        <h4 className="font-semibold text-slate-800">{t("cloudSyncTitle")}</h4>
                        <p className="text-sm text-slate-500">{t("cloudSyncHint")}</p>
                        {cloudDbStatus && (
                          <p className={"text-xs font-medium " + (cloudDbStatus.database && cloudDbStatus.database.ok ? "text-emerald-700" : "text-amber-700")}>
                            API: {cloudDbStatus.status || "—"}
                            {cloudDbStatus.database ? " · DB: " + (cloudDbStatus.database.ok ? (cloudDbStatus.database.provider || "ok") : "offline") : ""}
                            {cloudDbStatus.cloud_sync_configured === false ? " · " + t("cloudNotConfigured") : ""}
                          </p>
                        )}
                        {isRoot() && (
                          <>
                            <div className="max-w-md">
                              <label className="block text-xs text-slate-500 mb-1">{t("cloudSyncKey")}</label>
                              <input
                                type="password"
                                className="w-full h-9 rounded-lg border border-slate-300 px-3 text-sm"
                                placeholder={t("cloudSyncKeyPlaceholder")}
                                value={erpSyncKey}
                                onChange={(e) => setErpSyncKey(e.target.value)}
                              />
                            </div>
                            <div className="flex flex-wrap gap-2">
                              <button type="button" onClick={pushCloudBackup} className="px-4 py-2 text-sm rounded-lg bg-indigo-600 text-white hover:bg-indigo-700">{t("cloudPush")}</button>
                              <button type="button" onClick={pullCloudBackup} className="px-4 py-2 text-sm rounded-lg border border-indigo-300 bg-indigo-50 text-indigo-900 hover:bg-indigo-100">{t("cloudPull")}</button>
                            </div>
                          </>
                        )}
                      </div>

                      <div className="mt-6 pt-6 border-t border-slate-200 space-y-3">
                        <h4 className="font-semibold text-slate-800">{t("liveSyncTitle")}</h4>
                        {cloudOnlyMode && (
                          <p className="text-sm text-emerald-700 font-medium">{t("cloudOnlyMode")} — {t("cloudOnlyHint")}</p>
                        )}
                        <p className="text-sm text-slate-500">{t("liveSyncAutoHint")}</p>
                        {liveSyncEnabled && (
                          <p className={"text-xs font-medium " + (liveSyncStatus === "error" ? "text-red-600" : liveSyncStatus === "syncing" ? "text-blue-600" : "text-emerald-700")}>
                            {liveSyncStatus === "syncing" ? t("liveSyncSyncing") : liveSyncStatus === "error" ? t("liveSyncError") : (canLiveSyncPush() ? t("liveSyncOk") : t("liveSyncReadOnly"))}
                            {liveSyncLastAt ? " · " + t("liveSyncLast") + " " + liveSyncLastAt.toLocaleTimeString() : ""}
                          </p>
                        )}
                      </div>

                      {isRoot() && (
                        <div className="mt-6 pt-6 border-t border-slate-200 space-y-3">
                          <h4 className="font-semibold text-slate-800">{t("monthlyClearAll")}</h4>
                          <p className="text-sm text-slate-500">{t("monthlyClearAllConfirm")}</p>
                          <button type="button" onClick={clearAllMonthlyReportData} className="px-4 py-2 text-sm rounded-lg border border-red-300 bg-red-50 text-red-700 hover:bg-red-100">{t("monthlyClearAll")}</button>
                        </div>
                      )}
                    </div>
                  )}

                  {settingsTab === "management" && isRoot() && (
                    <div className="p-6 space-y-4">
                      <div className="flex justify-between items-start gap-3 flex-wrap">
                        <div>
                          <h3 className="font-semibold text-lg">{t("mgmtTitle")}</h3>
                          <p className="text-sm text-slate-500 mt-1">{t("mgmtHint")}</p>
                        </div>
                        <button type="button" onClick={() => setUserModal({ mode: "add", data: emptyUserForm() })} className="px-3 py-2 text-sm rounded-lg bg-blue-600 text-white">{t("addUser")}</button>
                      </div>
                      <div className="erp-list-scroll erp-list-scroll--settings erp-list-scroll--compact border rounded-lg">
                        <table className="w-full text-sm">
                          <thead className="bg-slate-50 text-slate-500">
                            <tr>
                              <th className="p-3 text-left">{t("userName")}</th>
                              <th className="p-3 text-left">{t("userLoginId")}</th>
                              <th className="p-3 text-left">{t("userRole")}</th>
                              <th className="p-3 text-left">{t("userActive")}</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y">
                            {users.map((u) => (
                              <tr key={u.id} className="hover:bg-slate-50 cursor-pointer" onClick={() => setUserModal({ mode: "edit", id: u.id, data: { name: u.name, email: u.email, password: u.password || "", role: u.role, is_active: u.is_active, permissions: u.permissions || emptyPermissions(), allowed_regions: u.allowed_regions || ["HK"] } })}>
                                <td className="p-3 font-medium">{u.name}</td>
                                <td className="p-3">{u.email}</td>
                                <td className="p-3"><span className={"px-2 py-0.5 rounded text-xs " + (u.role === "root" ? "bg-violet-100 text-violet-800" : "bg-slate-100")}>{u.role === "root" ? t("roleRoot") : t("roleUser")}</span></td>
                                <td className="p-3">{u.is_active ? t("yes") : t("no")}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}

                  {settingsTab === "audit" && (
                    <div className="p-6 space-y-4">
                      <div>
                        <h3 className="font-semibold text-lg">{t("auditTitle")}</h3>
                        <p className="text-sm text-slate-500 mt-1">{t("auditHint")}</p>
                      </div>
                      <div className="grid grid-cols-1 md:grid-cols-3 xl:grid-cols-6 gap-3 p-4 bg-slate-50 rounded-lg border">
                        <Field label={t("auditDateFrom")}><Input type="date" value={auditFilters.dateFrom} onChange={(e) => setAuditFilters({ ...auditFilters, dateFrom: e.target.value })} /></Field>
                        <Field label={t("auditDateTo")}><Input type="date" value={auditFilters.dateTo} onChange={(e) => setAuditFilters({ ...auditFilters, dateTo: e.target.value })} /></Field>
                        <Field label={t("auditUser")}>
                          <Select value={auditFilters.userId} onChange={(e) => setAuditFilters({ ...auditFilters, userId: e.target.value })}>
                            <option value="all">{t("allUsers")}</option>
                            {users.map((u) => <option key={u.id} value={String(u.id)}>{u.name}</option>)}
                          </Select>
                        </Field>
                        <Field label={t("auditModule")}>
                          <Select value={auditFilters.module} onChange={(e) => setAuditFilters({ ...auditFilters, module: e.target.value })}>
                            <option value="all">{t("allModules")}</option>
                            {ERP_MODULES.map((m) => <option key={m} value={m}>{moduleLabel(m, t)}</option>)}
                          </Select>
                        </Field>
                        <Field label={t("auditAction")}>
                          <Select value={auditFilters.action} onChange={(e) => setAuditFilters({ ...auditFilters, action: e.target.value })}>
                            <option value="all">{t("allActions")}</option>
                            <option value="create">{t("actionCreate")}</option>
                            <option value="update">{t("actionUpdate")}</option>
                            <option value="delete">{t("actionDelete")}</option>
                            <option value="mark_paid">{t("actionMarkPaid")}</option>
                            <option value="user_create">{t("actionUserCreate")}</option>
                            <option value="user_update">{t("actionUserUpdate")}</option>
                            <option value="user_delete">{t("actionUserDelete")}</option>
                            <option value="export">{t("actionExport")}</option>
                            <option value="import">{t("actionImport")}</option>
                          </Select>
                        </Field>
                        <Field label={t("auditSearch")}><Input value={auditFilters.q} onChange={(e) => setAuditFilters({ ...auditFilters, q: e.target.value })} placeholder="..." /></Field>
                      </div>
                      <div className="flex gap-2">
                        <button type="button" onClick={() => setAuditFilters({ dateFrom: "", dateTo: "", userId: "all", module: "all", action: "all", q: "" })} className="px-3 py-2 text-sm rounded-lg border">{t("auditReset")}</button>
                      </div>
                      <div className="erp-list-scroll erp-list-scroll--settings border rounded-lg">
                        <table className="w-full text-sm">
                          <thead className="bg-slate-50 text-slate-500">
                            <tr>
                              <th className="p-3 text-left">{t("auditWhen")}</th>
                              <th className="p-3 text-left">{t("auditUser")}</th>
                              <th className="p-3 text-left">{t("auditModule")}</th>
                              <th className="p-3 text-left">{t("auditAction")}</th>
                              <th className="p-3 text-left">{t("auditEntity")}</th>
                              <th className="p-3 text-left">{t("auditDetail")}</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y">
                            {filteredAuditLogs.length === 0 ? (
                              <tr><td colSpan="6" className="p-6 text-center text-slate-500">{t("auditNoRows")}</td></tr>
                            ) : filteredAuditLogs.map((row) => (
                              <tr key={row.id} className="hover:bg-slate-50 align-top">
                                <td className="p-3 font-mono text-xs whitespace-nowrap">{row.at.replace("T", " ").slice(0, 19)} UTC</td>
                                <td className="p-3">{row.user_name}<div className="text-xs text-slate-400">{row.user_email}</div></td>
                                <td className="p-3">{moduleLabel(row.module, t)}</td>
                                <td className="p-3"><span className="px-2 py-0.5 rounded bg-slate-100 text-xs">{actionLabel(row.action)}</span></td>
                                <td className="p-3">{row.entity}</td>
                                <td className="p-3 text-slate-600">{row.detail}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  )}
                </section>
              )}

              {userModal && (
                <Modal title={userModal.mode === "add" ? t("addUser") : t("editUser")} onClose={() => setUserModal(null)} wide>
                  <form onSubmit={saveUser} className="space-y-4">
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                      <Field label={t("userName")}><Input required value={userModal.data.name} onChange={(e) => setUserModal({ ...userModal, data: { ...userModal.data, name: e.target.value } })} /></Field>
                      <Field label={t("userLoginId")}><Input required type="text" autoComplete="username" value={userModal.data.email} onChange={(e) => setUserModal({ ...userModal, data: { ...userModal.data, email: e.target.value } })} /></Field>
                      <Field label={t("userPassword")}><Input required type="password" value={userModal.data.password} onChange={(e) => setUserModal({ ...userModal, data: { ...userModal.data, password: e.target.value } })} /></Field>
                      <Field label={t("userRole")}>
                        <Select value={userModal.data.role} onChange={(e) => setUserModal({ ...userModal, data: { ...userModal.data, role: e.target.value, permissions: e.target.value === "root" ? fullPermissions() : (userModal.data.permissions || emptyPermissions()), allowed_regions: e.target.value === "root" ? ["ALL"] : (userModal.data.allowed_regions || ["HK"]) } })}>
                          <option value="user">{t("roleUser")}</option>
                          <option value="root">{t("roleRoot")}</option>
                        </Select>
                      </Field>
                      <Field label={t("userActive")}>
                        <Select value={userModal.data.is_active ? "1" : "0"} onChange={(e) => setUserModal({ ...userModal, data: { ...userModal.data, is_active: e.target.value === "1" } })}>
                          <option value="1">{t("yes")}</option>
                          <option value="0">{t("no")}</option>
                        </Select>
                      </Field>
                    </div>
                    {userModal.data.role !== "root" && (
                      <div>
                        <p className="text-sm font-semibold mb-2">{t("regionAccess")}</p>
                        <p className="text-xs text-slate-500 mb-2">{t("regionScopeHint")}</p>
                        <div className="flex flex-wrap gap-3">
                          {ERP_REGIONS.map((rg) => (
                            <label key={rg.id} className="text-sm flex items-center gap-1.5">
                              <input type="checkbox" checked={!!(userModal.data.allowed_regions && userModal.data.allowed_regions.includes(rg.id))} onChange={() => toggleUserRegion(rg.id)} />
                              {i18nLang(lang) === "zh_TW" ? rg.labelZh : rg.labelEn}
                            </label>
                          ))}
                        </div>
                      </div>
                    )}
                    {userModal.data.role !== "root" && (
                      <div>
                        <p className="text-sm font-semibold mb-2">{t("accessRights")}</p>
                        <div className="erp-list-scroll erp-list-scroll--settings border rounded-lg">
                          <table className="w-full text-sm">
                            <thead className="bg-slate-50 text-slate-500">
                              <tr>
                                <th className="p-2 text-left">{t("auditModule")}</th>
                                <th className="p-2 text-center">{t("permView")}</th>
                                <th className="p-2 text-center">{t("permAdd")}</th>
                                <th className="p-2 text-center">{t("permEdit")}</th>
                                <th className="p-2 text-center">{t("permDelete")}</th>
                                <th className="p-2 text-center">{t("permExport")}</th>
                                <th className="p-2 text-center">{t("permImport")}</th>
                              </tr>
                            </thead>
                            <tbody className="divide-y">
                              {ERP_MODULES.map((mod) => (
                                <tr key={mod}>
                                  <td className="p-2">{moduleLabel(mod, t)}</td>
                                  {PERM_KEYS.map((perm) => (
                                    <td key={perm} className="p-2 text-center">
                                      <input type="checkbox" checked={!!(userModal.data.permissions && userModal.data.permissions[mod] && userModal.data.permissions[mod][perm])} onChange={() => toggleUserPerm(mod, perm)} />
                                    </td>
                                  ))}
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </div>
                    )}
                    <div className="flex justify-between gap-2">
                      {userModal.mode === "edit" && userModal.data.role !== "root" && userModal.id !== sessionUserId ? (
                        <button type="button" onClick={() => deleteUser(userModal.id)} className="px-4 py-2 rounded-lg bg-red-600 text-white hover:bg-red-700">{t("delete")}</button>
                      ) : <span />}
                      <div className="flex gap-2 ml-auto">
                        <button type="button" onClick={() => setUserModal(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button>
                        <button className="px-4 py-2 rounded-lg bg-blue-600 text-white">{t("save")}</button>
                      </div>
                    </div>
                  </form>
                </Modal>
              )}

              {importLoading && ReactDOM.createPortal(
                <div className="fixed inset-0 z-[60] bg-black/30 flex items-center justify-center p-4">
                  <div className="bg-white rounded-xl border border-slate-200 shadow-xl px-6 py-5 max-w-md text-center">
                    <p className="text-sm font-medium text-slate-800">{t("importLoading")}</p>
                    <p className="text-xs text-slate-500 mt-2">{t("importReading")}</p>
                  </div>
                </div>,
                document.body
              )}

              {importPreview && ReactDOM.createPortal(
                <Modal title={t("importPreviewTitle") + " — " + moduleLabel(importPreview.module, t)} onClose={() => { setImportPreview(null); setImportStatus(""); }} wide>
                  <p className="text-sm text-slate-500 mb-3">{t("importModeHint")} · {importPreview.fileName}</p>
                  {importPreview.parseErrors && importPreview.parseErrors.length > 0 ? (
                    <div className="rounded-lg border border-red-200 bg-red-50 p-4 mb-4">
                      <p className="text-sm font-semibold text-red-800 mb-2">{t("importRowError")}</p>
                      <ul className="text-sm text-red-700 list-disc pl-5 space-y-1">
                        {importPreview.parseErrors.map((msg, i) => <li key={i}>{msg}</li>)}
                      </ul>
                      <p className="text-xs text-red-600 mt-3">{t("importParseFailed")}</p>
                    </div>
                  ) : (
                    <>
                      <div className="erp-list-scroll erp-list-scroll--dash border rounded-lg mb-4">
                        <table className="w-full text-sm">
                          <thead className="bg-slate-50 text-slate-500">
                            <tr>
                              <th className="p-2 text-left">#</th>
                              <th className="p-2 text-left">{t("auditAction")}</th>
                              <th className="p-2 text-left">Key</th>
                              <th className="p-2 text-left">{t("auditDetail")}</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y">
                            {importPreview.preview.map((p) => (
                              <tr key={p.lineNo} className={p.action === "error" ? "bg-red-50" : p.action === "create" ? "bg-green-50" : "bg-amber-50"}>
                                <td className="p-2">{p.lineNo}</td>
                                <td className="p-2">{p.action === "error" ? t("importRowError") : p.action === "create" ? t("importRowCreate") : t("importRowUpdate")}</td>
                                <td className="p-2 font-mono text-xs">{p.key || "-"}</td>
                                <td className="p-2 text-xs">{p.errors.length ? p.errors.join("; ") : (p.data.company || p.data.job_no || "")}</td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <div className="flex justify-between items-center gap-2 flex-wrap">
                        <p className="text-sm text-slate-600">
                          {t("importSummary")}: {importPreview.preview.filter((p) => p.action === "create").length} {t("importRowCreate")},{" "}
                          {importPreview.preview.filter((p) => p.action === "update").length} {t("importRowUpdate")},{" "}
                          {importPreview.preview.filter((p) => p.action === "error").length} {t("importRowError")}
                        </p>
                        <div className="flex gap-2">
                          <button type="button" onClick={() => setImportPreview(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button>
                          <button type="button" onClick={applyImportPreview} disabled={!importPreview.preview.some((p) => p.action !== "error")} className="px-4 py-2 rounded-lg bg-emerald-600 text-white disabled:opacity-40">{t("importConfirm")}</button>
                        </div>
                      </div>
                    </>
                  )}
                  {importPreview.parseErrors && importPreview.parseErrors.length > 0 && (
                    <div className="flex justify-end mt-4">
                      <button type="button" onClick={() => setImportPreview(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button>
                    </div>
                  )}
                </Modal>,
                document.body
              )}

              {newTxOpen && (
                <Modal title={t("newTxTitle")} onClose={() => setNewTxOpen(false)}>
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {can("ar", "add") && <button onClick={() => { setArModal({ mode: "add", data: { job_id: "", job_no: "", customer: "", customer_po: "", job_completed: false, invoice_no: "", invoice_currency: "USD", invoice_amt: "", exchange_rate: "", base_amount: "", override_reason: "", invoice_date: "", due_date: "", payment_received_date: "", payment_status: "Awaiting Payment" } }); setNewTxOpen(false); }} className="rounded-lg border p-4 text-left hover:bg-slate-50">
                      <p className="font-semibold">{t("createArTitle")}</p><p className="text-xs text-slate-500">{t("createArDesc")}</p>
                    </button>}
                    {can("ap", "add") && <button onClick={() => { setApModal({ mode: "add", data: emptyApForm() }); setNewTxOpen(false); }} className="rounded-lg border p-4 text-left hover:bg-slate-50">
                      <p className="font-semibold">{t("createApTitle")}</p><p className="text-xs text-slate-500">{t("createApDesc")}</p>
                    </button>}
                  </div>
                </Modal>
              )}

              {clientModal && (
                <Modal title={clientModal.mode === "add" ? t("addClientTitle") : t("editClient")} onClose={() => setClientModal(null)}>
                  <form onSubmit={saveClient} className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <Field label={t("colCustomerNo")}><Input required value={clientModal.data.customer_no} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, customer_no: e.target.value } })} /></Field>
                    <Field label={t("colCompany")}><Input required value={clientModal.data.company} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, company: e.target.value } })} /></Field>
                    <Field label="Invoice Title"><Input value={clientModal.data.invoice_title || ""} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, invoice_title: e.target.value } })} /></Field>
                    <Field label={t("colGstNo")}><Input value={clientModal.data.gst_no} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, gst_no: e.target.value } })} /></Field>
                    <Field label={t("colPrimaryContact")}><Input value={clientModal.data.primary_contact} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, primary_contact: e.target.value } })} /></Field>
                    <Field label={t("colCompanyPhone")}><Input value={clientModal.data.company_phone} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, company_phone: e.target.value } })} /></Field>
                    <Field label={t("colMobilePhone")}><Input value={clientModal.data.mobile_phone} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, mobile_phone: e.target.value } })} /></Field>
                    <Field label={t("colEmail")}><Input type="email" value={clientModal.data.email} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, email: e.target.value } })} /></Field>
                    <Field label={t("address")}><Input value={clientModal.data.address} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, address: e.target.value } })} /></Field>
                    <Field label={t("colPostalCode")}><Input value={clientModal.data.postal_code} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, postal_code: e.target.value } })} /></Field>
                    <Field label={t("colAccountDeptContact")}><Input value={clientModal.data.account_dept_contact} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, account_dept_contact: e.target.value } })} /></Field>
                    <Field label={t("colPaymentTerms")}>
                      {(() => {
                        const raw = String(clientModal.data.payment_terms || "").trim();
                        const m = raw.match(/(\d+)/);
                        const days = m ? parseInt(m[1], 10) : null;
                        const quick = [30, 45, 60, 90];
                        const isQuick = days != null && quick.includes(days);
                        return (
                          <div className="flex flex-col gap-2">
                            <div className="flex flex-wrap gap-2">
                              {quick.map((d) => (
                                <button
                                  key={d}
                                  type="button"
                                  onClick={() => setClientModal({ ...clientModal, data: { ...clientModal.data, payment_terms: `${d} Days` } })}
                                  className={
                                    "px-2 py-1 text-xs rounded-lg border " +
                                    (isQuick && days === d ? "bg-blue-600 text-white border-blue-600" : "bg-white text-slate-700 border-slate-300 hover:bg-slate-50")
                                  }
                                >
                                  {d} Days
                                </button>
                              ))}
                              {!isQuick && <span className="text-[10px] text-slate-400 mt-2">Custom</span>}
                            </div>
                            <Input value={clientModal.data.payment_terms} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, payment_terms: e.target.value } })} />
                          </div>
                        );
                      })()}
                    </Field>
                    <div className="md:col-span-2 flex items-center gap-2">
                      <input type="checkbox" id="client-bu" checked={!!clientModal.data.is_bu} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, is_bu: e.target.checked, bu_no: e.target.checked ? (clientModal.data.bu_no || "") : "" } })} className="rounded border-slate-300" />
                      <label htmlFor="client-bu" className="text-sm text-slate-700">{t("colBu")}</label>
                      <span className="text-[10px] text-slate-400">{t("clientBuHint")}</span>
                    </div>
                    {clientModal.data.is_bu && (
                      <Field label={t("colBuNo")}><Input value={clientModal.data.bu_no || ""} onChange={(e) => setClientModal({ ...clientModal, data: { ...clientModal.data, bu_no: e.target.value } })} /></Field>
                    )}
                    <div className="md:col-span-2 flex justify-end gap-2"><button type="button" onClick={() => setClientModal(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button><button className="px-4 py-2 rounded-lg bg-blue-600 text-white">{t("save")}</button></div>
                  </form>
                </Modal>
              )}

              {quotationModal && (
                <Modal title={quotationModal.mode === "add" ? t("addQuotationTitle") : t("editQuotation")} onClose={() => setQuotationModal(null)}>
                  <form onSubmit={saveQuotation} className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <Field label={t("quotationNo")}><Input required value={quotationModal.data.quotation_no} onChange={(e) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, quotation_no: e.target.value } })} /></Field>
                    <Field label={t("clientGroup")}>
                      <SearchableSelect
                        required
                        value={quotationModal.data.client_id}
                        options={clientOptionsForQuotation}
                        placeholder={t("selectClientHint")}
                        noResultsText={t("noMatchFound")}
                        onChange={(nextValue) => {
                          const client = clients.find((c) => c.id === Number(nextValue));
                          setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, client_id: nextValue, company: client ? client.company : "" } });
                        }}
                      />
                    </Field>
                    <Field label={t("quotationDate")}><Input type="date" value={quotationModal.data.quotation_date} onChange={(e) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, quotation_date: e.target.value } })} /></Field>
                    <Field label={t("validUntil")}><Input type="date" value={quotationModal.data.valid_until} onChange={(e) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, valid_until: e.target.value } })} /></Field>
                    <Field label={t("amount")}><Input type="number" step="0.01" required value={quotationModal.data.amount} onChange={(e) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, amount: e.target.value } })} /></Field>
                    <Field label={t("currency")}>
                      <Select value={quotationModal.data.currency || "USD"} onChange={(e) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, currency: e.target.value } })}>
                        {Object.keys(fxUsdMap).map((c) => <option key={c} value={c}>{c}</option>)}
                      </Select>
                    </Field>
                    <Field label={t("quotationStatus")}>
                      <Select value={quotationModal.data.status} onChange={(e) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, status: e.target.value } })}>
                        {QUOTATION_STATUSES.map((s) => <option key={s} value={s}>{s}</option>)}
                      </Select>
                    </Field>
                    <Field label={t("colJobNo")}>
                      <SearchableSelect
                        value={quotationModal.data.job_no || ""}
                        options={quotationJobOptions}
                        placeholder={t("selectJobHint")}
                        noResultsText={t("noMatchFound")}
                        onChange={(nextValue) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, job_no: nextValue } })}
                      />
                    </Field>
                    <div className="md:col-span-2"><Field label={t("description")}><Input value={quotationModal.data.description} onChange={(e) => setQuotationModal({ ...quotationModal, data: { ...quotationModal.data, description: e.target.value } })} /></Field></div>
                    <div className="md:col-span-2 flex justify-end gap-2"><button type="button" onClick={() => setQuotationModal(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button><button className="px-4 py-2 rounded-lg bg-blue-600 text-white">{t("save")}</button></div>
                  </form>
                </Modal>
              )}

              {jobModal && (
                <Modal title={jobModal.mode === "add" ? t("addJobTitle") : t("editJob")} onClose={() => setJobModal(null)}>
                  <form onSubmit={saveJob} className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <Field label={t("colJobNo")}><Input required value={jobModal.data.job_no} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, job_no: e.target.value } })} /></Field>
                    <Field label={t("clientGroup")}>
                      <SearchableSelect
                        required
                        value={jobModal.data.client_id}
                        options={clientOptionsForJob}
                        placeholder={t("selectClientHint")}
                        noResultsText={t("noMatchFound")}
                        onChange={(nextValue) => {
                          const client = clients.find((c) => c.id === Number(nextValue));
                          setJobModal({ ...jobModal, data: { ...jobModal.data, client_id: nextValue, company: client ? client.company : "", region: client?.region || jobModal.data.region || newRecordRegion } });
                        }}
                      />
                    </Field>
                    <Field label="Type">
                      <Select required value={jobModal.data.job_type} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, job_type: e.target.value } })}>
                        {JOB_TYPES.map((t) => <option key={t} value={t}>{t}</option>)}
                      </Select>
                    </Field>
                    <Field label={t("currency")}>
                      <Select value={jobModal.data.currency || "USD"} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, currency: e.target.value } })}>
                        {Object.keys(fxUsdMap).map((c) => <option key={c} value={c}>{c}</option>)}
                      </Select>
                    </Field>
                    <div className="md:col-span-2 border rounded-lg p-3 bg-slate-50/50">
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-sm font-semibold text-slate-700">{t("jobQuotationLinesTitle")}</p>
                      </div>
                      <div className="flex gap-2 mb-2">
                        <Input
                          className="flex-1"
                          list="erp-accepted-quotations"
                          placeholder={t("quotationNo")}
                          value={jobModal.data._quotationDraft || ""}
                          onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, _quotationDraft: e.target.value } })}
                          onKeyDown={(e) => { if (e.key === "Enter") { e.preventDefault(); addQuotationDraftToJob(); } }}
                          onBlur={() => { if (jobModal.data._quotationDraft) addQuotationDraftToJob(); }}
                        />
                        <button type="button" className="text-xs px-2 py-1.5 rounded bg-blue-600 text-white hover:bg-blue-700 whitespace-nowrap" onClick={addQuotationDraftToJob}>{t("jobQuotationAdd")}</button>
                      </div>
                      <datalist id="erp-accepted-quotations">
                        {quotationsForJobRegion.filter((q) => q.status === "Accepted").map((q) => (
                          <option key={q.id} value={q.quotation_no}>{q.company} · {money(q.amount)} {q.currency}{q.job_no ? " · " + q.job_no : ""}</option>
                        ))}
                      </datalist>
                      {jobQuotationNos(jobModal.data).length ? (
                        <ul className="space-y-1 mb-2">
                          {jobQuotationNos(jobModal.data).map((qNo) => {
                            const q = findQuotationByNo(qNo, quotationsForJobRegion);
                            const others = q ? quotationJobNos(q).filter((n) => n !== jobModal.data.job_no) : [];
                            return (
                              <li key={qNo} className="flex items-start justify-between gap-2 text-xs bg-white border rounded px-2 py-1.5">
                                <div>
                                  <span className="font-medium text-slate-800">{qNo}</span>
                                  {q ? <span className="text-slate-500 ml-2">{money(q.amount)} {q.currency}</span> : null}
                                  {others.length ? <p className="text-[10px] text-amber-600 mt-0.5">{t("quotationMultiJobHint").replace("{jobs}", others.join(", "))}</p> : null}
                                </div>
                                <button type="button" className="text-red-500 hover:text-red-700 shrink-0" onClick={() => removeQuotationFromJobModal(qNo)} title={t("delete")}>×</button>
                              </li>
                            );
                          })}
                        </ul>
                      ) : null}
                      <p className="text-[10px] text-slate-400">{t("jobQuotationInputHint")}</p>
                    </div>
                    <Field label="Quotation Amount"><Input type="number" step="0.01" value={jobModal.data.amount} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, amount: e.target.value } })} /></Field>
                    <Field label={t("colCustomerPo")}><Input value={jobModal.data.customer_po} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, customer_po: e.target.value } })} /></Field>
                    <Field label={t("customerPoAmountTotal")}><Input type="number" step="0.01" readOnly value={jobModal.data.customer_po_amount || ""} className="bg-slate-50" /></Field>
                    <div className="md:col-span-2 border rounded-lg p-3 bg-slate-50/50">
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-sm font-semibold text-slate-700">{t("jobPoLinesTitle")}</p>
                        <button type="button" className="text-xs px-2 py-1 rounded bg-blue-600 text-white hover:bg-blue-700" onClick={() => setJobModal({ ...jobModal, data: { ...jobModal.data, po_lines: [...(jobModal.data.po_lines || []), emptyPoLine()] } })}>{t("jobPoAddLine")}</button>
                      </div>
                      <div className="overflow-x-auto">
                        <table className="w-full text-xs border-collapse">
                          <thead><tr className="text-slate-500">
                            <th className="p-1.5 text-left font-medium">{t("jobPoColNo")}</th>
                            <th className="p-1.5 text-right font-medium">{t("jobPoColAmount")}</th>
                            <th className="p-1.5 text-left font-medium">{t("jobPoColDate")}</th>
                            <th className="p-1.5 text-left font-medium">{t("jobPoColRemarks")}</th>
                            <th className="p-1.5 w-8"></th>
                          </tr></thead>
                          <tbody>
                            {(jobModal.data.po_lines || []).map((pl, idx) => (
                              <tr key={pl.id || idx}>
                                <td className="p-1"><input className="w-full border border-slate-300 rounded px-1.5 py-1" value={pl.po_no || ""} onChange={(e) => {
                                  const po_lines = [...(jobModal.data.po_lines || [])];
                                  po_lines[idx] = { ...po_lines[idx], po_no: e.target.value };
                                  const total = po_lines.reduce((s, l) => s + Number(l.amount || 0), 0);
                                  setJobModal({ ...jobModal, data: { ...jobModal.data, po_lines, customer_po: po_lines.map((l) => l.po_no).filter(Boolean).join(", "), customer_po_amount: total } });
                                }} /></td>
                                <td className="p-1"><input type="number" step="0.01" className="w-full border border-slate-300 rounded px-1.5 py-1 text-right" value={pl.amount ?? ""} onChange={(e) => {
                                  const po_lines = [...(jobModal.data.po_lines || [])];
                                  po_lines[idx] = { ...po_lines[idx], amount: e.target.value };
                                  const total = po_lines.reduce((s, l) => s + Number(l.amount || 0), 0);
                                  setJobModal({ ...jobModal, data: { ...jobModal.data, po_lines, customer_po_amount: total } });
                                }} /></td>
                                <td className="p-1"><input type="date" className="w-full border border-slate-300 rounded px-1.5 py-1" value={pl.received_date || ""} onChange={(e) => {
                                  const po_lines = [...(jobModal.data.po_lines || [])];
                                  po_lines[idx] = { ...po_lines[idx], received_date: e.target.value };
                                  const dates = po_lines.map((l) => l.received_date).filter(Boolean).sort();
                                  setJobModal({ ...jobModal, data: { ...jobModal.data, po_lines, start_date: dates[0] || jobModal.data.start_date } });
                                }} /></td>
                                <td className="p-1"><input className="w-full border border-slate-300 rounded px-1.5 py-1" value={pl.remarks || ""} onChange={(e) => {
                                  const po_lines = [...(jobModal.data.po_lines || [])];
                                  po_lines[idx] = { ...po_lines[idx], remarks: e.target.value };
                                  setJobModal({ ...jobModal, data: { ...jobModal.data, po_lines } });
                                }} /></td>
                                <td className="p-1 text-center"><button type="button" className="text-red-500 hover:text-red-700" onClick={() => {
                                  const po_lines = (jobModal.data.po_lines || []).filter((_, i) => i !== idx);
                                  const total = po_lines.reduce((s, l) => s + Number(l.amount || 0), 0);
                                  setJobModal({ ...jobModal, data: { ...jobModal.data, po_lines, customer_po: po_lines.map((l) => l.po_no).filter(Boolean).join(", "), customer_po_amount: total } });
                                }}>×</button></td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                      <p className="text-[10px] text-slate-400 mt-2">{t("jobPoLinesTitle")}: one row per PO; received date sets Monthly Report month.</p>
                    </div>
                    <Field label={t("colStartDate")}><Input type="date" value={jobModal.data.start_date} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, start_date: e.target.value } })} /></Field>
                    <Field label="Status">
                      <Select value={jobModal.data.status} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, status: e.target.value } })}>
                        <option>Open</option><option>In Progress</option><option>Completed</option><option>On Hold</option>
                      </Select>
                    </Field>
                    <div className="md:col-span-1"><Field label="Description"><Input value={jobModal.data.description} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, description: e.target.value } })} /></Field></div>
                    <div className="md:col-span-1"><Field label={t("ongoingColRemarks")}><Input value={jobModal.data.remarks || ""} onChange={(e) => setJobModal({ ...jobModal, data: { ...jobModal.data, remarks: e.target.value } })} /></Field></div>
                    <div className="md:col-span-2 flex justify-end gap-2"><button type="button" onClick={() => setJobModal(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button><button className="px-4 py-2 rounded-lg bg-blue-600 text-white">{t("save")}</button></div>
                  </form>
                </Modal>
              )}

              {vendorModal && (
                <Modal title={vendorModal.mode === "add" ? t("addVendorTitle") : t("editVendor")} onClose={() => setVendorModal(null)}>
                  <form onSubmit={saveVendor} className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <Field label={t("vendorNo")}><Input required value={vendorModal.data.vendor_no} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, vendor_no: e.target.value } })} /></Field>
                    <Field label="Vendor Name"><Input required value={vendorModal.data.name} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, name: e.target.value } })} /></Field>
                    <Field label="Email"><Input type="email" value={vendorModal.data.email} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, email: e.target.value } })} /></Field>
                    <Field label="Bank"><Input value={vendorModal.data.bank} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, bank: e.target.value } })} /></Field>
                    <Field label="Charge"><Input value={vendorModal.data.charge || ""} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, charge: e.target.value } })} placeholder="OUR" /></Field>
                    <Field label="SWIFT CODE"><Input value={vendorModal.data.swift_code} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, swift_code: e.target.value } })} /></Field>
                    <Field label="Contact"><Input value={vendorModal.data.contact} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, contact: e.target.value } })} /></Field>
                    <Field label={t("colFinanceContact")}><Input value={vendorModal.data.finance_contact || ""} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, finance_contact: e.target.value } })} /></Field>
                    <Field label={t("colFinanceEmail")}><Input type="email" value={vendorModal.data.finance_email || ""} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, finance_email: e.target.value } })} /></Field>
                    <Field label={t("phoneNo")}><Input value={vendorModal.data.phone} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, phone: e.target.value } })} /></Field>
                    <div className="md:col-span-2"><Field label={t("address")}><Input value={vendorModal.data.address || ""} onChange={(e) => setVendorModal({ ...vendorModal, data: { ...vendorModal.data, address: e.target.value } })} /></Field></div>
                    <div className="md:col-span-2 flex justify-end gap-2"><button type="button" onClick={() => setVendorModal(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button><button className="px-4 py-2 rounded-lg bg-blue-600 text-white">{t("save")}</button></div>
                  </form>
                </Modal>
              )}

              {arModal && (
                <Modal title={arModal.mode === "add" ? t("createArTitle") : t("editAr")} onClose={() => setArModal(null)} wide>
                  <form onSubmit={saveAR} className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <Field label="Link Job">
                      <SearchableSelect
                        value={arModal.data.job_id}
                        options={jobOptions.map((opt) => {
                          const row = sortedScopedJobs.find((j) => String(j.id) === String(opt.value));
                          return { ...opt, label: `${opt.label} · ${row ? row.job_type : ""}` };
                        })}
                        placeholder={t("selectJobOptionalHint")}
                        noResultsText={t("noMatchFound")}
                        onChange={(nextValue) => {
                          const job = jobs.find((j) => j.id === Number(nextValue));
                          const client = job ? clients.find((c) => c.id === job.client_id) : null;
                          const customerTitle = client?.invoice_title || (job ? job.company : arModal.data.customer);
                          setArModal({ ...arModal, data: patchArModalData({ ...arModal.data, job_id: nextValue, job_no: job ? job.job_no : "", customer: customerTitle, customer_po: job ? job.customer_po : arModal.data.customer_po }) });
                        }}
                      />
                    </Field>
                    <Field label="Customer"><Input value={arModal.data.customer} onChange={(e) => setArModal({ ...arModal, data: patchArModalData({ ...arModal.data, customer: e.target.value }) })} /></Field>
                    <Field label={t("colCustomerPo")}><Input value={arModal.data.customer_po} onChange={(e) => setArModal({ ...arModal, data: { ...arModal.data, customer_po: e.target.value } })} /></Field>
                    {(() => {
                      const linked = jobs.find((j) => j.id === Number(arModal.data.job_id));
                      return linked ? <Field label={t("jobStatus")}><Input readOnly value={linked.status} className="bg-slate-50" /></Field> : null;
                    })()}
                    <Field label={t("colInvoiceNo")}><Input required value={arModal.data.invoice_no} onChange={(e) => setArModal({ ...arModal, data: { ...arModal.data, invoice_no: e.target.value } })} /></Field>
                    <Field label={t("paymentStage")}>
                      <Select value={arModal.data.payment_stage || ""} onChange={(e) => setArModal({ ...arModal, data: { ...arModal.data, payment_stage: e.target.value } })}>
                        {PAYMENT_STAGES.map((s) => <option key={s || "none"} value={s}>{s || "—"}</option>)}
                      </Select>
                      <p className="text-[10px] text-slate-400 mt-1">{t("paymentStageHint")}</p>
                    </Field>
                    <Field label="Invoice Currency">
                      <Select value={arModal.data.invoice_currency || "USD"} onChange={(e) => setArModal({ ...arModal, data: { ...arModal.data, invoice_currency: e.target.value } })}>
                        {Object.keys(fxUsdMap).map((c) => <option key={c} value={c}>{c}</option>)}
                      </Select>
                    </Field>
                    <Field label={t("invoiceAmountOrig")}>
                      <Input type="number" step="0.01" required value={arModal.data.invoice_amt} onChange={(e) => calcBaseEditable({ invoice_amt: e.target.value }, arModal.data, (d) => setArModal((prev) => ({ ...prev, data: d })))} />
                      <p className="text-[10px] text-slate-400 mt-1">{t("invoiceAmountHint")}</p>
                    </Field>
                    <Field label="Exchange Rate"><Input type="number" step="0.0001" required value={arModal.data.exchange_rate} onChange={(e) => calcBaseEditable({ exchange_rate: e.target.value }, arModal.data, (d) => setArModal((prev) => ({ ...prev, data: d })))} /></Field>
                    <Field label={`Amount in ${regionListCurrency} (${t("amountManualOverride")})`}>
                      <Input type="number" step="0.01" value={arModal.data.base_amount} onChange={(e) => setArModal({ ...arModal, data: { ...arModal.data, base_amount: e.target.value } })} />
                      <p className="text-[10px] text-slate-400 mt-1">{t("baseAmountHint")}</p>
                    </Field>
                    <Field label="Invoice Date"><Input type="date" required value={arModal.data.invoice_date} onChange={(e) => setArModal({ ...arModal, data: patchArModalData({ ...arModal.data, invoice_date: e.target.value }) })} /></Field>
                    <Field label="Due Date">
                      <Input type="date" required value={arModal.data.due_date} onChange={(e) => {
                        const next = { ...arModal.data, due_date: e.target.value };
                        setArModal({ ...arModal, data: { ...next, payment_status: deriveArPaymentStatus(next) } });
                      }} />
                      <p className="text-[10px] text-slate-400 mt-1">{t("arDueAutoHint")}</p>
                    </Field>
                    <Field label={t("paymentReceived")}><Input type="date" value={arModal.data.payment_received_date} onChange={(e) => {
                      const v = e.target.value;
                      const next = { ...arModal.data, payment_received_date: v };
                      setArModal({ ...arModal, data: { ...next, payment_status: deriveArPaymentStatus(next) } });
                    }} /></Field>
                    <Field label={t("paymentStatus")}>
                      <p className="px-3 py-2 border rounded-lg bg-slate-50 text-sm">
                        <span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(deriveArPaymentStatus(arModal.data))}>{paymentStatusLabel(deriveArPaymentStatus(arModal.data), t)}</span>
                      </p>
                      <p className="text-[10px] text-slate-400 mt-1">{t("arPaymentStatusAutoHint")}</p>
                    </Field>
                    <div className="md:col-span-2 flex justify-end gap-2"><button type="button" onClick={() => setArModal(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button><button className="px-4 py-2 rounded-lg bg-blue-600 text-white">{t("save")}</button></div>
                  </form>
                </Modal>
              )}

              {apModal && (
                <Modal title={apModal.mode === "add" ? t("createApTitle") : t("editAp")} onClose={() => setApModal(null)} wide>
                  <form onSubmit={saveAP} className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    <Field label="Link Job">
                      <SearchableSelect
                        value={apModal.data.job_id}
                        options={jobOptions}
                        placeholder={t("selectJobHint")}
                        noResultsText={t("noMatchFound")}
                        onChange={(nextValue) => {
                          const job = jobs.find((j) => j.id === Number(nextValue));
                          setApModal({ ...apModal, data: { ...apModal.data, job_id: nextValue, job_no: job ? job.job_no : "" } });
                        }}
                      />
                    </Field>
                    <Field label={t("colCompanyName")}>
                      <SearchableSelect
                        required
                        value={apModal.data.company_name}
                        options={[
                          ...(apModal.data.company_name && !sortedScopedVendors.some((v) => v.name === apModal.data.company_name) ? [{
                            value: apModal.data.company_name,
                            label: apModal.data.company_name + (apModal.data.payee_type === "SI" ? " (" + t("payeeSi") + ")" : ""),
                            searchText: apModal.data.company_name
                          }] : []),
                          ...vendorOptions
                        ]}
                        placeholder={t("selectVendorHint")}
                        noResultsText={t("noMatchFound")}
                        onChange={(nextValue) => {
                          const patch = patchApFromPayeeMaster(nextValue, vendors);
                          setApModal({ ...apModal, data: { ...apModal.data, payee_type: "Vendor", si_id: "", ...patch } });
                        }}
                      />
                      {apModal.data.payee_type === "SI" && apModal.data.company_name && (
                        <p className="text-[10px] text-amber-600 mt-1">{t("payeeSi")}: {apModal.data.company_name}</p>
                      )}
                      <p className="text-[10px] text-slate-400 mt-1">{t("apPayeeAutoHint")}</p>
                    </Field>
                    <Field label={t("colInvoiceNo")}><Input required value={apModal.data.invoice_no} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, invoice_no: e.target.value } })} /></Field>
                    <Field label={t("paymentStage")}>
                      <Select value={apModal.data.payment_stage || ""} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, payment_stage: e.target.value } })}>
                        {PAYMENT_STAGES.map((s) => <option key={s || "none"} value={s}>{s || "—"}</option>)}
                      </Select>
                      <p className="text-[10px] text-slate-400 mt-1">{t("paymentStageHint")}</p>
                    </Field>
                    <Field label="Invoice Date"><Input type="date" required value={apModal.data.invoice_date} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, invoice_date: e.target.value } })} /></Field>
                    <Field label="Invoice Received Date"><Input type="date" value={apModal.data.invoice_received_date} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, invoice_received_date: e.target.value } })} /></Field>
                    <Field label="Due Date"><Input type="date" required value={apModal.data.due_date} onChange={(e) => {
                      const next = { ...apModal.data, due_date: e.target.value };
                      setApModal({ ...apModal, data: { ...next, payment_status: deriveApPaymentStatus(next) } });
                    }} /></Field>
                    <Field label="Currency">
                      <Select value={apModal.data.currency || "USD"} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, currency: e.target.value } })}>
                        {Object.keys(fxUsdMap).map((c) => <option key={c} value={c}>{c}</option>)}
                      </Select>
                    </Field>
                    <Field label={t("invoiceAmountOrig")}>
                      <Input type="number" step="0.01" required value={apModal.data.amount} onChange={(e) => calcBaseEditable({ amount: e.target.value }, apModal.data, (d) => setApModal((prev) => ({ ...prev, data: d })))} />
                      <p className="text-[10px] text-slate-400 mt-1">{t("invoiceAmountHint")}</p>
                    </Field>
                    <Field label="Exchange Rate"><Input type="number" step="0.0001" required value={apModal.data.exchange_rate} onChange={(e) => calcBaseEditable({ exchange_rate: e.target.value }, apModal.data, (d) => setApModal((prev) => ({ ...prev, data: d })))} /></Field>
                    <Field label={`Amount in ${regionListCurrency} (${t("amountManualOverride")})`}>
                      <Input type="number" step="0.01" value={apModal.data.base_amount} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, base_amount: e.target.value } })} />
                      <p className="text-[10px] text-slate-400 mt-1">{t("baseAmountHint")}</p>
                    </Field>
                    <Field label="Pay Date"><Input type="date" value={apModal.data.pay_date} onChange={(e) => {
                      const v = e.target.value;
                      const next = { ...apModal.data, pay_date: v };
                      setApModal({ ...apModal, data: { ...next, payment_status: deriveApPaymentStatus(next) } });
                    }} /></Field>
                    <Field label={t("paymentStatus")}>
                      <p className="px-3 py-2 border rounded-lg bg-slate-50 text-sm">
                        <span className={"px-2 py-0.5 rounded text-xs " + paymentStatusClass(deriveApPaymentStatus(apModal.data))}>{paymentStatusLabel(deriveApPaymentStatus(apModal.data), t)}</span>
                      </p>
                      <p className="text-[10px] text-slate-400 mt-1">{t("apPaymentStatusAutoHint")}</p>
                    </Field>
                    <Field label="Remarks"><Input value={apModal.data.remarks} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, remarks: e.target.value } })} /></Field>
                    <Field label="Bank"><Input value={apModal.data.bank} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, bank: e.target.value } })} className="bg-slate-50" /></Field>
                    <Field label="Charge"><Input value={apModal.data.charge} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, charge: e.target.value } })} placeholder="OUR" className="bg-slate-50" /></Field>
                    <Field label="SWIFT CODE"><Input value={apModal.data.swift_code} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, swift_code: e.target.value } })} className="bg-slate-50" /></Field>
                    <Field label="Payment Advice Email"><Input type="email" value={apModal.data.payment_advice_email} onChange={(e) => setApModal({ ...apModal, data: { ...apModal.data, payment_advice_email: e.target.value } })} className="bg-slate-50" /></Field>
                    <div className="md:col-span-2 flex justify-end gap-2"><button type="button" onClick={() => setApModal(null)} className="px-4 py-2 rounded-lg border">{t("cancel")}</button><button className="px-4 py-2 rounded-lg bg-blue-600 text-white">{t("save")}</button></div>
                  </form>
                </Modal>
              )}

              {renderDetailPanel()}
            </main>
          </div>
        );
      }

      ReactDOM.createRoot(document.getElementById("app")).render(<ErpRoot />);
