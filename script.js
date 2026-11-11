// Tắt kiểm tra TypeScript cho file này và khai báo các thư viện toàn cục
// @ts-nocheck
/* global flatpickr, XLSX, html2canvas, jspdf, qrcode, DEFAULT_PRODUCTS, LOGO_DATA_URI */

// Bọc toàn bộ code trong một IIFE để tránh làm ô nhiễm global scope
// Phụ thuộc (nạp trong index.html): flatpickr, html2canvas, jsPDF, SheetJS (XLSX),
// qrcode-generator, tariff-defaults.js (DEFAULT_PRODUCTS) và logo-data.js (LOGO_DATA_URI).
(function() {
    'use strict';

    // --- 1. STATE & CONFIG ---
    const BANK = {
        bin: '970418',                       // BIDV (mã Napas)
        account: '8608393979',
        accountName: 'CONG TY CO PHAN CANG SAI GON'
    };

    const i18n = {
        appTitle: { vi: 'Báo giá dịch vụ', en: 'Service Quotation' },
        developedBy: { vi: 'Phát triển bởi: Tiền/Cảng Tân Thuận', en: 'Developed by: Tien/Tan Thuan Port' },
        manageServices: { vi: 'Phương án (Tariff)', en: 'Manage Services (Tariff)' },
        history: { vi: 'Lịch sử báo giá', en: 'Quotation History' },
        formTitle: { vi: 'Thông tin chung', en: 'Informations' },
        quoteNo: { vi: 'Số báo giá', en: 'Quotation No.' },
        quoteDate: { vi: 'Ngày báo giá', en: 'Date' },
        customerName: { vi: 'Tên khách hàng', en: 'Customer Name' },
        taxCode: { vi: 'Mã số thuế', en: 'Tax Code' },
        address: { vi: 'Địa chỉ', en: 'Address' },
        phone: { vi: 'Điện thoại', en: 'Phone' },
        notes: { vi: 'Ghi chú', en: 'Notes' },
        thService: { vi: 'Phương án', en: 'Service' },
        thQty: { vi: 'Số lượng', en: 'Quantity' },
        thUnitPrice: { vi: 'Đơn giá', en: 'Unit Price' },
        thAmount: { vi: 'Thành tiền', en: 'Amount' },
        addRow: { vi: '+ Thêm dòng', en: '+ Add Row' },
        total: { vi: 'Tổng cộng:', en: 'Total:' },
        inWords: { vi: 'Bằng chữ:', en: 'In words:' },
        actionsTitle: { vi: 'Xem trước', en: 'Preview' },
        newQuote: { vi: 'Tạo mới', en: 'New' },
        saveQuote: { vi: 'Lưu báo giá', en: 'Save Quotation' },
        updatePreview: { vi: 'Cập nhật', en: 'Update Preview' },
        printPdf: { vi: 'In / Xuất PDF', en: 'Print / Export PDF' },
        previewPlaceholder: { vi: 'Bản xem trước (A4)', en: 'A4 preview' },
        modalProductTitle: { vi: 'Quản lý phương án (Tariff)', en: 'Manage Services (Tariff)' },
        importExcel: { vi: 'Import Excel', en: 'Import Excel' },
        exportExcel: { vi: 'Export Excel', en: 'Export Excel' },
        backupData: { vi: 'Sao lưu Dữ liệu', en: 'Backup Data' },
        restoreData: { vi: 'Phục hồi Dữ liệu', en: 'Restore Data' },
        modalName: { vi: 'Tên', en: 'Name' },
        modalUnit: { vi: 'ĐVT', en: 'Unit' },
        modalPrice: { vi: 'Đơn giá', en: 'Price' },
        btnModalSave: { vi: 'Lưu', en: 'Save' },
        btnCancel: { vi: 'Hủy', en: 'Cancel' },
        modalHistoryTitle: { vi: 'Dữ liệu báo giá', en: 'Quotation Data' },
        action: { vi: 'Hành động', en: 'Action' },
        date: { vi: 'Ngày', en: 'Date' },
        customer: { vi: 'Khách hàng', en: 'Customer' },
        view: { vi: 'Xem', en: 'View' },
        edit: { vi: 'Sửa', en: 'Edit' },
        delete: { vi: 'Xóa', en: 'Delete' },
        dong: { vi: 'đồng', en: 'Dong' },
        dateRangeFrom: { vi: 'Từ ngày:', en: 'From date:' },
        dateRangeTo: { vi: 'Đến ngày:', en: 'To date:' },
        dateTimeRangeFrom: { vi: 'Từ lúc:', en: 'From:' },
        dateTimeRangeTo: { vi: 'Đến lúc:', en: 'To:' },
        pickStartDate: { vi: 'Chọn ngày bắt đầu', en: 'Select start date' },
        pickEndDate: { vi: 'Chọn ngày kết thúc', en: 'Select end date' },
        pickStartDateTime: { vi: 'Chọn ngày & giờ bắt đầu', en: 'Select start date & time' },
        pickEndDateTime: { vi: 'Chọn ngày & giờ kết thúc', en: 'Select end date & time' },
        otherOption: { vi: 'Phương án khác', en: 'Other Service' },
        selectOption: { vi: '-- Chọn phương án --', en: '-- Select a service --' },
        storageKeywords: { vi: 'lưu bãi', en: 'storage' },
        powerKeywords: { vi: 'chạy điện', en: 'power supply' },
        containerQty: { vi: 'Số lượng container:', en: 'Number of containers:' },
        confirmClear: { vi: 'Tạo báo giá mới sẽ xóa toàn bộ thông tin hiện tại. Bạn có chắc chắn?', en: 'This will clear all current quotation data. Are you sure?' },
        confirmLangSwitch: { vi: 'Đổi ngôn ngữ sẽ xóa báo giá đang soạn. Bạn có chắc chắn?', en: 'Switching language will clear the quotation in progress. Are you sure?' },
        confirmDeleteService: { vi: 'Xóa phương án "{name}"?', en: 'Delete service "{name}"?' },
        confirmDeleteQuote: { vi: 'Bạn có chắc muốn xóa báo giá này không?', en: 'Are you sure you want to delete this quotation?' },
        confirmImport: { vi: 'Tìm thấy {count} phương án. Bạn có muốn thay thế danh sách hiện tại?', en: 'Found {count} services. Do you want to replace the current list?' },
        confirmImportSkipped: { vi: 'Tìm thấy {count} phương án hợp lệ ({skipped} dòng thiếu tên/đơn giá bị bỏ qua). Thay thế danh sách hiện tại?', en: 'Found {count} valid services ({skipped} rows with missing name/price skipped). Replace the current list?' },
        confirmRestore: { vi: 'Phục hồi dữ liệu sẽ ghi đè lên toàn bộ danh sách phương án và lịch sử báo giá hiện tại. Bạn có chắc chắn muốn tiếp tục?', en: 'Restoring data will overwrite all current services and quotation history. Are you sure you want to continue?' },
        toastReady: { vi: 'Sẵn sàng tạo báo giá', en: 'Ready to create quotations' },
        toastNewQuote: { vi: 'Đã tạo báo giá mới', en: 'New quotation created' },
        toastSaveSuccess: { vi: 'Đã lưu báo giá cho "{name}"!', en: 'Quotation saved for "{name}"!' },
        toastLoadSuccess: { vi: 'Đã tải báo giá của "{name}".', en: 'Loaded quotation for "{name}".' },
        toastDeleteSuccess: { vi: 'Đã xóa báo giá.', en: 'Quotation deleted.' },
        toastTariffUpdated: { vi: 'Danh sách phương án (Tariff) đã được cập nhật.', en: 'Service list (Tariff) has been updated.' },
        toastExportSuccess: { vi: 'Đã xuất file Excel thành công!', en: 'Excel file exported successfully!' },
        toastBackupSuccess: { vi: 'Đã sao lưu dữ liệu thành công!', en: 'Data backed up successfully!' },
        toastRestoreSuccess: { vi: 'Đã phục hồi dữ liệu thành công! Vui lòng tải lại trang.', en: 'Data restored successfully! Please reload the page.' },
        toastError: { vi: 'Có lỗi xảy ra. Vui lòng thử lại.', en: 'An error occurred. Please try again.' },
        toastValidationError: { vi: 'Vui lòng kiểm tra các trường sau: {fields}', en: 'Please check the following fields: {fields}' },
        toastNoCustomer: { vi: 'Tên khách hàng', en: 'Customer Name' },
        toastNoQuoteNo: { vi: 'Số báo giá', en: 'Quotation No.' },
        toastInvalidDate: { vi: 'Ngày báo giá', en: 'Quotation Date' },
        toastNoItems: { vi: 'Báo giá phải có ít nhất một phương án.', en: 'Quotation must have at least one service item.' },
        toastRowQty: { vi: 'Dòng {n}: số lượng phải lớn hơn 0', en: 'Row {n}: quantity must be greater than 0' },
        toastRowDates: { vi: 'Dòng {n}: chọn đủ thời gian bắt đầu - kết thúc hợp lệ', en: 'Row {n}: select a valid start - end period' },
        toastRowPrice: { vi: 'Dòng {n}: nhập đơn giá', en: 'Row {n}: enter a unit price' },
        toastImportNoData: { vi: 'Không tìm thấy dữ liệu trong file Excel.', en: 'No data found in the Excel file.' },
        toastImportBadFormat: { vi: 'File Excel cần có các cột: Name, Unit, Price (có dữ liệu hợp lệ).', en: 'The Excel file needs the columns: Name, Unit, Price (with valid data).' },
        toastInvalidFile: { vi: 'File không hợp lệ hoặc bị lỗi.', en: 'Invalid or corrupted file.' },
        toastDuplicateService: { vi: 'Phương án "{name}" đã tồn tại.', en: 'Service "{name}" already exists.' },
        toastInvalidService: { vi: 'Vui lòng nhập tên và đơn giá hợp lệ.', en: 'Please enter a valid name and price.' },
        toastQuoteNotFound: { vi: 'Không tìm thấy báo giá!', en: 'Quotation not found!' },
        toastLibMissing: { vi: 'Không tải được thư viện cần thiết. Kiểm tra kết nối Internet rồi tải lại trang.', en: 'A required library failed to load. Check your Internet connection and reload.' },
        toastFileProtocol: { vi: 'Không thể xuất khi mở trực tiếp file (file://). Hãy mở trang qua máy chủ web (http/https).', en: 'Export is not possible when opening the file directly (file://). Please serve the page over http/https.' },
        toastPreviewDone: { vi: 'Đã cập nhật bản xem trước.', en: 'Preview updated.' },
        toastPreviewError: { vi: 'Không tạo được bản xem trước. Xem Console (F12).', en: 'Error creating preview. Check Console (F12).' },
        toastPreparingPdf: { vi: 'Đang tạo PDF...', en: 'Preparing PDF...' },
        toastPdfReady: { vi: 'Đã tạo PDF.', en: 'PDF created.' },
        toastPdfDownloaded: { vi: 'Trình duyệt chặn cửa sổ mới - PDF đã được tải xuống.', en: 'Pop-up blocked - the PDF was downloaded instead.' },
        toastPdfError: { vi: 'Không xuất được PDF. Xem Console (F12).', en: 'Error creating PDF. Check Console (F12).' },
        processing: { vi: 'Đang xử lý...', en: 'Processing...' },
        historyNotFound: { vi: 'Chưa có dữ liệu.', en: 'No data yet.' },
        pdf: {
            companyName: { vi: 'TTĐH KHAI THÁC TÂN THUẬN', en: 'TAN THUAN OPERATIONS CENTER' },
            companyAddress: { vi: 'Địa chỉ: 18B Lưu Trọng Lư, Phường Tân Thuận, TP. HCM', en: 'Address: 18B Luu Trong Lu, Tan Thuan Ward, Ho Chi Minh City' },
            phone: { vi: 'Điện thoại', en: 'Phone' },
            address: { vi: 'Địa chỉ', en: 'Address' },
            email: { vi: 'Email', en: 'Email' },
            quoteNo: { vi: 'Số', en: 'No.' },
            date: { vi: 'Ngày', en: 'Date' },
            title: { vi: 'BÁO GIÁ', en: 'QUOTATION' },
            vatIncluded: { vi: '(Đã bao gồm 8% thuế VAT)', en: '(Includes 8% VAT)' },
            customerName: { vi: 'Tên khách hàng', en: 'Customer Name' },
            taxCode: { vi: 'Mã số thuế', en: 'Tax Code' },
            notes: { vi: 'Ghi chú', en: 'Notes' },
            thNo: { vi: 'STT', en: 'No.' },
            thService: { vi: 'Phương án', en: 'Service' },
            thUnit: { vi: 'ĐVT', en: 'Unit' },
            thQty: { vi: 'SL', en: 'QTY' },
            thUnitPrice: { vi: 'Đơn giá', en: 'Unit Price' },
            thAmount: { vi: 'Thành tiền', en: 'Amount' },
            totalPayment: { vi: 'Tổng cộng tiền thanh toán:', en: 'Total payment:' },
            inWords: { vi: 'Số tiền viết bằng chữ:', en: 'In words:' },
            paymentInfo: { vi: 'THÔNG TIN THANH TOÁN', en: 'PAYMENT INFORMATION' },
            bank: { vi: 'Ngân hàng', en: 'Bank' },
            accountName: { vi: 'Tên tài khoản', en: 'Account Name' },
            accountNumber: { vi: 'Số tài khoản', en: 'Account Number' },
            taxNote: { vi: 'QUÝ KHÁCH HÀNG VUI LÒNG CUNG CẤP MÃ SỐ THUẾ, ĐỊA CHỈ VÀ EMAIL ĐỂ NHẬN HÓA ĐƠN.', en: 'PLEASE PROVIDE YOUR TAX CODE, ADDRESS AND EMAIL TO RECEIVE THE INVOICE.' },
            otherOptionNote: { vi: '(Chi tiết trong Ghi chú)', en: '(Details in Notes)' },
            containerQtyDetail: { vi: '(SL: {containerQty} cont x {days} ngày)', en: '({containerQty} cont x {days} days)' },
            powerQtyDetail: { vi: '(SL: {containerQty} cont x {hours} giờ)', en: '({containerQty} cont x {hours} hrs)' },
            quoteNotesTitle: { vi: 'Lưu ý:', en: 'Notes:' },
            quoteNotesContent: {
                vi: `<li style="margin-bottom: 4px;">Chúng tôi báo giá theo thông tin khách hàng cung cấp, vui lòng kiểm tra lại phương án.</li><li style="margin-bottom: 4px;">Đối với phương án đóng/rút/sang container, chúng tôi không chịu trách nhiệm việc tháo gỡ, chằng buộc hàng hoá, cũng như các hư hỏng bên trong container như xước, gãy ván sàn...</li><li>Nếu phát sinh thêm phương án, sẽ căn cứ vào thực tế hiện trường để thu phí bổ sung.</li>`,
                en: `<li style="margin-bottom: 4px;">This quotation is based on the information provided by the customer. Please double-check the service details.</li><li style="margin-bottom: 4px;">For stuffing/unstuffing services, we are not responsible for unlashing/lashing cargo or any internal container damages such as scratches, broken floorboards, etc.</li><li>Additional services arising on-site will be charged based on the actual work performed.</li>`
            }
        }
    };

    const state = {
        currentLang: safeStorageGet('appLang') === 'en' ? 'en' : 'vi',
        products: [],
        rawTotalAmount: 0,
        pendingQuote: null,   // số báo giá đang gợi ý, chỉ "chốt" khi lưu
        datePicker: null
    };

    // --- 2. DOM ELEMENTS ---
    const elements = {
        langToggleBtn: document.getElementById('lang-toggle-btn'),
        appTitle: document.getElementById('app-title'),
        developedByLine: document.getElementById('developed-by-line'),
        tableBody: document.querySelector('#invoice-items-table tbody'),
        addRowBtn: document.getElementById('add-row-btn'),
        invoiceNumber: document.getElementById('invoice-number'),
        invoiceDate: document.getElementById('invoice-date'),
        customerName: document.getElementById('customer-name'),
        customerTaxCode: document.getElementById('customer-tax-code'),
        customerAddress: document.getElementById('customer-address'),
        customerPhone: document.getElementById('customer-phone'),
        customerNotes: document.getElementById('customer-notes'),
        totalAmount: document.getElementById('total-amount'),
        totalInWords: document.getElementById('total-in-words'),
        renderPreviewBtn: document.getElementById('render-preview-btn'),
        generatePdfBtn: document.getElementById('generate-pdf-btn'),
        newBtn: document.getElementById('new-btn'),
        saveBtn: document.getElementById('save-btn'),
        a4ContentWrapper: document.getElementById('a4-content-wrapper'),
        previewCanvas: document.getElementById('pdf-canvas-preview'),
        toast: document.getElementById('toast-notification'),
        productModal: document.getElementById('product-modal'),
        manageProductsBtn: document.getElementById('manage-products-btn'),
        closeProductModalBtn: document.getElementById('close-product-modal-btn'),
        productForm: document.getElementById('product-form'),
        productManagementTable: document.getElementById('product-management-table'),
        editIndex: document.getElementById('edit-index'),
        productName: document.getElementById('product-name'),
        productUnit: document.getElementById('product-unit'),
        productPrice: document.getElementById('product-price'),
        cancelEditBtn: document.getElementById('cancel-edit-btn'),
        importExcelBtn: document.getElementById('import-excel-btn'),
        exportExcelBtn: document.getElementById('export-excel-btn'),
        excelImporter: document.getElementById('excel-importer'),
        backupDataBtn: document.getElementById('backup-data-btn'),
        restoreDataBtn: document.getElementById('restore-data-btn'),
        jsonImporter: document.getElementById('json-importer'),
        historyModal: document.getElementById('history-modal'),
        historyBtn: document.getElementById('history-btn'),
        closeHistoryModalBtn: document.getElementById('close-history-modal-btn'),
        historyTableBody: document.getElementById('history-table-body'),
        formTitle: document.getElementById('form-title'),
        labelQuoteNo: document.getElementById('label-quote-no'),
        labelQuoteDate: document.getElementById('label-quote-date'),
        labelCustomerName: document.getElementById('label-customer-name'),
        labelTaxCode: document.getElementById('label-tax-code'),
        labelAddress: document.getElementById('label-address'),
        labelPhone: document.getElementById('label-phone'),
        labelNotes: document.getElementById('label-notes'),
        thService: document.getElementById('th-service'),
        thQty: document.getElementById('th-qty'),
        thUnitPrice: document.getElementById('th-unit-price'),
        thAmount: document.getElementById('th-amount'),
        labelTotal: document.getElementById('label-total'),
        labelInWords: document.getElementById('label-in-words'),
        actionsTitle: document.getElementById('actions-title'),
        previewPlaceholder: document.getElementById('preview-placeholder'),
        modalProductTitle: document.getElementById('modal-product-title'),
        labelModalName: document.getElementById('label-modal-name'),
        labelModalUnit: document.getElementById('label-modal-unit'),
        labelModalPrice: document.getElementById('label-modal-price'),
        btnModalSave: document.getElementById('btn-modal-save'),
        thModalName: document.getElementById('th-modal-name'),
        thModalUnit: document.getElementById('th-modal-unit'),
        thModalPrice: document.getElementById('th-modal-price'),
        thModalAction: document.getElementById('th-modal-action'),
        modalHistoryTitle: document.getElementById('modal-history-title'),
        thHistoryDate: document.getElementById('th-history-date'),
        thHistoryNo: document.getElementById('th-history-no'),
        thHistoryCustomer: document.getElementById('th-history-customer'),
        thHistoryTotal: document.getElementById('th-history-total'),
        thHistoryAction: document.getElementById('th-history-action'),
    };

    // --- 3. UTILITY FUNCTIONS ---
    /** Lấy chuỗi dịch theo ngôn ngữ hiện tại, thay {biến}. */
    function t(key, vars) {
        let s = i18n[key][state.currentLang];
        if (vars) Object.keys(vars).forEach(k => { s = s.split(`{${k}}`).join(vars[k]); });
        return s;
    }
    const tp = (key) => i18n.pdf[key][state.currentLang];

    /** Escape ký tự HTML để chèn an toàn vào innerHTML / thuộc tính. */
    function escapeHtml(value) {
        return String(value == null ? '' : value).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    }

    function safeStorageGet(key) { try { return localStorage.getItem(key); } catch (e) { return null; } }
    function safeStorageSet(key, value) { try { localStorage.setItem(key, value); return true; } catch (e) { console.warn('localStorage write failed', e); return false; } }
    function readJSON(key, fallback) {
        try {
            const raw = safeStorageGet(key);
            return raw ? JSON.parse(raw) : fallback;
        } catch (e) { return fallback; }
    }

    const formatCurrency = (amount) => new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(Number.isFinite(amount) ? amount : 0);

    let toastTimer = null;
    const showToast = (message, type = 'success') => {
        elements.toast.textContent = message;
        elements.toast.className = `show ${type}`;
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => { elements.toast.classList.remove('show'); }, 3500);
    };
    const showModal = (modalElement) => { modalElement.style.display = 'block'; };
    const hideModal = (modalElement) => { modalElement.style.display = 'none'; };

    /** Chuyển chuỗi DD/MM/YYYY thành Date. Trả về null nếu sai định dạng hoặc ngày không tồn tại (vd 32/13/2026). */
    function parseDateString(dateString) {
        const m = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.exec(String(dateString || '').trim());
        if (!m) return null;
        const d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]);
        const date = new Date(y, mo - 1, d);
        if (date.getFullYear() !== y || date.getMonth() !== mo - 1 || date.getDate() !== d) return null;
        return date;
    }

    /** Đọc số tiền bằng chữ tiếng Việt (làm tròn đến đồng). */
    function numberToWords_VI(n) {
        n = Math.round(Number(n));
        if (!Number.isFinite(n) || n === 0) return 'Không';
        const negative = n < 0;
        n = Math.abs(n);
        const digits = ['không', 'một', 'hai', 'ba', 'bốn', 'năm', 'sáu', 'bảy', 'tám', 'chín'];
        const scales = ['', 'ngàn', 'triệu', 'tỷ', 'ngàn tỷ', 'triệu tỷ'];

        // full = true: nhóm này đứng sau một nhóm lớn hơn nên phải đọc đủ "không trăm", "lẻ"
        function readTriple(num, full) {
            const h = Math.floor(num / 100), t10 = Math.floor((num % 100) / 10), u = num % 10;
            const parts = [];
            if (h > 0 || full) parts.push(digits[h] + ' trăm');
            if (t10 > 1) {
                parts.push(digits[t10] + ' mươi');
                if (u === 1) parts.push('mốt');
                else if (u === 5) parts.push('lăm');
                else if (u > 0) parts.push(digits[u]);
            } else if (t10 === 1) {
                parts.push('mười');
                if (u === 5) parts.push('lăm');
                else if (u > 0) parts.push(digits[u]);
            } else if (u > 0) {
                if (h > 0 || full) parts.push('lẻ');
                parts.push(digits[u]);
            }
            return parts.join(' ');
        }

        const groups = [];
        while (n > 0) { groups.push(n % 1000); n = Math.floor(n / 1000); }
        const words = [];
        for (let i = groups.length - 1; i >= 0; i--) {
            if (groups[i] === 0) continue;
            const full = i < groups.length - 1;
            words.push((readTriple(groups[i], full) + ' ' + (scales[i] || '')).trim());
        }
        const text = (negative ? 'âm ' : '') + words.join(' ');
        return text.charAt(0).toUpperCase() + text.slice(1);
    }

    /** Đọc số tiền bằng chữ tiếng Anh (làm tròn đến đồng). */
    function numberToWords_EN(n) {
        n = Math.round(Number(n));
        if (!Number.isFinite(n) || n === 0) return 'Zero';
        const negative = n < 0;
        n = Math.abs(n);
        const below20 = ['', 'One', 'Two', 'Three', 'Four', 'Five', 'Six', 'Seven', 'Eight', 'Nine', 'Ten', 'Eleven', 'Twelve', 'Thirteen', 'Fourteen', 'Fifteen', 'Sixteen', 'Seventeen', 'Eighteen', 'Nineteen'];
        const tens = ['', '', 'Twenty', 'Thirty', 'Forty', 'Fifty', 'Sixty', 'Seventy', 'Eighty', 'Ninety'];
        const scales = ['', 'Thousand', 'Million', 'Billion', 'Trillion', 'Quadrillion'];
        function helper(num) {
            if (num === 0) return '';
            if (num < 20) return below20[num] + ' ';
            if (num < 100) return tens[Math.floor(num / 10)] + ' ' + helper(num % 10);
            return below20[Math.floor(num / 100)] + ' Hundred ' + helper(num % 100);
        }
        let word = '', i = 0;
        while (n > 0) {
            if (n % 1000 !== 0) word = helper(n % 1000) + scales[i] + ' ' + word;
            n = Math.floor(n / 1000);
            i++;
        }
        return (negative ? 'Minus ' : '') + word.replace(/\s+/g, ' ').trim();
    }

    // --- 4. PRODUCT / ROW HELPERS ---
    const normText = (s) => String(s || '').normalize('NFC').toLowerCase();

    /** 'storage' | 'power' | 'normal' - quyết định cách tính số lượng của dòng. */
    function getServiceKind(product) {
        if (!product || product.isOther) return 'normal';
        const name = normText(product.name), unit = normText(product.unit);
        if (name.includes(normText(i18n.storageKeywords[state.currentLang])) || /ngày|day/.test(unit)) return 'storage';
        if (name.includes(normText(i18n.powerKeywords[state.currentLang])) || /giờ|hrs|hour/.test(unit)) return 'power';
        return 'normal';
    }

    function buildProductOptionsHtml(legacyProduct) {
        let html = `<option value="">${escapeHtml(t('selectOption'))}</option>`;
        state.products.forEach((p, i) => { html += `<option value="${i}">${escapeHtml(p.name)}</option>`; });
        if (legacyProduct) html += `<option value="legacy">${escapeHtml(legacyProduct.name)}</option>`;
        return html;
    }

    /** Sản phẩm đang chọn của một dòng (theo vị trí trong state.products, hoặc "legacy" nếu đã bị xóa khỏi tariff). */
    function getRowProduct(row) {
        const v = row.querySelector('.product-select').value;
        if (v === '') return null;
        if (v === 'legacy') return row._legacyProduct || null;
        return state.products[Number(v)] || null;
    }

    function getRowContainerQty(row) {
        return Math.max(1, parseFloat(row.querySelector('.container-qty-input').value) || 1);
    }

    // --- 5. UI UPDATE FUNCTIONS ---
    function createTableRow() {
        const row = document.createElement('tr');
        row.className = 'item-row';
        row.dataset.kind = 'normal';

        row.innerHTML = `
            <td class="px-4 py-2 align-top">
                <select class="product-select mt-1 block w-full rounded-md border-gray-300 shadow-sm sm:text-sm">
                    ${buildProductOptionsHtml(null)}
                </select>

                <div class="calculation-picker container-qty-picker">
                    <label>${escapeHtml(t('containerQty'))}</label>
                    <input type="number" class="container-qty-input mt-1 block w-full rounded-md border-gray-300 shadow-sm sm:text-sm" value="1" min="1" step="1">
                </div>

                <div class="calculation-picker date-range-picker">
                    <label>${escapeHtml(t('dateRangeFrom'))}</label>
                    <input type="text" class="start-date-input" placeholder="${escapeHtml(t('pickStartDate'))}">
                    <label>${escapeHtml(t('dateRangeTo'))}</label>
                    <input type="text" class="end-date-input" placeholder="${escapeHtml(t('pickEndDate'))}">
                </div>

                <div class="calculation-picker datetime-range-picker">
                    <label>${escapeHtml(t('dateTimeRangeFrom'))}</label>
                    <input type="text" class="start-datetime-input" placeholder="${escapeHtml(t('pickStartDateTime'))}">
                    <label>${escapeHtml(t('dateTimeRangeTo'))}</label>
                    <input type="text" class="end-datetime-input" placeholder="${escapeHtml(t('pickEndDateTime'))}">
                </div>
            </td>
            <td class="px-4 py-2 align-top"><input type="number" class="qty-input mt-1 block w-full rounded-md border-gray-300 shadow-sm sm:text-sm" value="1" min="0" step="0.5"></td>
            <td class="px-4 py-2 align-top"><input type="number" class="price-input mt-1 block w-full rounded-md border-gray-300 shadow-sm sm:text-sm bg-gray-100" value="0" min="0" readonly></td>
            <td class="px-4 py-2 align-top"><span class="line-total font-bold text-gray-900">0 đ</span></td>
            <td class="px-2 py-2 align-top text-center"><button type="button" class="remove-row-btn text-red-500 hover:text-red-700 font-bold" aria-label="Remove">×</button></td>
        `;
        elements.tableBody.appendChild(row);
        initFlatpickr(row);
        return row;
    }

    function initFlatpickr(rowElement) {
        const fire = (selectedDates, dateStr, instance) => {
            instance.element.dispatchEvent(new Event('input', { bubbles: true }));
        };
        const dateOnlyOptions = { dateFormat: 'd/m/Y', onChange: fire };
        const dateTimeOptions = { enableTime: true, dateFormat: 'd/m/Y H:i', time_24hr: true, minuteIncrement: 30, onChange: fire };

        rowElement.querySelectorAll('.start-date-input, .end-date-input').forEach(input => flatpickr(input, dateOnlyOptions));
        rowElement.querySelectorAll('.start-datetime-input, .end-datetime-input').forEach(input => flatpickr(input, dateTimeOptions));
    }

    /** Áp dụng sản phẩm đã chọn lên dòng: hiện/ẩn bộ chọn ngày, khóa/mở ô số lượng và đơn giá. */
    function applyRowProduct(row, { keepPrice = false } = {}) {
        const product = getRowProduct(row);
        const kind = getServiceKind(product);
        const prevKind = row.dataset.kind || 'normal';
        row.dataset.kind = kind;

        const qtyInput = row.querySelector('.qty-input');
        const priceInput = row.querySelector('.price-input');

        // Lưu bãi / chạy điện: có số lượng container; thời gian (ngày/giờ) tự tính từ bộ chọn
        row.querySelector('.container-qty-picker').classList.toggle('visible', kind !== 'normal');
        row.querySelector('.date-range-picker').classList.toggle('visible', kind === 'storage');
        row.querySelector('.datetime-range-picker').classList.toggle('visible', kind === 'power');

        const qtyLocked = kind !== 'normal';
        qtyInput.readOnly = qtyLocked;
        qtyInput.classList.toggle('bg-gray-100', qtyLocked);
        if (!qtyLocked && prevKind !== 'normal') qtyInput.value = 1;

        const priceEditable = !!(product && product.isOther);
        priceInput.readOnly = !priceEditable;
        priceInput.classList.toggle('bg-gray-100', !priceEditable);
        if (!keepPrice) priceInput.value = product ? (priceEditable ? 0 : product.price) : 0;

        if (qtyLocked) updateRowQuantity(row);
    }

    /** Tính số ngày ("lưu bãi") hoặc số giờ ("chạy điện") của dòng. Chưa đủ / sai khoảng thời gian -> 0. */
    function updateRowQuantity(row) {
        const kind = row.dataset.kind;
        const qtyInput = row.querySelector('.qty-input');

        if (kind === 'storage') {
            const fpStart = row.querySelector('.start-date-input')._flatpickr;
            const fpEnd = row.querySelector('.end-date-input')._flatpickr;
            let days = 0;
            if (fpStart && fpEnd && fpStart.selectedDates.length && fpEnd.selectedDates.length) {
                const startDate = new Date(fpStart.selectedDates[0]); startDate.setHours(0, 0, 0, 0);
                const endDate = new Date(fpEnd.selectedDates[0]); endDate.setHours(0, 0, 0, 0);
                if (endDate >= startDate) days = Math.round((endDate.getTime() - startDate.getTime()) / 86400000) + 1;
            }
            qtyInput.value = days;
        } else if (kind === 'power') {
            const fpStart = row.querySelector('.start-datetime-input')._flatpickr;
            const fpEnd = row.querySelector('.end-datetime-input')._flatpickr;
            let hours = 0;
            if (fpStart && fpEnd && fpStart.selectedDates.length && fpEnd.selectedDates.length) {
                const diffMinutes = (fpEnd.selectedDates[0].getTime() - fpStart.selectedDates[0].getTime()) / 60000;
                if (diffMinutes >= 0) {
                    const totalHours = Math.round((diffMinutes / 60) * 2) / 2;
                    hours = totalHours > 0 ? totalHours : 0.5;
                }
            }
            qtyInput.value = hours;
        }
    }

    function recalculateTotals() {
        let grandTotal = 0;
        elements.tableBody.querySelectorAll('.item-row').forEach(row => {
            const price = parseFloat(row.querySelector('.price-input').value) || 0;
            const qty = parseFloat(row.querySelector('.qty-input').value) || 0;
            // Lưu bãi / chạy điện: thành tiền = số container x số ngày (giờ) x đơn giá
            const multiplier = row.dataset.kind !== 'normal' ? getRowContainerQty(row) : 1;
            const lineTotal = Math.round(multiplier * qty * price);
            row.querySelector('.line-total').textContent = formatCurrency(lineTotal);
            grandTotal += lineTotal;
        });
        state.rawTotalAmount = grandTotal;
        elements.totalAmount.textContent = formatCurrency(grandTotal);
        const words = state.currentLang === 'vi' ? numberToWords_VI(grandTotal) : numberToWords_EN(grandTotal);
        elements.totalInWords.textContent = words + ' ' + t('dong') + '.';
    }

    /** Dựng lại danh sách chọn của mọi dòng sau khi tariff thay đổi (thêm/sửa/xóa/import). */
    function refreshRowSelects() {
        elements.tableBody.querySelectorAll('.item-row').forEach(row => {
            const prev = getRowProduct(row);
            const wasEmpty = !prev;
            let legacy = null;
            const idx = prev ? state.products.indexOf(prev) : -1;
            if (prev && idx === -1) { legacy = prev; row._legacyProduct = legacy; }
            else if (!prev) { row._legacyProduct = null; }
            const select = row.querySelector('.product-select');
            select.innerHTML = buildProductOptionsHtml(legacy);
            select.value = wasEmpty ? '' : (idx !== -1 ? String(idx) : 'legacy');
            applyRowProduct(row, { keepPrice: !!(prev && prev.isOther) || !!legacy });
        });
        recalculateTotals();
    }

    function formHasData() {
        const fields = [elements.customerName, elements.customerTaxCode, elements.customerAddress, elements.customerPhone, elements.customerNotes];
        if (fields.some(el => el.value.trim() !== '')) return true;
        return Array.from(elements.tableBody.querySelectorAll('.item-row')).some(row => row.querySelector('.product-select').value !== '');
    }

    /** Xóa form, đặt ngày hôm nay, gợi ý số báo giá mới (chưa chốt), tạo 1 dòng trống. */
    function startNewQuote() {
        elements.customerName.value = '';
        elements.customerTaxCode.value = '';
        elements.customerAddress.value = '';
        elements.customerPhone.value = '';
        elements.customerNotes.value = '';
        state.datePicker.setDate(new Date(), false);
        suggestQuoteNumber();
        elements.tableBody.innerHTML = '';
        createTableRow();
        recalculateTotals();
    }

    function resetForm(confirmNeeded = true) {
        const doReset = () => {
            startNewQuote();
            showToast(t('toastNewQuote'), 'info');
        };
        if (confirmNeeded) {
            if (confirm(t('confirmClear'))) doReset();
        } else {
            doReset();
        }
    }

    function renderProductManagementTable() {
        elements.productManagementTable.innerHTML = '';
        state.products.forEach((p, index) => {
            if (p.isOther) return;
            const row = elements.productManagementTable.insertRow();
            row.innerHTML = `<td class="p-2">${escapeHtml(p.name)}</td><td class="p-2 text-center">${escapeHtml(p.unit)}</td><td class="p-2 text-right">${escapeHtml(formatCurrency(p.price))}</td><td class="p-2 text-center"><button type="button" class="edit-product-btn px-2 py-1 bg-yellow-500 text-white rounded text-sm" data-index="${index}">${escapeHtml(t('edit'))}</button> <button type="button" class="delete-product-btn px-2 py-1 bg-red-600 text-white rounded text-sm" data-index="${index}">${escapeHtml(t('delete'))}</button></td>`;
        });
    }

    function renderHistoryTable() {
        const savedQuotations = getSavedQuotations();
        savedQuotations.sort((a, b) => String(b.id).localeCompare(String(a.id)));
        elements.historyTableBody.innerHTML = savedQuotations.length === 0
            ? `<tr><td colspan="5" class="text-center p-4">${escapeHtml(t('historyNotFound'))}</td></tr>` : '';
        savedQuotations.forEach(q => {
            const row = elements.historyTableBody.insertRow();
            row.innerHTML = `<td class="p-2">${escapeHtml(q.invoiceDate)}</td><td class="p-2">${escapeHtml(q.invoiceNumber)}</td><td class="p-2">${escapeHtml(q.customerName)}</td><td class="p-2 text-right">${escapeHtml(q.totalAmount)}</td><td class="p-2 text-center"><button type="button" class="view-history-btn px-2 py-1 bg-blue-500 text-white rounded text-sm" data-id="${escapeHtml(q.id)}">${escapeHtml(t('view'))}</button> <button type="button" class="delete-history-btn px-2 py-1 bg-red-600 text-white rounded text-sm" data-id="${escapeHtml(q.id)}">${escapeHtml(t('delete'))}</button></td>`;
        });
    }

    function resetProductForm() {
        elements.productForm.reset();
        elements.editIndex.value = -1;
        elements.cancelEditBtn.classList.add('hidden');
    }

    function translateUI() {
        document.documentElement.lang = state.currentLang;
        elements.appTitle.textContent = t('appTitle');
        elements.developedByLine.textContent = t('developedBy');
        elements.manageProductsBtn.textContent = t('manageServices');
        elements.historyBtn.textContent = t('history');
        elements.formTitle.textContent = t('formTitle');
        elements.labelQuoteNo.textContent = t('quoteNo');
        elements.labelQuoteDate.textContent = t('quoteDate');
        elements.labelCustomerName.textContent = t('customerName');
        elements.labelTaxCode.textContent = t('taxCode');
        elements.labelAddress.textContent = t('address');
        elements.labelPhone.textContent = t('phone');
        elements.labelNotes.textContent = t('notes');
        elements.thService.textContent = t('thService');
        elements.thQty.textContent = t('thQty');
        elements.thUnitPrice.textContent = t('thUnitPrice');
        elements.thAmount.textContent = t('thAmount');
        elements.addRowBtn.textContent = t('addRow');
        elements.labelTotal.textContent = t('total');
        elements.labelInWords.textContent = t('inWords');
        elements.actionsTitle.textContent = t('actionsTitle');
        elements.newBtn.textContent = t('newQuote');
        elements.saveBtn.textContent = t('saveQuote');
        elements.renderPreviewBtn.textContent = t('updatePreview');
        elements.generatePdfBtn.textContent = t('printPdf');
        elements.previewPlaceholder.textContent = t('previewPlaceholder');
        elements.modalProductTitle.textContent = t('modalProductTitle');
        elements.importExcelBtn.textContent = t('importExcel');
        elements.exportExcelBtn.textContent = t('exportExcel');
        elements.backupDataBtn.textContent = t('backupData');
        elements.restoreDataBtn.textContent = t('restoreData');
        elements.labelModalName.textContent = t('modalName');
        elements.labelModalUnit.textContent = t('modalUnit');
        elements.labelModalPrice.textContent = t('modalPrice');
        elements.btnModalSave.textContent = t('btnModalSave');
        elements.cancelEditBtn.textContent = t('btnCancel');
        elements.thModalName.textContent = t('modalName');
        elements.thModalUnit.textContent = t('modalUnit');
        elements.thModalPrice.textContent = t('modalPrice');
        elements.thModalAction.textContent = t('action');
        elements.modalHistoryTitle.textContent = t('modalHistoryTitle');
        elements.thHistoryDate.textContent = t('date');
        elements.thHistoryNo.textContent = t('quoteNo');
        elements.thHistoryCustomer.textContent = t('customer');
        elements.thHistoryTotal.textContent = t('total');
        elements.thHistoryAction.textContent = t('action');
        recalculateTotals();
    }

    // --- 6. DATA MANAGEMENT (LocalStorage) ---
    function saveProducts() {
        const toSave = state.products.filter(p => !p.isOther).map(p => ({ name: p.name, unit: p.unit, price: p.price }));
        safeStorageSet(`tanthuan_products_${state.currentLang}`, JSON.stringify(toSave));
    }

    function loadProducts() {
        const stored = readJSON(`tanthuan_products_${state.currentLang}`, null);
        // Danh sách rỗng (vd. từ file sao lưu khi chưa từng chỉnh tariff) -> dùng biểu giá mặc định
        const source = (Array.isArray(stored) && stored.length) ? stored : (DEFAULT_PRODUCTS[state.currentLang] || []);
        state.products = source
            .filter(p => p && p.name)
            .map(p => ({ name: String(p.name), unit: String(p.unit || 'cont'), price: Number(p.price) || 0 }));
        state.products.push({ name: t('otherOption'), unit: state.currentLang === 'vi' ? 'cont' : 'Turn', price: 0, isOther: true });
    }

    function getSavedQuotations() {
        const list = readJSON(`tanthuan_quotations_${state.currentLang}`, []);
        return Array.isArray(list) ? list : [];
    }
    function setSavedQuotations(list) {
        return safeStorageSet(`tanthuan_quotations_${state.currentLang}`, JSON.stringify(list));
    }

    function validateForm() {
        const errors = [];
        const fieldsToValidate = [
            { element: elements.customerName, errorKey: 'toastNoCustomer' },
            { element: elements.invoiceNumber, errorKey: 'toastNoQuoteNo' },
            { element: elements.invoiceDate, errorKey: 'toastInvalidDate', validator: (el) => !!parseDateString(el.value) }
        ];

        fieldsToValidate.forEach(field => field.element.classList.remove('input-error'));
        elements.tableBody.querySelectorAll('.input-error').forEach(el => el.classList.remove('input-error'));

        fieldsToValidate.forEach(field => {
            const isValid = field.validator ? field.validator(field.element) : field.element.value.trim() !== '';
            if (!isValid) {
                errors.push(t(field.errorKey));
                field.element.classList.add('input-error');
            }
        });

        const rows = Array.from(elements.tableBody.querySelectorAll('.item-row'));
        let hasItems = false;
        rows.forEach((row, i) => {
            const product = getRowProduct(row);
            if (!product) return;
            hasItems = true;
            const n = i + 1;
            const kind = getServiceKind(product);
            const qty = parseFloat(row.querySelector('.qty-input').value) || 0;
            const price = parseFloat(row.querySelector('.price-input').value) || 0;
            if (qty <= 0) {
                errors.push(kind === 'normal' ? t('toastRowQty', { n }) : t('toastRowDates', { n }));
                row.querySelector(kind === 'normal' ? '.qty-input' : '.product-select').classList.add('input-error');
            }
            if (product.isOther && price <= 0) {
                errors.push(t('toastRowPrice', { n }));
                row.querySelector('.price-input').classList.add('input-error');
            }
        });
        if (!hasItems) errors.push(t('toastNoItems'));

        if (errors.length > 0) {
            showToast(t('toastValidationError', { fields: errors.join(', ') }), 'error');
            return false;
        }
        return true;
    }

    function collectItems() {
        const items = [];
        elements.tableBody.querySelectorAll('.item-row').forEach(row => {
            const product = getRowProduct(row);
            if (!product) return;
            const val = (sel) => { const el = row.querySelector(sel); return el ? el.value : ''; };
            items.push({
                name: product.name,
                unit: product.unit,
                quantity: parseFloat(val('.qty-input')) || 0,
                price: parseFloat(val('.price-input')) || 0,
                containerQty: parseInt(val('.container-qty-input'), 10) || 1,
                startDate: val('.start-date-input'),
                endDate: val('.end-date-input'),
                startDateTime: val('.start-datetime-input'),
                endDateTime: val('.end-datetime-input')
            });
        });
        return items;
    }

    function saveQuotation() {
        if (!validateForm()) return;

        const quotation = {
            id: `QT-${Date.now()}`,
            invoiceNumber: elements.invoiceNumber.value.trim(),
            invoiceDate: elements.invoiceDate.value,
            customerName: elements.customerName.value.trim(),
            customerTaxCode: elements.customerTaxCode.value.trim(),
            customerAddress: elements.customerAddress.value.trim(),
            customerPhone: elements.customerPhone.value.trim(),
            customerNotes: elements.customerNotes.value.trim(),
            totalAmount: elements.totalAmount.textContent,
            items: collectItems()
        };

        const saved = getSavedQuotations();
        // Lưu lại cùng số báo giá -> cập nhật bản cũ thay vì tạo bản trùng
        const existing = saved.findIndex(q => q.invoiceNumber === quotation.invoiceNumber);
        if (existing !== -1) { quotation.id = saved[existing].id; saved[existing] = quotation; }
        else saved.push(quotation);

        if (!setSavedQuotations(saved)) { showToast(t('toastError'), 'error'); return; }
        commitQuoteNumber();
        showToast(t('toastSaveSuccess', { name: quotation.customerName }));
    }

    function loadQuotation(id) {
        const q = getSavedQuotations().find(x => x.id === id);
        if (!q) { showToast(t('toastQuoteNotFound'), 'error'); return; }

        elements.invoiceNumber.value = q.invoiceNumber || '';
        state.datePicker.setDate(q.invoiceDate || new Date(), false, 'd/m/Y');
        elements.customerName.value = q.customerName || '';
        elements.customerTaxCode.value = q.customerTaxCode || '';
        elements.customerAddress.value = q.customerAddress || '';
        elements.customerPhone.value = q.customerPhone || '';
        elements.customerNotes.value = q.customerNotes || '';
        elements.tableBody.innerHTML = '';

        (q.items || []).forEach(item => {
            const row = createTableRow();
            const select = row.querySelector('.product-select');
            const otherName = t('otherOption');
            const price = Number(item.price) || 0;

            let idx = -1;
            if (item.name === otherName) idx = state.products.findIndex(p => p.isOther);
            else {
                idx = state.products.findIndex(p => !p.isOther && p.name === item.name && p.price === price);
                if (idx === -1) idx = state.products.findIndex(p => !p.isOther && p.name === item.name);
            }
            if (idx === -1) {
                // Phương án đã bị xóa khỏi tariff: giữ nguyên dòng báo giá cũ
                row._legacyProduct = { name: item.name, unit: item.unit || 'cont', price };
                select.innerHTML = buildProductOptionsHtml(row._legacyProduct);
                select.value = 'legacy';
            } else {
                select.value = String(idx);
            }
            applyRowProduct(row, { keepPrice: true });

            // Giá tại thời điểm báo giá được giữ nguyên, không lấy theo tariff hiện tại
            row.querySelector('.price-input').value = price;
            const setPicker = (sel, value) => {
                const input = row.querySelector(sel);
                if (value && input && input._flatpickr) input._flatpickr.setDate(value, false);
            };
            setPicker('.start-date-input', item.startDate);
            setPicker('.end-date-input', item.endDate);
            setPicker('.start-datetime-input', item.startDateTime);
            setPicker('.end-datetime-input', item.endDateTime);
            row.querySelector('.container-qty-input').value = item.containerQty || 1;
            row.querySelector('.qty-input').value = item.quantity;
        });
        if (!elements.tableBody.querySelector('.item-row')) createTableRow();
        state.pendingQuote = null;
        recalculateTotals();
        showToast(t('toastLoadSuccess', { name: q.customerName }));
    }

    function deleteQuotation(id) {
        setSavedQuotations(getSavedQuotations().filter(q => q.id !== id));
        showToast(t('toastDeleteSuccess'), 'info');
    }

    // --- Số báo giá: chỉ gợi ý khi mở/ tạo mới, chỉ "chốt" (tăng bộ đếm) khi lưu ---
    function peekNextQuoteNumber() {
        const today = new Date();
        const dateStr = `${String(today.getDate()).padStart(2, '0')}${String(today.getMonth() + 1).padStart(2, '0')}`;
        const info = readJSON('lastQuoteInfo', { date: '', counter: 0 });
        const counter = (info && info.date === dateStr) ? (Number(info.counter) || 0) + 1 : 1;
        const prefix = state.currentLang === 'vi' ? 'BG' : 'QT';
        return { number: `${prefix}${dateStr}-${String(counter).padStart(3, '0')}`, dateStr, counter };
    }

    function suggestQuoteNumber() {
        state.pendingQuote = peekNextQuoteNumber();
        elements.invoiceNumber.value = state.pendingQuote.number;
    }

    function commitQuoteNumber() {
        const p = state.pendingQuote;
        if (p && elements.invoiceNumber.value.trim() === p.number) {
            safeStorageSet('lastQuoteInfo', JSON.stringify({ date: p.dateStr, counter: p.counter }));
            state.pendingQuote = null;
        }
    }

    // --- 7. QR THANH TOÁN (VietQR / EMVCo, tạo ngay trên trình duyệt) ---
    function emvTlv(id, value) { return id + String(value.length).padStart(2, '0') + value; }

    function crc16ccitt(str) {
        let crc = 0xFFFF;
        for (let i = 0; i < str.length; i++) {
            crc ^= str.charCodeAt(i) << 8;
            for (let b = 0; b < 8; b++) crc = (crc & 0x8000) ? ((crc << 1) ^ 0x1021) & 0xFFFF : (crc << 1) & 0xFFFF;
        }
        return crc.toString(16).toUpperCase().padStart(4, '0');
    }

    function buildVietQrPayload(amount, description) {
        const beneficiary = emvTlv('00', BANK.bin) + emvTlv('01', BANK.account);
        const merchant = emvTlv('00', 'A000000727') + emvTlv('01', beneficiary) + emvTlv('02', 'QRIBFTTA');
        const hasAmount = amount > 0;
        let payload = emvTlv('00', '01') + emvTlv('01', hasAmount ? '12' : '11') + emvTlv('38', merchant) + emvTlv('53', '704');
        if (hasAmount) payload += emvTlv('54', String(Math.round(amount)));
        payload += emvTlv('58', 'VN');
        const info = String(description || '').normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[đĐ]/g, 'd')
            .replace(/[^A-Za-z0-9 ._-]/g, '').slice(0, 25);
        if (info) payload += emvTlv('62', emvTlv('08', info));
        payload += '6304';
        return payload + crc16ccitt(payload);
    }

    /** Trả về data URL của mã QR, hoặc '' nếu thư viện QR không có. */
    function getQrCodeDataUrl(amount, description) {
        if (typeof qrcode === 'undefined') return '';
        try {
            const qr = qrcode(0, 'M');
            qr.addData(buildVietQrPayload(amount, description));
            qr.make();
            return qr.createDataURL(8, 4);
        } catch (e) {
            console.warn('Could not generate QR code', e);
            return '';
        }
    }

    // --- 8. PDF & PREVIEW GENERATION ---
    /** Chờ ảnh tải xong. Ảnh lỗi/quá chậm không được làm treo hoặc làm hỏng cả bản PDF. */
    function waitForImageLoad(img) {
        return new Promise(resolve => {
            if (img.complete) return resolve();
            const done = () => resolve();
            img.addEventListener('load', done, { once: true });
            img.addEventListener('error', done, { once: true });
            setTimeout(done, 8000);
        });
    }

    function generateA4HTML(qrDataUrl) {
        let itemsHTML = '';
        let stt = 0;
        elements.tableBody.querySelectorAll('.item-row').forEach(row => {
            const product = getRowProduct(row);
            if (!product) return;
            stt++;

            const kind = getServiceKind(product);
            const qty = row.querySelector('.qty-input').value;
            const price = parseFloat(row.querySelector('.price-input').value) || 0;
            const lineTotal = row.querySelector('.line-total').textContent;
            const containerQty = getRowContainerQty(row);

            let displayName = escapeHtml(product.name);
            if (product.isOther) displayName += ` ${escapeHtml(tp('otherOptionNote'))}`;
            if (kind === 'storage') {
                displayName += ' ' + escapeHtml(tp('containerQtyDetail').replace('{containerQty}', containerQty).replace('{days}', qty));
            } else if (kind === 'power') {
                displayName += ' ' + escapeHtml(tp('powerQtyDetail').replace('{containerQty}', containerQty).replace('{hours}', qty));
            }

            itemsHTML += `<tr class="border-b"><td style="padding: 8px; border: 1px solid #ccc; text-align: center;">${stt}</td><td style="padding: 8px; border: 1px solid #ccc; word-break: break-word; white-space: normal;">${displayName}</td><td style="padding: 8px; border: 1px solid #ccc; text-align: center;">${escapeHtml(product.unit)}</td><td style="padding: 8px; border: 1px solid #ccc; text-align: center;">${escapeHtml(qty)}</td><td style="padding: 8px; border: 1px solid #ccc; text-align: right;">${escapeHtml(formatCurrency(price))}</td><td style="padding: 8px; border: 1px solid #ccc; text-align: right;">${escapeHtml(lineTotal)}</td></tr>`;
        });

        const taxCode = elements.customerTaxCode.value.trim();
        const taxCodeHTML = taxCode
            ? `<tr><td style="border: none; padding: 2px 0; width: 15%; font-weight: bold; white-space: nowrap; vertical-align: top;">${tp('taxCode')}:</td><td style="border: none; padding: 2px 0;">${escapeHtml(taxCode)}</td></tr>`
            : '';
        const notesHtml = escapeHtml(elements.customerNotes.value).replace(/\n/g, '<br>');
        const qrHtml = qrDataUrl
            ? `<img src="${qrDataUrl}" alt="QR Code" style="width: 120px; height: 120px; display: inline-block;" id="qr-code-image">`
            : '';
        const logoSrc = typeof LOGO_DATA_URI !== 'undefined' ? LOGO_DATA_URI : 'logo.png';

        return `
            <div style="font-family: 'Roboto', sans-serif; color: #000;">
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 1.5rem;"><tr><td style="width: 25%; vertical-align: middle; padding-right: 15px;"><img src="${logoSrc}" alt="Logo" style="max-width: 100px; height: auto;"></td><td style="width: 45%; vertical-align: middle;"><h3 style="font-size: 1.125rem; font-weight: 700; margin:0 0 5px 0;">${tp('companyName')}</h3><p style="margin: 0; font-size: 10pt;">${tp('phone')}: 02838728546</p><p style="margin: 0; font-size: 10pt;">${tp('companyAddress')}</p><p style="margin: 0; font-size: 10pt;">${tp('email')}: cms.cont@tanthuanport.vn</p></td><td style="width: 30%; vertical-align: middle; text-align: right;"><p style="margin: 0; font-size: 11pt;"><strong>${tp('quoteNo')}:</strong> ${escapeHtml(elements.invoiceNumber.value)}</p><p style="margin: 0; font-size: 11pt;"><strong>${tp('date')}:</strong> ${escapeHtml(elements.invoiceDate.value)}</p></td></tr></table>
                <div style="text-align: center; margin-bottom: 1.5rem;"><h1 style="font-size: 2.25rem; font-weight: 700; margin-bottom: 0;">${tp('title')}</h1><p style="font-size: 11pt; font-style: italic; margin-top: 4px;">${tp('vatIncluded')}</p></div>
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 1.2rem; font-size: 11pt;"><tr><td style="border: none; padding: 2px 0; width: 15%; font-weight: bold; white-space: nowrap; vertical-align: top;">${tp('customerName')}:</td><td style="border: none; padding: 2px 0;">${escapeHtml(elements.customerName.value)}</td></tr>${taxCodeHTML}<tr><td style="border: none; padding: 2px 0; width: 15%; font-weight: bold; white-space: nowrap; vertical-align: top;">${tp('address')}:</td><td style="border: none; padding: 2px 0;">${escapeHtml(elements.customerAddress.value)}</td></tr><tr><td style="border: none; padding: 2px 0; width: 15%; font-weight: bold; white-space: nowrap; vertical-align: top;">${tp('phone')}:</td><td style="border: none; padding: 2px 0;">${escapeHtml(elements.customerPhone.value)}</td></tr><tr><td style="border: none; padding: 2px 0; width: 15%; font-weight: bold; white-space: nowrap; vertical-align: top;">${tp('notes')}:</td><td style="border: none; padding: 2px 0; word-break: break-word; white-space: normal;">${notesHtml}</td></tr></table>
                <table style="width: 100%; border-collapse: collapse; margin-bottom: 1rem; font-size: 11pt;"><thead style="background-color: #e0e0e0 !important; -webkit-print-color-adjust: exact;"><tr><th style="width: 5%; text-align: center; padding: 8px; border: 1px solid #ccc;">${tp('thNo')}</th><th style="width: 45%; text-align: left; padding: 8px; border: 1px solid #ccc;">${tp('thService')}</th><th style="width: 10%; text-align: center; padding: 8px; border: 1px solid #ccc;">${tp('thUnit')}</th><th style="width: 10%; text-align: center; padding: 8px; border: 1px solid #ccc;">${tp('thQty')}</th><th style="width: 15%; text-align: right; padding: 8px; border: 1px solid #ccc;">${tp('thUnitPrice')}</th><th style="width: 15%; text-align: right; padding: 8px; border: 1px solid #ccc;">${tp('thAmount')}</th></tr></thead><tbody>${itemsHTML}</tbody></table>
                <table style="width: 100%; border-collapse: collapse; margin-top: 20px; border: 1px solid #ccc; font-size: 11pt;"><tr><td style="padding: 8px; width: 70%; border-right: 1px solid #ccc;">${tp('totalPayment')}</td><td style="padding: 8px; width: 30%; text-align: right; font-weight: bold;">${escapeHtml(elements.totalAmount.textContent)}</td></tr><tr style="background-color: #f3f4f6 !important; -webkit-print-color-adjust: exact;"><td colspan="2" style="padding: 8px;"><span style="font-style: normal;">${tp('inWords')} </span><span style="font-style: italic;">${escapeHtml(elements.totalInWords.textContent)}</span></td></tr></table>
                <div style="margin-top: 20px; font-size: 10pt; border: 1px solid #ddd; padding: 10px 15px; border-radius: 5px; background-color: #f9f9f9 !important; -webkit-print-color-adjust: exact;"><h4 style="font-weight: bold; margin: 0 0 8px 0; font-size: 11pt;">${tp('quoteNotesTitle')}</h4><ul class="custom-list">${tp('quoteNotesContent')}</ul></div>
                <table style="width: 100%; border-collapse: collapse; margin-top: 1.5rem; border-top: 1px solid #ccc; padding-top: 1rem;"><tr><td style="width: 70%; border: none; vertical-align: top;"><h4 style="font-weight: 700; font-size: 1rem; margin-bottom: 0.5rem;">${tp('paymentInfo')}</h4><p style="margin: 0; font-size: 11pt;">${tp('bank')}: <strong>BIDV - Chau Thanh Sai Gon Branch</strong></p><p style="margin: 0; font-size: 11pt;">${tp('accountName')}: <strong>${BANK.accountName}</strong></p><p style="margin: 0; font-size: 11pt;">${tp('accountNumber')}: <strong style="color: #d0021b !important;">${BANK.account}</strong></p><p style="margin-top: 2rem; font-size: 0.875rem; font-style: italic;">${tp('taxNote')}</p></td><td style="width: 30%; border: none; text-align: right; vertical-align: top;">${qrHtml}</td></tr></table>
            </div>
        `;
    }

    function assertLibraries() {
        if (typeof html2canvas === 'undefined' || !window.jspdf) {
            const err = new Error('LIB_MISSING');
            err.code = 'LIB_MISSING';
            throw err;
        }
    }

    /** Dựng nội dung A4 ẩn rồi chụp thành canvas. */
    async function renderQuoteCanvas(scale) {
        assertLibraries();
        const qrDataUrl = getQrCodeDataUrl(state.rawTotalAmount, elements.invoiceNumber.value);
        elements.a4ContentWrapper.innerHTML = generateA4HTML(qrDataUrl);

        await new Promise(resolve => setTimeout(resolve, 0));
        const images = Array.from(elements.a4ContentWrapper.querySelectorAll('img')).map(waitForImageLoad);
        await Promise.all([...images, document.fonts ? document.fonts.ready : Promise.resolve()]);

        return html2canvas(elements.a4ContentWrapper, {
            scale,
            useCORS: true,
            backgroundColor: '#ffffff',
            scrollX: 0,
            scrollY: 0,
            windowWidth: elements.a4ContentWrapper.scrollWidth,
            windowHeight: elements.a4ContentWrapper.scrollHeight
        });
    }

    /**
     * Tạo PDF A4 nhiều trang từ canvas. Điểm ngắt trang được chọn ở ranh giới giữa các dòng/khối
     * để không cắt ngang chữ; không kéo giãn ảnh.
     */
    function buildPdfFromCanvas(canvas, wrapper) {
        const { jsPDF } = window.jspdf;
        const doc = new jsPDF({ orientation: 'p', unit: 'mm', format: 'a4' });
        const PAGE_W = 210, PAGE_H = 297, MARGIN_TOP = 12, MARGIN_BOTTOM = 12;

        const wrapRect = wrapper.getBoundingClientRect();
        const cssW = wrapRect.width;
        const scale = canvas.width / cssW;
        const mmToCss = cssW / PAGE_W;
        const wrapTop = wrapRect.top;

        // Chiều cao nội dung thật (bỏ phần min-height trống của khung A4) + lề dưới của khung
        const inner = wrapper.firstElementChild;
        const paddingBottom = parseFloat(getComputedStyle(wrapper).paddingBottom) || 0;
        const contentBottom = inner ? inner.getBoundingClientRect().bottom - wrapTop + paddingBottom : canvas.height / scale;
        const totalCss = Math.min(canvas.height / scale, contentBottom);

        const breakPoints = Array.from(wrapper.querySelectorAll(':scope > div > *, tr'))
            .map(el => el.getBoundingClientRect().bottom - wrapTop)
            .filter(v => v > 0)
            .sort((a, b) => a - b);

        const segments = [];
        let start = 0;
        if (totalCss <= PAGE_H * mmToCss + 1) {
            segments.push([0, totalCss]);   // vừa một trang A4
            start = totalCss;
        }
        while (start < totalCss - 1) {
            const avail = (segments.length === 0 ? PAGE_H - MARGIN_BOTTOM : PAGE_H - MARGIN_TOP - MARGIN_BOTTOM) * mmToCss;
            if (start + avail >= totalCss - 1) { segments.push([start, totalCss]); break; }
            const limit = start + avail;
            let end = limit;
            for (let i = breakPoints.length - 1; i >= 0; i--) {
                if (breakPoints[i] <= limit && breakPoints[i] > start + avail * 0.4) { end = breakPoints[i]; break; }
            }
            segments.push([start, end]);
            start = end;
        }

        segments.forEach(([s, e], i) => {
            if (i > 0) doc.addPage();
            const sy = Math.round(s * scale);
            const sh = Math.min(canvas.height - sy, Math.round((e - s) * scale));
            if (sh <= 0) return;
            const part = document.createElement('canvas');
            part.width = canvas.width;
            part.height = sh;
            const ctx = part.getContext('2d');
            ctx.fillStyle = '#ffffff';
            ctx.fillRect(0, 0, part.width, part.height);
            ctx.drawImage(canvas, 0, sy, canvas.width, sh, 0, 0, canvas.width, sh);
            const heightMm = sh * PAGE_W / canvas.width;
            doc.addImage(part.toDataURL('image/jpeg', 0.92), 'JPEG', 0, i === 0 ? 0 : MARGIN_TOP, PAGE_W, heightMm, undefined, 'FAST');
            if (segments.length > 1) {
                doc.setFontSize(9);
                doc.setTextColor(120);
                doc.text(`${i + 1}/${segments.length}`, PAGE_W / 2, PAGE_H - 5, { align: 'center' });
            }
        });
        return doc;
    }

    function reportExportError(error, fallbackKey) {
        console.error(error);
        if (error && error.code === 'LIB_MISSING') showToast(t('toastLibMissing'), 'error');
        else if (location.protocol === 'file:') showToast(t('toastFileProtocol'), 'error');
        else showToast(t(fallbackKey), 'error');
    }

    async function renderPreviewToCanvas() {
        if (!validateForm()) return;
        showToast(t('updatePreview'), 'info');
        try {
            const canvas = await renderQuoteCanvas(1.5);
            elements.previewCanvas.width = canvas.width;
            elements.previewCanvas.height = canvas.height;
            elements.previewCanvas.style.height = 'auto';
            const ctx = elements.previewCanvas.getContext('2d');
            ctx.clearRect(0, 0, canvas.width, canvas.height);
            ctx.drawImage(canvas, 0, 0);
            showToast(t('toastPreviewDone'));
        } catch (error) {
            reportExportError(error, 'toastPreviewError');
        }
    }

    async function generatePdf() {
        if (!validateForm()) return;

        // Mở tab ngay trong thao tác click (trước các bước bất đồng bộ) để trình duyệt không chặn pop-up
        let pdfWindow = null;
        try { pdfWindow = window.open('', '_blank'); } catch (e) { pdfWindow = null; }
        if (pdfWindow) {
            try { pdfWindow.document.title = elements.invoiceNumber.value; pdfWindow.document.body.textContent = t('toastPreparingPdf'); } catch (e) { /* ignore */ }
        }

        elements.generatePdfBtn.disabled = true;
        elements.generatePdfBtn.textContent = t('processing');
        showToast(t('toastPreparingPdf'), 'info');

        try {
            const canvas = await renderQuoteCanvas(2);
            const doc = buildPdfFromCanvas(canvas, elements.a4ContentWrapper);
            const blob = doc.output('blob');
            const url = URL.createObjectURL(blob);
            const filename = `${(elements.invoiceNumber.value || 'BaoGia').replace(/[^\w.-]+/g, '_')}.pdf`;

            if (pdfWindow && !pdfWindow.closed) {
                pdfWindow.location.href = url;
                showToast(t('toastPdfReady'));
            } else {
                const a = document.createElement('a');
                a.href = url;
                a.download = filename;
                document.body.appendChild(a);
                a.click();
                a.remove();
                showToast(t('toastPdfDownloaded'), 'info');
            }
            setTimeout(() => URL.revokeObjectURL(url), 5 * 60 * 1000);
        } catch (error) {
            if (pdfWindow && !pdfWindow.closed) pdfWindow.close();
            reportExportError(error, 'toastPdfError');
        } finally {
            elements.generatePdfBtn.disabled = false;
            elements.generatePdfBtn.textContent = t('printPdf');
        }
    }

    // --- 9. EVENT HANDLERS ---
    function handleTableRowInput(e) {
        const row = e.target.closest('.item-row');
        if (!row) return;
        const cls = e.target.classList;
        if (cls.contains('start-date-input') || cls.contains('end-date-input')
            || cls.contains('start-datetime-input') || cls.contains('end-datetime-input')) {
            updateRowQuantity(row);
        }
        recalculateTotals();
    }

    function handleTableRowChange(e) {
        if (!e.target.classList.contains('product-select')) return;
        const row = e.target.closest('.item-row');
        applyRowProduct(row);
        const product = getRowProduct(row);
        if (product && product.isOther) row.querySelector('.price-input').focus();
        recalculateTotals();
    }

    function handleTableRowClick(e) {
        if (e.target.classList.contains('remove-row-btn')) {
            e.target.closest('.item-row').remove();
            recalculateTotals();
        }
    }

    function handleLangToggle() {
        if (formHasData() && !confirm(t('confirmLangSwitch'))) return;
        state.currentLang = state.currentLang === 'vi' ? 'en' : 'vi';
        safeStorageSet('appLang', state.currentLang);
        loadProducts();
        translateUI();
        startNewQuote();
    }

    function handleProductFormSubmit(e) {
        e.preventDefault();
        const name = elements.productName.value.trim().replace(/\s+/g, ' ');
        const unit = elements.productUnit.value.trim() || 'cont';
        const price = parseFloat(elements.productPrice.value);
        const index = parseInt(elements.editIndex.value, 10);

        if (!name || !Number.isFinite(price) || price < 0) { showToast(t('toastInvalidService'), 'error'); return; }
        const dup = state.products.findIndex((p, i) => i !== index && normText(p.name).trim() === normText(name));
        if (dup !== -1) { showToast(t('toastDuplicateService', { name }), 'error'); return; }

        if (index === -1) {
            state.products.unshift({ name, unit, price });
        } else {
            Object.assign(state.products[index], { name, unit, price });   // giữ nguyên tham chiếu để các dòng đang chọn vẫn khớp
        }

        saveProducts();
        renderProductManagementTable();
        refreshRowSelects();
        showToast(t('toastTariffUpdated'));
        resetProductForm();
    }

    function handleProductTableClick(e) {
        const raw = e.target.dataset.index;
        if (raw === undefined) return;
        const index = Number(raw);
        const product = state.products[index];
        if (!product) return;

        if (e.target.classList.contains('delete-product-btn')) {
            if (confirm(t('confirmDeleteService', { name: product.name }))) {
                state.products.splice(index, 1);
                saveProducts();
                renderProductManagementTable();
                refreshRowSelects();
                showToast(t('toastTariffUpdated'));
            }
        }
        if (e.target.classList.contains('edit-product-btn')) {
            elements.productName.value = product.name;
            elements.productUnit.value = product.unit;
            elements.productPrice.value = product.price;
            elements.editIndex.value = index;
            elements.cancelEditBtn.classList.remove('hidden');
        }
    }

    function handleExcelImport(event) {
        const file = event.target.files[0];
        if (!file) return;
        if (typeof XLSX === 'undefined') { showToast(t('toastLibMissing'), 'error'); event.target.value = ''; return; }
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
                const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
                if (!json || json.length === 0) {
                    showToast(t('toastImportNoData'), 'error');
                    return;
                }
                const valid = [];
                let skipped = 0;
                json.forEach(item => {
                    const name = String(item['Name'] == null ? '' : item['Name']).trim().replace(/\s+/g, ' ');
                    const unit = String(item['Unit'] == null ? '' : item['Unit']).trim() || 'cont';
                    const price = parseFloat(item['Price']);
                    if (!name || !Number.isFinite(price) || price < 0) { skipped++; return; }
                    valid.push({ name, unit, price });
                });
                if (valid.length === 0) {
                    showToast(t('toastImportBadFormat'), 'error');
                    return;
                }
                const message = skipped > 0
                    ? t('confirmImportSkipped', { count: valid.length, skipped })
                    : t('confirmImport', { count: valid.length });
                if (confirm(message)) {
                    state.products = valid;
                    state.products.push({ name: t('otherOption'), unit: state.currentLang === 'vi' ? 'cont' : 'Turn', price: 0, isOther: true });
                    saveProducts();
                    renderProductManagementTable();
                    refreshRowSelects();
                    showToast(t('toastTariffUpdated'));
                }
            } catch (error) {
                showToast(t('toastError'), 'error');
                console.error('Excel import error:', error);
            } finally {
                event.target.value = '';
            }
        };
        reader.readAsArrayBuffer(file);
    }

    function handleExcelExport() {
        if (typeof XLSX === 'undefined') { showToast(t('toastLibMissing'), 'error'); return; }
        const dataToExport = state.products.filter(p => !p.isOther).map(p => ({ 'Name': p.name, 'Unit': p.unit, 'Price': p.price }));
        const worksheet = XLSX.utils.json_to_sheet(dataToExport);
        const workbook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(workbook, worksheet, 'Services');
        XLSX.writeFile(workbook, 'Service_List_TanThuan.xlsx');
        showToast(t('toastExportSuccess'));
    }

    function handleHistoryTableClick(e) {
        const id = e.target.dataset.id;
        if (!id) return;
        if (e.target.classList.contains('view-history-btn')) {
            loadQuotation(id);
            hideModal(elements.historyModal);
        }
        if (e.target.classList.contains('delete-history-btn')) {
            if (confirm(t('confirmDeleteQuote'))) {
                deleteQuotation(id);
                renderHistoryTable();
            }
        }
    }

    function handleWindowClick(event) {
        if (event.target === elements.productModal) { hideModal(elements.productModal); resetProductForm(); }
        if (event.target === elements.historyModal) { hideModal(elements.historyModal); }
    }

    function backupData() {
        const dataToBackup = {
            products_vi: readJSON('tanthuan_products_vi', []),
            products_en: readJSON('tanthuan_products_en', []),
            quotations_vi: readJSON('tanthuan_quotations_vi', []),
            quotations_en: readJSON('tanthuan_quotations_en', []),
            lastQuoteInfo: readJSON('lastQuoteInfo', {})
        };
        const blob = new Blob([JSON.stringify(dataToBackup, null, 2)], { type: 'application/json' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `tanthuan_backup_${new Date().toISOString().slice(0, 10)}.json`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
        showToast(t('toastBackupSuccess'));
    }

    function handleRestoreData(event) {
        const file = event.target.files[0];
        if (!file) return;
        if (!confirm(t('confirmRestore'))) {
            event.target.value = '';
            return;
        }
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = JSON.parse(e.target.result);
                if (!data || !Array.isArray(data.products_vi) || !Array.isArray(data.quotations_vi)) {
                    throw new Error('Invalid backup file structure.');
                }
                safeStorageSet('tanthuan_products_vi', JSON.stringify(data.products_vi));
                safeStorageSet('tanthuan_products_en', JSON.stringify(Array.isArray(data.products_en) ? data.products_en : []));
                safeStorageSet('tanthuan_quotations_vi', JSON.stringify(data.quotations_vi));
                safeStorageSet('tanthuan_quotations_en', JSON.stringify(Array.isArray(data.quotations_en) ? data.quotations_en : []));
                safeStorageSet('lastQuoteInfo', JSON.stringify(data.lastQuoteInfo || {}));
                showToast(t('toastRestoreSuccess'), 'success');
                setTimeout(() => window.location.reload(), 2000);
            } catch (error) {
                showToast(t('toastInvalidFile'), 'error');
                console.error('Restore error:', error);
            } finally {
                event.target.value = '';
            }
        };
        reader.readAsText(file);
    }

    // --- 10. INITIALIZATION ---
    function bindEvents() {
        elements.langToggleBtn.addEventListener('click', handleLangToggle);
        elements.addRowBtn.addEventListener('click', () => createTableRow());
        elements.tableBody.addEventListener('click', handleTableRowClick);
        elements.tableBody.addEventListener('input', handleTableRowInput);
        elements.tableBody.addEventListener('change', handleTableRowChange);
        elements.newBtn.addEventListener('click', () => resetForm(true));
        elements.saveBtn.addEventListener('click', saveQuotation);
        elements.renderPreviewBtn.addEventListener('click', renderPreviewToCanvas);
        elements.generatePdfBtn.addEventListener('click', generatePdf);
        elements.manageProductsBtn.onclick = () => { showModal(elements.productModal); renderProductManagementTable(); };
        elements.closeProductModalBtn.onclick = () => { hideModal(elements.productModal); resetProductForm(); };
        elements.cancelEditBtn.onclick = resetProductForm;
        elements.productForm.addEventListener('submit', handleProductFormSubmit);
        elements.productManagementTable.addEventListener('click', handleProductTableClick);
        elements.importExcelBtn.onclick = () => elements.excelImporter.click();
        elements.exportExcelBtn.addEventListener('click', handleExcelExport);
        elements.excelImporter.addEventListener('change', handleExcelImport);
        elements.backupDataBtn.addEventListener('click', backupData);
        elements.restoreDataBtn.onclick = () => elements.jsonImporter.click();
        elements.jsonImporter.addEventListener('change', handleRestoreData);
        elements.historyBtn.onclick = () => { showModal(elements.historyModal); renderHistoryTable(); };
        elements.closeHistoryModalBtn.onclick = () => { hideModal(elements.historyModal); };
        elements.historyTableBody.addEventListener('click', handleHistoryTableClick);
        window.addEventListener('click', handleWindowClick);
    }

    function init() {
        state.datePicker = flatpickr(elements.invoiceDate, { dateFormat: 'd/m/Y', defaultDate: 'today' });
        loadProducts();
        translateUI();
        startNewQuote();
        showToast(t('toastReady'));
    }

    // --- APP ENTRY POINT ---
    init();
    bindEvents();

})(); // End of IIFE
