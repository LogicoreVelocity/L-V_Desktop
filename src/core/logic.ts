import ExcelJS from 'exceljs'

type PageKey =
  | 'dashboard'
  | 'employees'
  | 'finance'
  | 'assets'
  | 'documents'
  | 'templates'
  | 'inputs'
  | 'deductions'
  | 'payroll'
  | 'wps'
  | 'slips'
  | 'backup'
  | 'settings'
  | 'about'

type EmployeeStatus = 'Active' | 'Inactive'

type Employee = {
  id: string
  employeeCode: string
  payrollCode: string
  wpsPersonId: string
  wpsRoutingCode: string
  wpsAccountNumber: string
  employeeName: string
  roleDesignation: string
  dateOfJoining: string
  mobileNumber: string
  basicSalary: number
  status: EmployeeStatus | string
  email: string
  emiratesId: string
  emiratesIdExpiry: string
  visaExpiryDate: string
  passportNumber: string
  passportExpiry: string
  emergencyContactName: string
  emergencyContactMobile: string
  emergencyContactName2: string
  emergencyContactMobile2: string
  licenseNumber: string
  licenseExpiry: string
  permitNumber: string
  permitExpiry: string
  workHours: string
  salaryType: 'Fixed Salary' | 'Per Day'
  accommodationAllowance: number
  travelAllowance: number
  incentive: number
  dueFromLastMonth: number
  daysPresent: number
  workingDays: number
  allowance: number
  remarks: string
}

type Deduction = {
  id: string
  employeeId: string
  month: string
  year: string
  advanceDeduction: number
  fineDeduction: number
  salikDeduction: number
  otherDeduction: number
  reasonRemarks: string
  referenceNo: string
  transactionId: string
  fineDate: string
  fineReceiptNo: string
  fineTransactionNo: string
  salikDate: string
  salikReceiptNo: string
  salikTransactionNo: string
  attachmentName: string
}

type PayrollInput = {
  id: string
  employeeId: string
  month: string
  year: string
  daysPresent: number
  workingDays: number
  allowance: number
  incentive: number
  attendanceBonus: number
  performanceBonus: number
  overtimePay: number
  otherEarning: number
  arrears: number
  remarks: string
}

type FinanceRecord = {
  id: string
  date: string
  type: 'Income' | 'Expense' | 'Transfer' | string
  category: string
  party: string
  referenceNo: string
  description: string
  amount: number
  vatAmount: number
  paymentMode: string
  status: 'Pending' | 'Paid' | 'Reconciled' | string
  attachmentName: string
}

type AssetRecord = {
  id: string
  assetCode: string
  assetName: string
  category: string
  assignedToEmployeeId: string
  purchaseDate: string
  purchaseCost: number
  serialNumber: string
  location: string
  condition: 'Good' | 'Service Due' | 'Damaged' | 'Retired' | string
  warrantyExpiry: string
  status: 'Active' | 'Maintenance' | 'Disposed' | string
  notes: string
}

type DocumentRecord = {
  id: string
  documentDate: string
  documentType: string
  linkedModule: string
  linkedReference: string
  title: string
  fileName: string
  fileSizeKb: number
  notes: string
}

type Settings = {
  companyName: string
  companyAddress: string
  companyPhone: string
  companyEmail: string
  wpsEmployerId: string
  wpsRoutingCode: string
  wpsEmployerReference: string
  defaultPayrollMonth: string
  defaultPayrollYear: string
  currency: string
}

type StoredState = {
  employees: Employee[]
  financeRecords: FinanceRecord[]
  assetRecords: AssetRecord[]
  documentRecords: DocumentRecord[]
  payrollInputs: PayrollInput[]
  deductions: Deduction[]
  settings: Settings
}

type PayrollRow = Employee & {
  advanceDeduction: number
  fineDeduction: number
  salikDeduction: number
  otherDeduction: number
  attendanceBonus: number
  performanceBonus: number
  overtimePay: number
  otherEarning: number
  arrears: number
  salaryEarned: number
  grossSalary: number
  payableBeforeDeductions: number
  totalDeduction: number
  netSalary: number
  month: string
  year: string
}

type WpsValidation = {
  errors: string[]
  warnings: string[]
}

type SifRecordSet = {
  filename: string
  content: string
  rows: {
    employeeCode: string
    employeeName: string
    fixedIncome: number
    variableIncome: number
    leaveDays: number
    totalPay: number
  }[]
  validation: WpsValidation
}

type Notice = {
  tone: 'success' | 'error' | 'warning'
  message: string
  details: string[]
}

const storageKey = 'logicore-excel-manager-v2'

const now = new Date()
const currentMonth = String(now.getMonth() + 1).padStart(2, '0')
const currentYear = String(now.getFullYear())

const defaultSettings: Settings = {
  companyName: 'Your Company',
  companyAddress: '',
  companyPhone: '',
  companyEmail: '',
  wpsEmployerId: '',
  wpsRoutingCode: '',
  wpsEmployerReference: '',
  defaultPayrollMonth: currentMonth,
  defaultPayrollYear: currentYear,
  currency: 'AED',
}

const pages: { key: PageKey; label: string }[] = [
  { key: 'dashboard', label: 'Dashboard' },
  { key: 'employees', label: 'Employees / Drivers' },
  { key: 'finance', label: 'Finance Register' },
  { key: 'assets', label: 'Asset Register' },
  { key: 'documents', label: 'Documents / PDFs' },
  { key: 'templates', label: 'Import Templates' },
  { key: 'inputs', label: 'Payroll Inputs' },
  { key: 'deductions', label: 'Advances & Deductions' },
  { key: 'payroll', label: 'Payroll' },
  { key: 'wps', label: 'WPS / SIF' },
  { key: 'slips', label: 'Salary Slips' },
  { key: 'backup', label: 'Backup / Restore' },
  { key: 'settings', label: 'Settings' },
  { key: 'about', label: 'Notes' },
]

const employeeFields: {
  key: keyof Employee
  label: string
  type: 'text' | 'number' | 'date' | 'email' | 'select'
  required: boolean
}[] = [
  { key: 'employeeCode', label: 'Employee Code', type: 'text', required: true },
  { key: 'payrollCode', label: 'Payroll Code', type: 'text', required: false },
  { key: 'wpsPersonId', label: 'WPS Person ID', type: 'text', required: false },
  { key: 'wpsRoutingCode', label: 'WPS Agent Routing Code', type: 'text', required: false },
  { key: 'wpsAccountNumber', label: 'WPS IBAN / Account', type: 'text', required: false },
  { key: 'employeeName', label: 'Employee Name', type: 'text', required: true },
  {
    key: 'roleDesignation',
    label: 'Role / Designation',
    type: 'text',
    required: true,
  },
  { key: 'dateOfJoining', label: 'Date of Joining', type: 'date', required: true },
  { key: 'mobileNumber', label: 'Mobile Number', type: 'text', required: true },
  { key: 'basicSalary', label: 'Basic Salary', type: 'number', required: true },
  { key: 'status', label: 'Status', type: 'select', required: true },
  { key: 'email', label: 'Email', type: 'email', required: false },
  { key: 'emiratesId', label: 'Emirates ID', type: 'text', required: false },
  { key: 'emiratesIdExpiry', label: 'Emirates ID Expiry', type: 'date', required: false },
  { key: 'visaExpiryDate', label: 'Visa Expiry Date', type: 'date', required: false },
  { key: 'passportNumber', label: 'Passport Number', type: 'text', required: false },
  { key: 'passportExpiry', label: 'Passport Expiry', type: 'date', required: false },
  {
    key: 'emergencyContactName',
    label: 'Emergency Contact Name',
    type: 'text',
    required: false,
  },
  {
    key: 'emergencyContactMobile',
    label: 'Emergency Contact Mobile',
    type: 'text',
    required: false,
  },
  {
    key: 'emergencyContactName2',
    label: 'Emergency Contact Name 2',
    type: 'text',
    required: false,
  },
  {
    key: 'emergencyContactMobile2',
    label: 'Emergency Contact Mobile 2',
    type: 'text',
    required: false,
  },
  { key: 'licenseNumber', label: 'License Number', type: 'text', required: false },
  { key: 'licenseExpiry', label: 'License Expiry', type: 'date', required: false },
  { key: 'permitNumber', label: 'Permit Number', type: 'text', required: false },
  { key: 'permitExpiry', label: 'Permit Expiry', type: 'date', required: false },
  { key: 'workHours', label: 'Work Hours', type: 'text', required: false },
  { key: 'salaryType', label: 'Salary Type', type: 'select', required: false },
  { key: 'daysPresent', label: 'Days Present', type: 'number', required: false },
  { key: 'workingDays', label: 'Working Days', type: 'number', required: false },
  {
    key: 'accommodationAllowance',
    label: 'Accommodation Allowance',
    type: 'number',
    required: false,
  },
  {
    key: 'travelAllowance',
    label: 'Travel Allowance',
    type: 'number',
    required: false,
  },
  { key: 'incentive', label: 'Incentive', type: 'number', required: false },
  { key: 'dueFromLastMonth', label: 'Due From Last Month', type: 'number', required: false },
  { key: 'allowance', label: 'Allowance', type: 'number', required: false },
  { key: 'remarks', label: 'Remarks', type: 'text', required: false },
]

const employeeTemplateColumns = employeeFields.map((field) => field.label)

const deductionsTemplateColumns = [
  'Invoice No',
  'Invoice Type',
  'Date',
  'Customer',
  'Driver',
  'Agreement No',
  'Vehicle',
  'Amount',
]

const payrollInputTemplateColumns = [
  'Employee Code',
  'Payroll Code',
  'Employee Name',
  'Month',
  'Year',
  'Days Present',
  'Working Days',
  'Allowance',
  'Incentive',
  'Attendance Bonus',
  'Performance Bonus / Tip',
  'Overtime Pay',
  'Other Earnings',
  'Arrears / Due From Last Month',
  'Remarks',
]

const financeTemplateColumns = [
  'Date',
  'Type',
  'Category',
  'Party',
  'Reference No',
  'Description',
  'Amount',
  'VAT Amount',
  'Payment Mode',
  'Status',
  'Attachment Name',
]

const assetTemplateColumns = [
  'Asset Code',
  'Asset Name',
  'Category',
  'Assigned Employee Code',
  'Purchase Date',
  'Purchase Cost',
  'Serial Number',
  'Location',
  'Condition',
  'Status',
  'Warranty Expiry',
  'Notes',
]

const documentTemplateColumns = [
  'Document Date',
  'Document Type',
  'Linked Module',
  'Linked Reference',
  'Title',
  'File Name',
  'File Size KB',
  'Notes',
]

const payrollExportColumns = [
  'Sl',
  'Visa',
  'Name',
  'Project',
  'Project ID',
  'Monthly Rental DDN or Salary',
  'Total No. Days',
  'Salary Type',
  'Earnings',
  'Allowance',
  'Performance Bonus / Tip',
  'Attendance Bonus',
  'Incentive',
  'O/T',
  'Other Earnings',
  'Arrears',
  'LOP',
  'Eligible Days',
  'Gross Earning',
  'COD',
  'Client Penalty',
  'Vehicle Rent',
  'Advance Deduction',
  'Salik Deduction',
  'Other Deduction',
  'Net Salary',
  'Employee Code',
  'Payroll Code',
  'Remarks',
]

const maxExcelFileSizeBytes = 10 * 1024 * 1024

const monthOptions = [
  { value: '01', label: 'January' },
  { value: '02', label: 'February' },
  { value: '03', label: 'March' },
  { value: '04', label: 'April' },
  { value: '05', label: 'May' },
  { value: '06', label: 'June' },
  { value: '07', label: 'July' },
  { value: '08', label: 'August' },
  { value: '09', label: 'September' },
  { value: '10', label: 'October' },
  { value: '11', label: 'November' },
  { value: '12', label: 'December' },
]

function blankEmployee(): Employee {
  return {
    id: crypto.randomUUID(),
    employeeCode: '',
    payrollCode: '',
    wpsPersonId: '',
    wpsRoutingCode: '',
    wpsAccountNumber: '',
    employeeName: '',
    roleDesignation: '',
    dateOfJoining: '',
    mobileNumber: '',
    basicSalary: 0,
    status: 'Active',
    email: '',
    emiratesId: '',
    emiratesIdExpiry: '',
    visaExpiryDate: '',
    passportNumber: '',
    passportExpiry: '',
    emergencyContactName: '',
    emergencyContactMobile: '',
    emergencyContactName2: '',
    emergencyContactMobile2: '',
    licenseNumber: '',
    licenseExpiry: '',
    permitNumber: '',
    permitExpiry: '',
    workHours: '',
    salaryType: 'Fixed Salary',
    daysPresent: 0,
    workingDays: 0,
    accommodationAllowance: 0,
    travelAllowance: 0,
    incentive: 0,
    dueFromLastMonth: 0,
    allowance: 0,
    remarks: '',
  }
}

function normalizeStoredEmployee(employee: Partial<Employee>): Employee {
  return {
    ...blankEmployee(),
    ...employee,
    id: employee.id || crypto.randomUUID(),
    employeeCode: text(employee.employeeCode),
    payrollCode: text(employee.payrollCode),
    wpsPersonId: text(employee.wpsPersonId),
    wpsRoutingCode: text(employee.wpsRoutingCode),
    wpsAccountNumber: text(employee.wpsAccountNumber),
    status: text(employee.status) || 'Active',
    salaryType: employee.salaryType === 'Per Day' ? 'Per Day' : 'Fixed Salary',
    basicSalary: numberValue(employee.basicSalary),
    accommodationAllowance: numberValue(employee.accommodationAllowance),
    travelAllowance: numberValue(employee.travelAllowance),
    incentive: numberValue(employee.incentive),
    dueFromLastMonth: numberValue(employee.dueFromLastMonth),
    daysPresent: numberValue(employee.daysPresent),
    workingDays: numberValue(employee.workingDays),
    allowance: numberValue(employee.allowance),
  }
}

function blankDeduction(month: string, year: string, employeeId = ''): Deduction {
  return {
    id: crypto.randomUUID(),
    employeeId,
    month,
    year,
    advanceDeduction: 0,
    fineDeduction: 0,
    salikDeduction: 0,
    otherDeduction: 0,
    reasonRemarks: '',
    referenceNo: '',
    transactionId: '',
    fineDate: '',
    fineReceiptNo: '',
    fineTransactionNo: '',
    salikDate: '',
    salikReceiptNo: '',
    salikTransactionNo: '',
    attachmentName: '',
  }
}

function blankPayrollInput(month: string, year: string, employeeId = ''): PayrollInput {
  return {
    id: crypto.randomUUID(),
    employeeId,
    month,
    year,
    daysPresent: 0,
    workingDays: 0,
    allowance: 0,
    incentive: 0,
    attendanceBonus: 0,
    performanceBonus: 0,
    overtimePay: 0,
    otherEarning: 0,
    arrears: 0,
    remarks: '',
  }
}

function blankFinanceRecord(): FinanceRecord {
  return {
    id: crypto.randomUUID(),
    date: new Date().toISOString().slice(0, 10),
    type: 'Expense',
    category: '',
    party: '',
    referenceNo: '',
    description: '',
    amount: 0,
    vatAmount: 0,
    paymentMode: '',
    status: 'Pending',
    attachmentName: '',
  }
}

function blankAssetRecord(): AssetRecord {
  return {
    id: crypto.randomUUID(),
    assetCode: '',
    assetName: '',
    category: '',
    assignedToEmployeeId: '',
    purchaseDate: '',
    purchaseCost: 0,
    serialNumber: '',
    location: '',
    condition: 'Good',
    warrantyExpiry: '',
    status: 'Active',
    notes: '',
  }
}

function blankDocumentRecord(): DocumentRecord {
  return {
    id: crypto.randomUUID(),
    documentDate: new Date().toISOString().slice(0, 10),
    documentType: 'PDF',
    linkedModule: '',
    linkedReference: '',
    title: '',
    fileName: '',
    fileSizeKb: 0,
    notes: '',
  }
}

function normalizeStoredPayrollInput(input: Partial<PayrollInput>): PayrollInput {
  return {
    ...blankPayrollInput(text(input.month) || currentMonth, text(input.year) || currentYear),
    ...input,
    id: input.id || crypto.randomUUID(),
    employeeId: text(input.employeeId),
    month: normalizeMonth(input.month),
    year: normalizeYear(input.year),
    daysPresent: numberValue(input.daysPresent),
    workingDays: numberValue(input.workingDays),
    allowance: numberValue(input.allowance),
    incentive: numberValue(input.incentive),
    attendanceBonus: numberValue(input.attendanceBonus),
    performanceBonus: numberValue(input.performanceBonus),
    overtimePay: numberValue(input.overtimePay),
    otherEarning: numberValue(input.otherEarning),
    arrears: numberValue(input.arrears),
    remarks: text(input.remarks),
  }
}

function normalizeStoredFinanceRecord(record: Partial<FinanceRecord>): FinanceRecord {
  return {
    ...blankFinanceRecord(),
    ...record,
    id: record.id || crypto.randomUUID(),
    date: normalizeDate(record.date),
    amount: numberValue(record.amount),
    vatAmount: numberValue(record.vatAmount),
  }
}

function normalizeStoredAssetRecord(record: Partial<AssetRecord>): AssetRecord {
  return {
    ...blankAssetRecord(),
    ...record,
    id: record.id || crypto.randomUUID(),
    purchaseDate: normalizeDate(record.purchaseDate),
    purchaseCost: numberValue(record.purchaseCost),
    warrantyExpiry: normalizeDate(record.warrantyExpiry),
  }
}

function normalizeStoredDocumentRecord(record: Partial<DocumentRecord>): DocumentRecord {
  return {
    ...blankDocumentRecord(),
    ...record,
    id: record.id || crypto.randomUUID(),
    documentDate: normalizeDate(record.documentDate),
    fileSizeKb: numberValue(record.fileSizeKb),
  }
}

function getInitialState(): StoredState {
  const fallback: StoredState = {
    employees: [],
    financeRecords: [],
    assetRecords: [],
    documentRecords: [],
    payrollInputs: [],
    deductions: [],
    settings: defaultSettings,
  }

  try {
    const raw = localStorage.getItem(storageKey)
    if (!raw) return fallback
    const parsed = JSON.parse(raw) as Partial<StoredState>

    return {
      employees: Array.isArray(parsed.employees)
        ? parsed.employees.map((employee) => normalizeStoredEmployee(employee))
        : [],
      financeRecords: Array.isArray(parsed.financeRecords)
        ? parsed.financeRecords.map((record) => normalizeStoredFinanceRecord(record))
        : [],
      assetRecords: Array.isArray(parsed.assetRecords)
        ? parsed.assetRecords.map((record) => normalizeStoredAssetRecord(record))
        : [],
      documentRecords: Array.isArray(parsed.documentRecords)
        ? parsed.documentRecords.map((record) => normalizeStoredDocumentRecord(record))
        : [],
      payrollInputs: Array.isArray(parsed.payrollInputs)
        ? parsed.payrollInputs.map((input) => normalizeStoredPayrollInput(input))
        : [],
      deductions: Array.isArray(parsed.deductions) ? parsed.deductions : [],
      settings: { ...defaultSettings, ...(parsed.settings ?? {}) },
    }
  } catch {
    return fallback
  }
}

function saveState(state: StoredState) {
  localStorage.setItem(storageKey, JSON.stringify(state))
}

function text(value: unknown) {
  return String(value ?? '').trim()
}

function normalizeHeader(value: string) {
  return value.trim().toLowerCase().replace(/\s+/g, ' ')
}

function rowValue(row: Record<string, unknown>, ...candidates: string[]) {
  const keys = Object.keys(row)
  for (const candidate of candidates) {
    const match = keys.find((key) => normalizeHeader(key) === normalizeHeader(candidate))
    if (match) return row[match]
  }
  return ''
}

function numberValue(value: unknown) {
  const parsed = Number(String(value ?? '').replace(/,/g, '').trim())
  return Number.isFinite(parsed) ? parsed : 0
}

function normalizeMonth(value: unknown) {
  const raw = text(value)
  const isoDateMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (isoDateMatch) return isoDateMatch[2]
  const monthByName = monthOptions.find(
    (option) => option.label.toLowerCase() === raw.toLowerCase(),
  )
  if (monthByName) return monthByName.value

  const parsed = Number(raw)
  if (Number.isFinite(parsed) && parsed >= 1 && parsed <= 12) {
    return String(parsed).padStart(2, '0')
  }

  return currentMonth
}

function normalizeYear(value: unknown) {
  const raw = text(value)
  const isoDateMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/)
  if (isoDateMatch) return isoDateMatch[1]
  if (/^\d{4}$/.test(raw)) return raw
  return currentYear
}

function excelSerialDateToIso(value: number) {
  const excelEpoch = Date.UTC(1899, 11, 30)
  const date = new Date(excelEpoch + value * 24 * 60 * 60 * 1000)
  return formatSifDate(date)
}

function normalizeDate(value: unknown) {
  if (value instanceof Date) return value.toISOString().slice(0, 10)

  if (typeof value === 'number') {
    return excelSerialDateToIso(value)
  }

  return text(value)
}

function digitsOnly(value: string) {
  return value.replace(/\D/g, '')
}

function formatEmiratesId(value: string) {
  const digits = digitsOnly(value).slice(0, 15)
  const segments = [3, 4, 7, 1]
  let cursor = 0
  const parts: string[] = []

  segments.forEach((length) => {
    const part = digits.slice(cursor, cursor + length)
    if (part) parts.push(part)
    cursor += length
  })

  return parts.join('-')
}

function normalizeRoleDesignation(roleDesignation: string) {
  return roleDesignation.trim().toLowerCase()
}

function roleCategory(roleDesignation: string) {
  const normalized = normalizeRoleDesignation(roleDesignation)
  if (normalized.includes('driver')) return 'Driver'
  if (normalized.includes('rider')) return 'Rider'
  if (normalized.includes('helper')) return 'Helper'
  if (normalized.includes('admin')) return 'Admin'
  if (normalized.includes('account')) return 'Accountant'
  if (normalized.includes('supervisor')) return 'Supervisor'
  if (normalized.includes('sale')) return 'Sales'
  return 'General'
}

function monthName(month: string) {
  return monthOptions.find((option) => option.value === month)?.label ?? month
}

function formatMoney(value: number, currency: string) {
  return new Intl.NumberFormat('en-AE', {
    style: 'currency',
    currency: currency || 'AED',
    maximumFractionDigits: 2,
  }).format(value || 0)
}

function roundMoney(value: number) {
  return Math.round((value + Number.EPSILON) * 100) / 100
}

function formatSifAmount(value: number) {
  return roundMoney(Math.max(0, value || 0)).toFixed(2)
}

function safeSifCell(value: unknown, maxLength?: number) {
  const cleaned = text(value).replace(/[\r\n,]/g, ' ').replace(/\s+/g, ' ').trim()
  return typeof maxLength === 'number' ? cleaned.slice(0, maxLength) : cleaned
}

function normalizeWpsId(value: string, length: number) {
  const cleaned = text(value).replace(/[^a-zA-Z0-9]/g, '').toUpperCase()
  return cleaned ? cleaned.padStart(length, '0').slice(-length) : ''
}

function normalizeWpsRoutingCode(value: string) {
  return digitsOnly(value).slice(0, 9)
}

function normalizeWpsAccount(value: string) {
  return text(value).replace(/[^a-zA-Z0-9]/g, '').toUpperCase().slice(0, 23)
}

function payPeriodStart(month: string, year: string) {
  return new Date(Number(year), Number(month) - 1, 1)
}

function payPeriodEnd(month: string, year: string) {
  return new Date(Number(year), Number(month), 0)
}

function formatSifDate(date: Date) {
  return [
    date.getFullYear(),
    String(date.getMonth() + 1).padStart(2, '0'),
    String(date.getDate()).padStart(2, '0'),
  ].join('-')
}

function formatSifSalaryMonth(month: string, year: string) {
  return `${String(month).padStart(2, '0')}${year}`
}

function formatSifCreationTime(date: Date) {
  return `${String(date.getHours()).padStart(2, '0')}${String(date.getMinutes()).padStart(2, '0')}`
}

function formatSifTimestamp(date: Date) {
  return [
    String(date.getFullYear()).slice(-2),
    String(date.getMonth() + 1).padStart(2, '0'),
    String(date.getDate()).padStart(2, '0'),
    String(date.getHours()).padStart(2, '0'),
    String(date.getMinutes()).padStart(2, '0'),
    String(date.getSeconds()).padStart(2, '0'),
  ].join('')
}

function calculateWpsPayment(row: PayrollRow) {
  const fixedBase =
    row.basicSalary + row.accommodationAllowance + row.travelAllowance + row.allowance
  const fixedEarned =
    row.salaryType === 'Per Day' && row.workingDays > 0
      ? (fixedBase / row.workingDays) * Math.max(0, row.daysPresent)
      : fixedBase
  const totalPay = roundMoney(Math.max(0, row.netSalary))
  const fixedIncome = roundMoney(Math.min(totalPay, Math.max(0, fixedEarned)))
  const variableIncome = roundMoney(Math.max(0, totalPay - fixedIncome))
  const leaveDays =
    row.workingDays > 0 ? Math.max(0, Math.round(row.workingDays - row.daysPresent)) : 0

  return {
    fixedIncome,
    variableIncome,
    leaveDays,
    totalPay,
  }
}

function selectOptions(fieldKey: keyof Employee) {
  if (fieldKey === 'status') return ['Active', 'Inactive']
  if (fieldKey === 'salaryType') return ['Fixed Salary', 'Per Day']
  return []
}

function workbookColumnWidth(rows: Record<string, unknown>[], header: string) {
  const values = rows.map((row) => text(row[header]))
  const longestValue = values.reduce((max, value) => Math.max(max, value.length), 0)
  return Math.min(Math.max(header.length, longestValue, 12) + 2, 28)
}

function validateExcelFile(file: File) {
  const filename = file.name.toLowerCase()
  if (!filename.endsWith('.xlsx')) {
    return 'Only .xlsx files are supported.'
  }

  if (file.size === 0) {
    return 'The selected Excel file is empty.'
  }

  if (file.size > maxExcelFileSizeBytes) {
    return 'The selected Excel file is too large. Keep it under 10 MB.'
  }

  return null
}

function findMissingColumns(headers: string[], expectedColumns: string[]) {
  const normalizedHeaders = new Set(headers.map(normalizeHeader))
  return expectedColumns.filter(
    (column) => !normalizedHeaders.has(normalizeHeader(column)),
  )
}

async function readWorkbook(file: File) {
  const workbook = new ExcelJS.Workbook()
  await workbook.xlsx.load(await file.arrayBuffer())
  return workbook
}

function readFirstWorksheet(workbook: ExcelJS.Workbook) {
  const worksheet = workbook.worksheets[0]
  if (!worksheet) {
    throw new Error('The Excel workbook does not contain any sheets.')
  }

  return worksheet
}

function excelCellValue(value: ExcelJS.CellValue): unknown {
  if (value instanceof Date) return value
  if (!value || typeof value !== 'object') return value ?? ''
  if ('result' in value) return value.result ?? ''
  if ('text' in value) return value.text ?? ''
  if ('richText' in value) return value.richText.map((item) => item.text).join('')
  return String(value)
}

function worksheetToJsonRows(worksheet: ExcelJS.Worksheet) {
  const headers: string[] = []
  worksheet.getRow(1).eachCell({ includeEmpty: true }, (cell, columnNumber) => {
    headers[columnNumber - 1] = text(excelCellValue(cell.value))
  })

  const rows: Record<string, unknown>[] = []
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return
    const record: Record<string, unknown> = {}
    let hasValue = false

    headers.forEach((header, index) => {
      if (!header) return
      const value = excelCellValue(row.getCell(index + 1).value)
      record[header] = value
      if (text(value)) hasValue = true
    })

    if (hasValue) rows.push(record)
  })

  return rows
}

function resolveEmployeeForDeductionRow(
  row: Record<string, unknown>,
  employees: Employee[],
) {
  const employeeCode = text(rowValue(row, 'Employee Code'))
  const agreementNo = text(rowValue(row, 'Agreement No', 'Payroll Code', 'Project ID'))
  const driverName = text(rowValue(row, 'Driver', 'Employee Name', 'Name'))

  const byEmployeeCode = employeeCode
    ? employees.find(
        (item) => item.employeeCode.toLowerCase() === employeeCode.toLowerCase(),
      )
    : undefined
  if (byEmployeeCode) {
    return { employee: byEmployeeCode, matchedBy: 'Employee Code', lookupValue: employeeCode }
  }

  const byAgreement = agreementNo
    ? employees.find((item) => {
        const payrollKey = (item.payrollCode || item.employeeCode).toLowerCase()
        return payrollKey === agreementNo.toLowerCase()
      })
    : undefined
  if (byAgreement) {
    return { employee: byAgreement, matchedBy: 'Agreement No', lookupValue: agreementNo }
  }

  const byDriverName = driverName
    ? employees.find(
        (item) => item.employeeName.toLowerCase() === driverName.toLowerCase(),
      )
    : undefined
  if (byDriverName) {
    return { employee: byDriverName, matchedBy: 'Driver', lookupValue: driverName }
  }

  return {
    employee: undefined,
    matchedBy: employeeCode ? 'Employee Code' : agreementNo ? 'Agreement No' : 'Driver',
    lookupValue: employeeCode || agreementNo || driverName,
  }
}

function resolveEmployeeByCodeOrName(value: unknown, employees: Employee[]) {
  const lookup = text(value).toLowerCase()
  if (!lookup) return undefined

  return employees.find(
    (employee) =>
      employee.employeeCode.toLowerCase() === lookup ||
      employee.payrollCode.toLowerCase() === lookup ||
      employee.employeeName.toLowerCase() === lookup,
  )
}

async function writeWorkbook(
  filename: string,
  sheetName: string,
  rows: Record<string, unknown>[],
) {
  const workbook = new ExcelJS.Workbook()
  const worksheet = workbook.addWorksheet(sheetName)
  const headers = rows[0] ? Object.keys(rows[0]) : []

  if (headers.length) {
    worksheet.addRow(headers)
    rows.forEach((row) => {
      worksheet.addRow(headers.map((header) => row[header] ?? ''))
    })
    worksheet.columns = headers.map((header) => ({
      width: workbookColumnWidth(rows, header),
    }))
    worksheet.autoFilter = {
      from: { row: 1, column: 1 },
      to: { row: Math.max(rows.length + 1, 1), column: headers.length },
    }
    worksheet.getRow(1).font = { bold: true }
  }

  const buffer = await workbook.xlsx.writeBuffer()
  downloadText(
    filename,
    buffer as unknown as BlobPart,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  )
}

function writeTemplate(filename: string, sheetName: string, columns: string[]) {
  const exampleRow = Object.fromEntries(
    columns.map((column) => {
      if (column === 'Month') return [column, 'June']
      if (column === 'Year') return [column, currentYear]
      if (column.includes('Date') || column.includes('Expiry')) {
        return [column, '2026-06-30']
      }
      if (
        column.includes('Salary') ||
        column.includes('Allowance') ||
        column.includes('Bonus') ||
        column.includes('Incentive') ||
        column.includes('Overtime') ||
        column.includes('Earnings') ||
        column.includes('Arrears') ||
        column.includes('Deduction') ||
        column.includes('Days') ||
        column.includes('Due')
      ) {
        return [column, 0]
      }
      if (column === 'Status') return [column, 'Active']
      if (column === 'Salary Type') return [column, 'Fixed Salary']
      return [column, '']
    }),
  )
  writeWorkbook(filename, sheetName, [exampleRow])
}

function downloadJson(filename: string, payload: unknown) {
  const blob = new Blob([JSON.stringify(payload, null, 2)], {
    type: 'application/json;charset=utf-8',
  })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

function downloadText(
  filename: string,
  payload: BlobPart,
  type = 'text/plain;charset=utf-8',
) {
  const blob = new Blob([payload], { type })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

function employeeToImportRow(employee: Employee) {
  return {
    'Employee Code': employee.employeeCode,
    'Payroll Code': employee.payrollCode,
    'WPS Person ID': employee.wpsPersonId,
    'WPS Agent Routing Code': employee.wpsRoutingCode,
    'WPS IBAN / Account': employee.wpsAccountNumber,
    'Employee Name': employee.employeeName,
    'Role / Designation': employee.roleDesignation,
    'Date of Joining': employee.dateOfJoining,
    'Mobile Number': employee.mobileNumber,
    'Basic Salary': employee.basicSalary,
    Status: employee.status,
    Email: employee.email,
    'Emirates ID': employee.emiratesId,
    'Emirates ID Expiry': employee.emiratesIdExpiry,
    'Visa Expiry Date': employee.visaExpiryDate,
    'Passport Number': employee.passportNumber,
    'Passport Expiry': employee.passportExpiry,
    'Emergency Contact Name': employee.emergencyContactName,
    'Emergency Contact Mobile': employee.emergencyContactMobile,
    'Emergency Contact Name 2': employee.emergencyContactName2,
    'Emergency Contact Mobile 2': employee.emergencyContactMobile2,
    'License Number': employee.licenseNumber,
    'License Expiry': employee.licenseExpiry,
    'Permit Number': employee.permitNumber,
    'Permit Expiry': employee.permitExpiry,
    'Work Hours': employee.workHours,
    'Salary Type': employee.salaryType,
    'Days Present': employee.daysPresent,
    'Working Days': employee.workingDays,
    'Accommodation Allowance': employee.accommodationAllowance,
    'Travel Allowance': employee.travelAllowance,
    Incentive: employee.incentive,
    'Due From Last Month': employee.dueFromLastMonth,
    Allowance: employee.allowance,
    Remarks: employee.remarks,
  }
}

function deductionToExportRow(deduction: Deduction, employee?: Employee) {
  const amount =
    deduction.advanceDeduction +
    deduction.fineDeduction +
    deduction.salikDeduction +
    deduction.otherDeduction
  const invoiceType =
    deduction.fineDeduction > 0
      ? 'Miscellaneous Invoice'
      : deduction.salikDeduction > 0
        ? 'Tolls Invoice'
        : 'Deduction Invoice'
  const invoiceDate = deduction.fineDate || deduction.salikDate
  return {
    'Invoice No': deduction.referenceNo || deduction.transactionId,
    'Invoice Type': invoiceType,
    Date: invoiceDate,
    Customer: '',
    Driver: employee?.employeeName ?? '',
    'Agreement No': employee?.payrollCode || employee?.employeeCode || '',
    Vehicle: employee?.permitNumber || employee?.licenseNumber || '',
    Amount: amount,
    'Employee Code': employee?.employeeCode ?? '',
    'Employee Name': employee?.employeeName ?? '',
    Month: monthName(deduction.month),
    Year: deduction.year,
    'Transaction ID': deduction.transactionId,
    'Fine Receipt No': deduction.fineReceiptNo,
    'Fine Transaction No': deduction.fineTransactionNo,
    'Salik Receipt No': deduction.salikReceiptNo,
    'Salik Transaction No': deduction.salikTransactionNo,
    'Reason / Remarks': deduction.reasonRemarks,
    'Reference No': deduction.referenceNo,
    'Attachment Name': deduction.attachmentName,
  }
}

function payrollInputToExportRow(input: PayrollInput, employee?: Employee) {
  return {
    'Employee Code': employee?.employeeCode ?? '',
    'Payroll Code': employee?.payrollCode || employee?.employeeCode || '',
    'Employee Name': employee?.employeeName ?? '',
    Month: monthName(input.month),
    Year: input.year,
    'Days Present': input.daysPresent,
    'Working Days': input.workingDays,
    Allowance: input.allowance,
    Incentive: input.incentive,
    'Attendance Bonus': input.attendanceBonus,
    'Performance Bonus / Tip': input.performanceBonus,
    'Overtime Pay': input.overtimePay,
    'Other Earnings': input.otherEarning,
    'Arrears / Due From Last Month': input.arrears,
    Remarks: input.remarks,
  }
}

function financeRecordToExportRow(record: FinanceRecord) {
  return {
    Date: record.date,
    Type: record.type,
    Category: record.category,
    Party: record.party,
    'Reference No': record.referenceNo,
    Description: record.description,
    Amount: record.amount,
    'VAT Amount': record.vatAmount,
    'Payment Mode': record.paymentMode,
    Status: record.status,
    'Attachment Name': record.attachmentName,
  }
}

function assetRecordToExportRow(record: AssetRecord, employee?: Employee) {
  return {
    'Asset Code': record.assetCode,
    'Asset Name': record.assetName,
    Category: record.category,
    'Assigned Employee Code': employee?.employeeCode ?? '',
    'Assigned Employee Name': employee?.employeeName ?? '',
    'Purchase Date': record.purchaseDate,
    'Purchase Cost': record.purchaseCost,
    'Serial Number': record.serialNumber,
    Location: record.location,
    Condition: record.condition,
    Status: record.status,
    'Warranty Expiry': record.warrantyExpiry,
    Notes: record.notes,
  }
}

function documentRecordToExportRow(record: DocumentRecord) {
  return {
    'Document Date': record.documentDate,
    'Document Type': record.documentType,
    'Linked Module': record.linkedModule,
    'Linked Reference': record.linkedReference,
    Title: record.title,
    'File Name': record.fileName,
    'File Size KB': record.fileSizeKb,
    Notes: record.notes,
  }
}

function payrollToExportRow(row: PayrollRow, index: number) {
  const lossOfPay =
    row.salaryType === 'Per Day' && row.workingDays > 0
      ? Math.max(
          0,
          row.basicSalary +
            row.accommodationAllowance +
            row.travelAllowance +
            row.allowance +
            row.incentive -
            row.salaryEarned,
        )
      : 0
  return {
    Sl: index + 1,
    Visa: row.emiratesId,
    Name: row.employeeName,
    Project: row.roleDesignation,
    'Project ID': row.payrollCode || row.employeeCode,
    'Monthly Rental DDN or Salary': row.basicSalary,
    'Total No. Days': row.workingDays,
    'Salary Type': row.salaryType,
    Earnings: row.salaryEarned,
    Allowance: row.allowance,
    'Performance Bonus / Tip': row.performanceBonus,
    'Attendance Bonus': row.attendanceBonus,
    Incentive: row.incentive,
    'O/T': row.overtimePay,
    'Other Earnings': row.otherEarning,
    Arrears: row.dueFromLastMonth,
    LOP: lossOfPay,
    'Eligible Days': row.daysPresent || row.workingDays,
    'Gross Earning': row.grossSalary,
    COD: 0,
    'Client Penalty': row.fineDeduction,
    'Vehicle Rent': 0,
    'Advance Deduction': row.advanceDeduction,
    'Salik Deduction': row.salikDeduction,
    'Other Deduction': row.otherDeduction,
    'Net Salary': row.netSalary,
    'Employee Code': row.employeeCode,
    'Payroll Code': row.payrollCode || row.employeeCode,
    Remarks: row.remarks,
  }
}

function buildPayrollRows(
  employees: Employee[],
  payrollInputs: PayrollInput[],
  deductions: Deduction[],
  month: string,
  year: string,
): PayrollRow[] {
  return employees.map((employee) => {
    const periodInput = payrollInputs.find(
      (input) =>
        input.employeeId === employee.id &&
        input.month === month &&
        input.year === year,
    )
    const matchingDeductions = deductions.filter(
      (deduction) =>
        deduction.employeeId === employee.id &&
        deduction.month === month &&
        deduction.year === year,
    )
    const totals = matchingDeductions.reduce(
      (sum, deduction) => ({
        advanceDeduction: sum.advanceDeduction + deduction.advanceDeduction,
        fineDeduction: sum.fineDeduction + deduction.fineDeduction,
        salikDeduction: sum.salikDeduction + deduction.salikDeduction,
        otherDeduction: sum.otherDeduction + deduction.otherDeduction,
      }),
      {
        advanceDeduction: 0,
        fineDeduction: 0,
        salikDeduction: 0,
        otherDeduction: 0,
      },
    )
    const daysPresent = periodInput?.daysPresent || employee.daysPresent
    const workingDays = periodInput?.workingDays || employee.workingDays
    const allowance = periodInput?.allowance ?? employee.allowance
    const incentive = periodInput?.incentive ?? employee.incentive
    const attendanceBonus = periodInput?.attendanceBonus ?? 0
    const performanceBonus = periodInput?.performanceBonus ?? 0
    const overtimePay = periodInput?.overtimePay ?? 0
    const otherEarning = periodInput?.otherEarning ?? 0
    const arrears = periodInput?.arrears ?? employee.dueFromLastMonth
    const remarks = [employee.remarks, periodInput?.remarks].filter(Boolean).join(' | ')
    const baseSalary =
      employee.basicSalary +
      employee.accommodationAllowance +
      employee.travelAllowance +
      allowance
    const salaryEarned =
      employee.salaryType === 'Per Day' &&
      workingDays > 0 &&
      daysPresent >= 0
        ? (baseSalary / workingDays) * daysPresent
        : baseSalary
    const grossSalary =
      salaryEarned +
      incentive +
      attendanceBonus +
      performanceBonus +
      overtimePay +
      otherEarning
    const payableBeforeDeductions = grossSalary + arrears
    const totalDeduction =
      totals.advanceDeduction +
      totals.fineDeduction +
      totals.salikDeduction +
      totals.otherDeduction

    return {
        ...employee,
        daysPresent,
        workingDays,
        allowance,
        incentive,
        dueFromLastMonth: arrears,
        remarks,
        ...totals,
        attendanceBonus,
        performanceBonus,
        overtimePay,
        otherEarning,
        arrears,
        salaryEarned,
        grossSalary,
        payableBeforeDeductions,
        totalDeduction,
        netSalary: payableBeforeDeductions - totalDeduction,
        month,
        year,
      }
  })
}

function validateEmployee(
  employee: Employee,
  employees: Employee[],
  currentEmployeeId?: string,
) {
  const errors: string[] = []
  const warnings: string[] = []

  employeeFields
    .filter((field) => field.required)
    .forEach((field) => {
      const value = employee[field.key]
      if (field.key === 'basicSalary') {
        if (!Number.isFinite(employee.basicSalary) || employee.basicSalary <= 0) {
          errors.push('Basic Salary is required and must be numeric.')
        }
        return
      }
      if (!text(value)) errors.push(`${field.label} is required.`)
    })

  const otherEmployees = employees.filter((item) => item.id !== currentEmployeeId)
  const code = employee.employeeCode.toLowerCase()
  if (code && otherEmployees.some((item) => item.employeeCode.toLowerCase() === code)) {
    errors.push('Employee Code must be unique.')
  }

  const emiratesId = employee.emiratesId.toLowerCase()
  if (
    emiratesId &&
    otherEmployees.some((item) => item.emiratesId.toLowerCase() === emiratesId)
  ) {
    errors.push('Emirates ID must be unique when provided.')
  }

  const passportNumber = employee.passportNumber.toLowerCase()
  if (
    passportNumber &&
    otherEmployees.some((item) => item.passportNumber.toLowerCase() === passportNumber)
  ) {
    errors.push('Passport Number must be unique when provided.')
  }

  const mobileNumber = employee.mobileNumber.toLowerCase()
  if (
    mobileNumber &&
    otherEmployees.some((item) => item.mobileNumber.toLowerCase() === mobileNumber)
  ) {
    warnings.push('Mobile number is duplicated. This is a warning only.')
  }

  if (employee.emiratesId) {
    const formatted = formatEmiratesId(employee.emiratesId)
    if (formatted.length !== 18) {
      warnings.push('Emirates ID should be in the 000-0000-0000000-0 format.')
    }
  } else {
    warnings.push('Missing Emirates ID for payroll/data-keeping.')
  }

  if (!employee.emiratesIdExpiry) {
    warnings.push('Missing Emirates ID expiry date.')
  }

  if (!employee.visaExpiryDate) {
    warnings.push('Missing visa expiry date.')
  }

  if (!employee.passportNumber) {
    warnings.push('Missing passport number for payroll/data-keeping.')
  }

  if (!employee.passportExpiry) {
    warnings.push('Missing passport expiry date.')
  }

  if (!employee.emergencyContactName || !employee.emergencyContactMobile) {
    warnings.push('At least one emergency contact name and mobile should be captured.')
  }

  if (!employee.workHours) {
    warnings.push('Missing work hours.')
  }

  if (!employee.wpsPersonId) {
    warnings.push('Missing WPS Person ID. This is required before SIF export.')
  } else if (normalizeWpsId(employee.wpsPersonId, 14).length !== 14) {
    warnings.push('WPS Person ID should be 14 alphanumeric characters after padding.')
  }

  if (!employee.wpsRoutingCode) {
    warnings.push('Missing employee WPS agent routing code. The employer routing code can be used as fallback.')
  } else if (normalizeWpsRoutingCode(employee.wpsRoutingCode).length !== 9) {
    warnings.push('Employee WPS agent routing code should be exactly 9 digits.')
  }

  if (!employee.wpsAccountNumber) {
    warnings.push('Missing WPS IBAN / account number. This is required before SIF export.')
  }

  const category = roleCategory(employee.roleDesignation)
  if (category === 'Driver' || category === 'Rider') {
    if (!employee.licenseNumber) {
      warnings.push('License number is expected for driver/rider workflows.')
    }
    if (!employee.licenseExpiry) {
      warnings.push('License expiry is expected for driver/rider workflows.')
    }
  }

  if (
    employee.salaryType === 'Per Day' &&
    (!Number.isFinite(employee.workingDays) || employee.workingDays <= 0)
  ) {
    errors.push('Working Days must be greater than zero for Per Day salary type.')
  }

  if (employee.daysPresent < 0 || employee.workingDays < 0) {
    errors.push('Days Present and Working Days cannot be negative.')
  }

  if (
    employee.salaryType === 'Fixed Salary' &&
    employee.workingDays > 0 &&
    employee.daysPresent > 0 &&
    employee.daysPresent < employee.workingDays
  ) {
    warnings.push(
      'Fixed Salary employee has fewer days present than working days. Payroll keeps the fixed salary unless you change the salary type.',
    )
  }

  return { errors, warnings }
}

function validateDeduction(
  deduction: Deduction,
  employees: Employee[],
  existingDeductions: Deduction[],
  currentEmployeeId?: string,
) {
  const errors: string[] = []
  const warnings: string[] = []

  if (!deduction.employeeId) {
    errors.push('Employee is required.')
  } else if (!employees.some((employee) => employee.id === deduction.employeeId)) {
    errors.push('Selected employee was not found.')
  }

  if (!monthOptions.some((option) => option.value === deduction.month)) {
    errors.push('Month is required and must be valid.')
  }

  if (!/^\d{4}$/.test(deduction.year)) {
    errors.push('Year is required and must be a 4-digit year.')
  }

  if (
    deduction.transactionId &&
    existingDeductions.some(
      (item) =>
        item.id !== deduction.id &&
        item.transactionId &&
        item.transactionId.toLowerCase() === deduction.transactionId.toLowerCase(),
    )
  ) {
    errors.push('Transaction ID must be unique when provided.')
  }

  ;([
    ['Advance Deduction', deduction.advanceDeduction],
    ['Fine Deduction', deduction.fineDeduction],
    ['Salik Deduction', deduction.salikDeduction],
    ['Other Deduction', deduction.otherDeduction],
  ] as const).forEach(([label, value]) => {
    if (!Number.isFinite(value)) {
      errors.push(`${label} must be numeric.`)
      return
    }
    if (value < 0) {
      errors.push(`${label} cannot be negative.`)
    }
  })

  if (
    deduction.advanceDeduction === 0 &&
    deduction.fineDeduction === 0 &&
    deduction.salikDeduction === 0 &&
    deduction.otherDeduction === 0
  ) {
    warnings.push('This deduction row totals zero.')
  }

  if (currentEmployeeId && deduction.employeeId !== currentEmployeeId) {
    warnings.push('Deduction employee reference changed during validation.')
  }

  if (deduction.fineDeduction > 0 && !deduction.fineReceiptNo && !deduction.fineTransactionNo) {
    warnings.push('Fine deduction should include a receipt number or transaction number.')
  }

  if (deduction.salikDeduction > 0 && !deduction.salikReceiptNo && !deduction.salikTransactionNo) {
    warnings.push('Salik deduction should include a receipt number or transaction number.')
  }

  return { errors, warnings }
}

function validatePayrollInput(
  input: PayrollInput,
  employees: Employee[],
  existingInputs: PayrollInput[],
) {
  const errors: string[] = []
  const warnings: string[] = []

  if (!input.employeeId) {
    errors.push('Employee is required.')
  } else if (!employees.some((employee) => employee.id === input.employeeId)) {
    errors.push('Selected employee was not found.')
  }

  if (!monthOptions.some((option) => option.value === input.month)) {
    errors.push('Month is required and must be valid.')
  }

  if (!/^\d{4}$/.test(input.year)) {
    errors.push('Year is required and must be a 4-digit year.')
  }

  if (
    input.employeeId &&
    existingInputs.some(
      (item) =>
        item.id !== input.id &&
        item.employeeId === input.employeeId &&
        item.month === input.month &&
        item.year === input.year,
    )
  ) {
    errors.push('Only one payroll input row is allowed per employee/month/year.')
  }

  ;([
    ['Days Present', input.daysPresent],
    ['Working Days', input.workingDays],
    ['Allowance', input.allowance],
    ['Incentive', input.incentive],
    ['Attendance Bonus', input.attendanceBonus],
    ['Performance Bonus / Tip', input.performanceBonus],
    ['Overtime Pay', input.overtimePay],
    ['Other Earnings', input.otherEarning],
    ['Arrears / Due From Last Month', input.arrears],
  ] as const).forEach(([label, value]) => {
    if (!Number.isFinite(value)) {
      errors.push(`${label} must be numeric.`)
      return
    }
    if (value < 0) {
      errors.push(`${label} cannot be negative.`)
    }
  })

  if (input.workingDays > 0 && input.daysPresent > input.workingDays) {
    warnings.push('Days Present is greater than Working Days.')
  }

  if (
    input.daysPresent === 0 &&
    input.workingDays === 0 &&
    input.allowance === 0 &&
    input.incentive === 0 &&
    input.attendanceBonus === 0 &&
    input.performanceBonus === 0 &&
    input.overtimePay === 0 &&
    input.otherEarning === 0 &&
    input.arrears === 0
  ) {
    warnings.push('This payroll input row has no monthly values.')
  }

  return { errors, warnings }
}

function validateWpsPayroll(
  rows: PayrollRow[],
  settings: Settings,
): WpsValidation {
  const errors: string[] = []
  const warnings: string[] = []
  const employerId = normalizeWpsId(settings.wpsEmployerId, 13)
  const employerRoutingCode = normalizeWpsRoutingCode(settings.wpsRoutingCode)

  if (!rows.length) {
    errors.push('No payroll rows are available for this period.')
  }

  if (!employerId) {
    errors.push('WPS Employer ID is required in Settings.')
  } else if (employerId.length !== 13) {
    errors.push('WPS Employer ID must be 13 alphanumeric characters or fewer.')
  }

  if (!employerRoutingCode) {
    errors.push('WPS employer agent routing code is required in Settings.')
  } else if (employerRoutingCode.length !== 9) {
    errors.push('WPS employer agent routing code must be exactly 9 digits.')
  }

  const personIds = new Set<string>()
  rows.forEach((row) => {
    const rowLabel = `${row.employeeCode || 'Employee'} - ${row.employeeName || 'Unnamed'}`
    const personId = normalizeWpsId(row.wpsPersonId, 14)
    const routingCode = normalizeWpsRoutingCode(row.wpsRoutingCode || settings.wpsRoutingCode)
    const accountNumber = normalizeWpsAccount(row.wpsAccountNumber)
    const payment = calculateWpsPayment(row)

    if (!personId) {
      errors.push(`${rowLabel}: WPS Person ID is required.`)
    } else if (personId.length !== 14) {
      errors.push(`${rowLabel}: WPS Person ID must be 14 alphanumeric characters or fewer.`)
    } else if (personIds.has(personId)) {
      errors.push(`${rowLabel}: WPS Person ID must be unique in the SIF file.`)
    } else {
      personIds.add(personId)
    }

    if (!routingCode) {
      errors.push(`${rowLabel}: WPS agent routing code is required.`)
    } else if (routingCode.length !== 9) {
      errors.push(`${rowLabel}: WPS agent routing code must be exactly 9 digits.`)
    }

    if (!accountNumber) {
      errors.push(`${rowLabel}: WPS IBAN / account number is required.`)
    } else if (accountNumber.startsWith('AE') && accountNumber.length !== 23) {
      warnings.push(`${rowLabel}: UAE IBAN values normally contain 23 characters.`)
    } else if (accountNumber.length > 23) {
      errors.push(`${rowLabel}: WPS account number must be 23 characters or fewer.`)
    }

    if (payment.totalPay <= 0) {
      errors.push(`${rowLabel}: net salary must be greater than zero for WPS export.`)
    }

    if (row.netSalary < 0) {
      errors.push(`${rowLabel}: net salary cannot be negative.`)
    }
  })

  return { errors, warnings }
}

function buildSifRecords(
  rows: PayrollRow[],
  settings: Settings,
  month: string,
  year: string,
): SifRecordSet {
  const validation = validateWpsPayroll(rows, settings)
  const createdAt = new Date()
  const employerId = normalizeWpsId(settings.wpsEmployerId, 13)
  const employerRoutingCode = normalizeWpsRoutingCode(settings.wpsRoutingCode)
  const periodStart = payPeriodStart(month, year)
  const periodEnd = payPeriodEnd(month, year)
  const daysInPeriod = periodEnd.getDate()
  const salaryMonth = formatSifSalaryMonth(month, year)
  const creationDate = formatSifDate(createdAt)
  const creationTime = formatSifCreationTime(createdAt)
  const details = rows.map((row) => {
    const routingCode = normalizeWpsRoutingCode(row.wpsRoutingCode || settings.wpsRoutingCode)
    const personId = normalizeWpsId(row.wpsPersonId, 14)
    const accountNumber = normalizeWpsAccount(row.wpsAccountNumber)
    const payment = calculateWpsPayment(row)

    return {
      row,
      routingCode,
      personId,
      accountNumber,
      ...payment,
    }
  })
  const lines = details.map((item) =>
    [
      'EDR',
      item.personId,
      item.routingCode,
      item.accountNumber,
      formatSifDate(periodStart),
      formatSifDate(periodEnd),
      String(daysInPeriod),
      formatSifAmount(item.fixedIncome),
      formatSifAmount(item.variableIncome),
      String(item.leaveDays),
    ].join(','),
  )
  const totalSalary = details.reduce((sum, item) => sum + item.totalPay, 0)
  lines.push(
    [
      'SCR',
      employerId,
      employerRoutingCode,
      creationDate,
      creationTime,
      salaryMonth,
      String(details.length).padStart(2, '0'),
      formatSifAmount(totalSalary),
      'AED',
      safeSifCell(settings.wpsEmployerReference || settings.companyName, 35),
    ].join(','),
  )

  return {
    filename: `wps-${employerId || 'employer'}-${formatSifTimestamp(createdAt)}.SIF`,
    content: `${lines.join('\r\n')}\r\n`,
    rows: details.map((item) => ({
      employeeCode: item.row.employeeCode,
      employeeName: item.row.employeeName,
      fixedIncome: item.fixedIncome,
      variableIncome: item.variableIncome,
      leaveDays: item.leaveDays,
      totalPay: item.totalPay,
    })),
    validation,
  }
}


export type {
  PageKey,
  EmployeeStatus,
  Employee,
  Deduction,
  PayrollInput,
  FinanceRecord,
  AssetRecord,
  DocumentRecord,
  Settings,
  StoredState,
  PayrollRow,
  WpsValidation,
  SifRecordSet,
  Notice,
}

export {
  storageKey,
  now,
  currentMonth,
  currentYear,
  defaultSettings,
  pages,
  employeeFields,
  employeeTemplateColumns,
  deductionsTemplateColumns,
  payrollInputTemplateColumns,
  financeTemplateColumns,
  assetTemplateColumns,
  documentTemplateColumns,
  payrollExportColumns,
  maxExcelFileSizeBytes,
  monthOptions,
  blankEmployee,
  normalizeStoredEmployee,
  blankDeduction,
  blankPayrollInput,
  blankFinanceRecord,
  blankAssetRecord,
  blankDocumentRecord,
  normalizeStoredPayrollInput,
  normalizeStoredFinanceRecord,
  normalizeStoredAssetRecord,
  normalizeStoredDocumentRecord,
  getInitialState,
  saveState,
  text,
  normalizeHeader,
  rowValue,
  numberValue,
  normalizeMonth,
  normalizeYear,
  excelSerialDateToIso,
  normalizeDate,
  digitsOnly,
  formatEmiratesId,
  normalizeRoleDesignation,
  roleCategory,
  monthName,
  formatMoney,
  roundMoney,
  formatSifAmount,
  safeSifCell,
  normalizeWpsId,
  normalizeWpsRoutingCode,
  normalizeWpsAccount,
  payPeriodStart,
  payPeriodEnd,
  formatSifDate,
  formatSifSalaryMonth,
  formatSifCreationTime,
  formatSifTimestamp,
  calculateWpsPayment,
  selectOptions,
  workbookColumnWidth,
  validateExcelFile,
  findMissingColumns,
  readWorkbook,
  readFirstWorksheet,
  excelCellValue,
  worksheetToJsonRows,
  resolveEmployeeForDeductionRow,
  resolveEmployeeByCodeOrName,
  writeWorkbook,
  writeTemplate,
  downloadJson,
  downloadText,
  employeeToImportRow,
  deductionToExportRow,
  payrollInputToExportRow,
  financeRecordToExportRow,
  assetRecordToExportRow,
  documentRecordToExportRow,
  payrollToExportRow,
  buildPayrollRows,
  validateEmployee,
  validateDeduction,
  validatePayrollInput,
  validateWpsPayroll,
  buildSifRecords,
}
