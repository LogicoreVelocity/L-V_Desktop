import { useMemo, useState, type ChangeEvent } from 'react'
import { jsPDF } from 'jspdf'
import autoTable from 'jspdf-autotable'
import {
  storageKey,
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
  rowValue,
  numberValue,
  normalizeMonth,
  normalizeYear,
  normalizeDate,
  formatEmiratesId,
  roleCategory,
  monthName,
  formatMoney,
  normalizeWpsId,
  normalizeWpsRoutingCode,
  normalizeWpsAccount,
  selectOptions,
  validateExcelFile,
  findMissingColumns,
  readWorkbook,
  readFirstWorksheet,
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
  buildSifRecords,
} from './core/logic'
import type {
  PageKey,
  Employee,
  Deduction,
  PayrollInput,
  FinanceRecord,
  AssetRecord,
  DocumentRecord,
  Settings,
  StoredState,
  PayrollRow,
  SifRecordSet,
  Notice,
} from './core/logic'
import './App.css'

function App() {
  const [storedState] = useState<StoredState>(() => getInitialState())
  const [activePage, setActivePage] = useState<PageKey>('dashboard')
  const [employees, setEmployees] = useState<Employee[]>(storedState.employees)
  const [financeRecords, setFinanceRecords] = useState<FinanceRecord[]>(
    storedState.financeRecords,
  )
  const [assetRecords, setAssetRecords] = useState<AssetRecord[]>(
    storedState.assetRecords,
  )
  const [documentRecords, setDocumentRecords] = useState<DocumentRecord[]>(
    storedState.documentRecords,
  )
  const [payrollInputs, setPayrollInputs] = useState<PayrollInput[]>(
    storedState.payrollInputs,
  )
  const [deductions, setDeductions] = useState<Deduction[]>(storedState.deductions)
  const [settings, setSettings] = useState<Settings>(storedState.settings)
  const [employeeSearch, setEmployeeSearch] = useState('')
  const [employeeStatusFilter, setEmployeeStatusFilter] = useState('All')
  const [financeSearch, setFinanceSearch] = useState('')
  const [assetSearch, setAssetSearch] = useState('')
  const [documentSearch, setDocumentSearch] = useState('')
  const [payrollInputSearch, setPayrollInputSearch] = useState('')
  const [deductionSearch, setDeductionSearch] = useState('')
  const [payrollMonth, setPayrollMonth] = useState(settings.defaultPayrollMonth)
  const [payrollYear, setPayrollYear] = useState(settings.defaultPayrollYear)
  const [slipMonth, setSlipMonth] = useState(settings.defaultPayrollMonth)
  const [slipYear, setSlipYear] = useState(settings.defaultPayrollYear)
  const [selectedSlipEmployeeId, setSelectedSlipEmployeeId] = useState('')
  const [draftEmployee, setDraftEmployee] = useState<Employee>(blankEmployee)
  const [draftFinanceRecord, setDraftFinanceRecord] = useState<FinanceRecord>(
    blankFinanceRecord,
  )
  const [draftAssetRecord, setDraftAssetRecord] = useState<AssetRecord>(
    blankAssetRecord,
  )
  const [draftDocumentRecord, setDraftDocumentRecord] = useState<DocumentRecord>(
    blankDocumentRecord,
  )
  const [draftPayrollInput, setDraftPayrollInput] = useState<PayrollInput>(() =>
    blankPayrollInput(settings.defaultPayrollMonth, settings.defaultPayrollYear),
  )
  const [draftDeduction, setDraftDeduction] = useState<Deduction>(() =>
    blankDeduction(settings.defaultPayrollMonth, settings.defaultPayrollYear),
  )
  const [notice, setNotice] = useState<Notice | null>(null)

  const persist = (next: Partial<StoredState>) => {
    const state = {
      employees,
      financeRecords,
      assetRecords,
      documentRecords,
      payrollInputs,
      deductions,
      settings,
      ...next,
    }
    saveState(state)
  }

  const updateEmployees = (nextEmployees: Employee[]) => {
    setEmployees(nextEmployees)
    persist({ employees: nextEmployees })
  }

  const updateFinanceRecords = (nextFinanceRecords: FinanceRecord[]) => {
    setFinanceRecords(nextFinanceRecords)
    persist({ financeRecords: nextFinanceRecords })
  }

  const updateAssetRecords = (nextAssetRecords: AssetRecord[]) => {
    setAssetRecords(nextAssetRecords)
    persist({ assetRecords: nextAssetRecords })
  }

  const updateDocumentRecords = (nextDocumentRecords: DocumentRecord[]) => {
    setDocumentRecords(nextDocumentRecords)
    persist({ documentRecords: nextDocumentRecords })
  }

  const updatePayrollInputs = (nextPayrollInputs: PayrollInput[]) => {
    setPayrollInputs(nextPayrollInputs)
    persist({ payrollInputs: nextPayrollInputs })
  }

  const updateDeductions = (nextDeductions: Deduction[]) => {
    setDeductions(nextDeductions)
    persist({ deductions: nextDeductions })
  }

  const updateSettings = (nextSettings: Settings) => {
    setSettings(nextSettings)
    persist({ settings: nextSettings })
  }

  const validationSummary = useMemo(() => {
    const errors: string[] = []
    const warnings: string[] = []
    employees.forEach((employee) => {
      const result = validateEmployee(employee, employees, employee.id)
      result.errors.forEach((error) =>
        errors.push(`${employee.employeeCode || employee.employeeName || 'Employee'}: ${error}`),
      )
      result.warnings.forEach((warning) =>
        warnings.push(`${employee.employeeCode || employee.employeeName || 'Employee'}: ${warning}`),
      )
    })
    return { errors, warnings }
  }, [employees])

  const payrollRows = useMemo(
    () => buildPayrollRows(employees, payrollInputs, deductions, payrollMonth, payrollYear),
    [deductions, employees, payrollInputs, payrollMonth, payrollYear],
  )

  const slipRows = useMemo(
    () => buildPayrollRows(employees, payrollInputs, deductions, slipMonth, slipYear),
    [deductions, employees, payrollInputs, slipMonth, slipYear],
  )

  const dashboardRows = useMemo(
    () =>
      buildPayrollRows(
        employees,
        payrollInputs,
        deductions,
        settings.defaultPayrollMonth,
        settings.defaultPayrollYear,
      ),
    [
      deductions,
      employees,
      payrollInputs,
      settings.defaultPayrollMonth,
      settings.defaultPayrollYear,
    ],
  )

  const dashboardStats = useMemo(() => {
    const totalGrossSalary = dashboardRows.reduce((sum, row) => sum + row.grossSalary, 0)
    const totalDeductions = dashboardRows.reduce(
      (sum, row) => sum + row.totalDeduction,
      0,
    )
    const totalNetSalary = dashboardRows.reduce((sum, row) => sum + row.netSalary, 0)
    const financeNet = financeRecords.reduce(
      (sum, record) =>
        sum + (record.type === 'Income' ? record.amount : -record.amount),
      0,
    )
    const assetValue = assetRecords.reduce((sum, record) => sum + record.purchaseCost, 0)

    return {
      totalEmployees: employees.length,
      financeRecords: financeRecords.length,
      assetRecords: assetRecords.length,
      documentRecords: documentRecords.length,
      financeNet,
      assetValue,
      activeEmployees: employees.filter((employee) => employee.status === 'Active').length,
      inactiveEmployees: employees.filter((employee) => employee.status === 'Inactive')
        .length,
      totalGrossSalary,
      totalDeductions,
      totalNetSalary,
      missingEmiratesId: employees.filter(
        (employee) =>
          !employee.emiratesId || !employee.emiratesIdExpiry || !employee.visaExpiryDate,
      ).length,
      missingPassportNumber: employees.filter(
        (employee) => !employee.passportNumber || !employee.passportExpiry,
      ).length,
      missingEmergencyContact: employees.filter(
        (employee) => !employee.emergencyContactName || !employee.emergencyContactMobile,
      ).length,
      missingLicenseInfo: employees.filter(
        (employee) =>
          (roleCategory(employee.roleDesignation) === 'Driver' ||
            roleCategory(employee.roleDesignation) === 'Rider') &&
          (!employee.licenseNumber || !employee.licenseExpiry),
      ).length,
      missingWorkHours: employees.filter((employee) => !employee.workHours).length,
    }
  }, [assetRecords, dashboardRows, documentRecords, employees, financeRecords])

  const filteredEmployees = useMemo(() => {
    const query = employeeSearch.toLowerCase().trim()
    return employees.filter((employee) => {
      const matchesStatus =
        employeeStatusFilter === 'All' || employee.status === employeeStatusFilter
      const matchesSearch =
        !query ||
        [
          employee.employeeCode,
          employee.employeeName,
          employee.roleDesignation,
          employee.mobileNumber,
          employee.email,
          employee.emiratesId,
          employee.passportNumber,
        ]
          .join(' ')
          .toLowerCase()
          .includes(query)
      return matchesStatus && matchesSearch
    })
  }, [employeeSearch, employeeStatusFilter, employees])

  const filteredFinanceRecords = useMemo(() => {
    const query = financeSearch.toLowerCase().trim()
    return financeRecords.filter((record) => {
      const haystack = [
        record.date,
        record.type,
        record.category,
        record.party,
        record.referenceNo,
        record.description,
        record.paymentMode,
        record.status,
      ]
        .join(' ')
        .toLowerCase()
      return !query || haystack.includes(query)
    })
  }, [financeRecords, financeSearch])

  const filteredAssetRecords = useMemo(() => {
    const query = assetSearch.toLowerCase().trim()
    return assetRecords.filter((record) => {
      const employee = employees.find((item) => item.id === record.assignedToEmployeeId)
      const haystack = [
        record.assetCode,
        record.assetName,
        record.category,
        employee?.employeeCode,
        employee?.employeeName,
        record.serialNumber,
        record.location,
        record.condition,
        record.status,
      ]
        .join(' ')
        .toLowerCase()
      return !query || haystack.includes(query)
    })
  }, [assetRecords, assetSearch, employees])

  const filteredDocumentRecords = useMemo(() => {
    const query = documentSearch.toLowerCase().trim()
    return documentRecords.filter((record) => {
      const haystack = [
        record.documentDate,
        record.documentType,
        record.linkedModule,
        record.linkedReference,
        record.title,
        record.fileName,
        record.notes,
      ]
        .join(' ')
        .toLowerCase()
      return !query || haystack.includes(query)
    })
  }, [documentRecords, documentSearch])

  const filteredPayrollInputs = useMemo(() => {
    const query = payrollInputSearch.toLowerCase().trim()
    return payrollInputs.filter((input) => {
      const employee = employees.find((item) => item.id === input.employeeId)
      const haystack = [
        employee?.employeeCode,
        employee?.payrollCode,
        employee?.employeeName,
        monthName(input.month),
        input.year,
        input.remarks,
      ]
        .join(' ')
        .toLowerCase()
      return !query || haystack.includes(query)
    })
  }, [employees, payrollInputSearch, payrollInputs])

  const filteredDeductions = useMemo(() => {
    const query = deductionSearch.toLowerCase().trim()
    return deductions.filter((deduction) => {
      const employee = employees.find((item) => item.id === deduction.employeeId)
      const haystack = [
        employee?.employeeCode,
        employee?.employeeName,
        deduction.reasonRemarks,
        deduction.referenceNo,
        deduction.month,
        deduction.year,
      ]
        .join(' ')
        .toLowerCase()
      return !query || haystack.includes(query)
    })
  }, [deductionSearch, deductions, employees])

  const deductionMonthlySummary = useMemo(() => {
    const byPeriod = new Map<string, number>()
    deductions.forEach((deduction) => {
      const key = `${monthName(deduction.month)} ${deduction.year}`
      const amount =
        deduction.advanceDeduction +
        deduction.fineDeduction +
        deduction.salikDeduction +
        deduction.otherDeduction
      byPeriod.set(key, (byPeriod.get(key) ?? 0) + amount)
    })
    return [...byPeriod.entries()]
  }, [deductions])

  const wpsPreview = useMemo(
    () => buildSifRecords(payrollRows, settings, payrollMonth, payrollYear),
    [payrollRows, payrollMonth, payrollYear, settings],
  )

  const selectedSlip = slipRows.find((row) => row.id === selectedSlipEmployeeId) ?? slipRows[0]

  const setEmployeeField = (employeeId: string, key: keyof Employee, value: string) => {
    const nextEmployees = employees.map((employee) =>
      employee.id === employeeId
        ? {
            ...employee,
            [key]:
              key === 'emiratesId'
                ? formatEmiratesId(value)
                : key === 'wpsPersonId'
                  ? normalizeWpsId(value, 14)
                : key === 'wpsRoutingCode'
                  ? normalizeWpsRoutingCode(value)
                  : key === 'wpsAccountNumber'
                    ? normalizeWpsAccount(value)
                : key === 'basicSalary' ||
                    key === 'allowance' ||
                    key === 'accommodationAllowance' ||
                    key === 'travelAllowance' ||
                    key === 'incentive' ||
                    key === 'dueFromLastMonth' ||
                    key === 'daysPresent' ||
                    key === 'workingDays'
                  ? numberValue(value)
                  : value,
          }
        : employee,
    )
    updateEmployees(nextEmployees)
  }

  const addEmployee = () => {
    const employee = { ...draftEmployee, id: crypto.randomUUID() }
    const result = validateEmployee(employee, employees)
    if (result.errors.length > 0) {
      setNotice({
        tone: 'error',
        message: 'Employee could not be added.',
        details: result.errors,
      })
      return
    }

    updateEmployees([...employees, employee])
    setDraftEmployee(blankEmployee())
    setNotice({
      tone: result.warnings.length ? 'warning' : 'success',
      message: result.warnings.length ? 'Employee added with warnings.' : 'Employee added.',
      details: result.warnings,
    })
  }

  const deleteEmployee = (employeeId: string) => {
    const employee = employees.find((item) => item.id === employeeId)
    if (!employee) return

    if (
      !window.confirm(
        `Delete ${employee.employeeName} and related payroll/deduction rows from local storage?`,
      )
    ) {
      return
    }

    const nextEmployees = employees.filter((item) => item.id !== employeeId)
    const nextPayrollInputs = payrollInputs.filter((input) => input.employeeId !== employeeId)
    const nextAssetRecords = assetRecords.map((asset) =>
      asset.assignedToEmployeeId === employeeId
        ? { ...asset, assignedToEmployeeId: '' }
        : asset,
    )
    const nextDeductions = deductions.filter(
      (deduction) => deduction.employeeId !== employeeId,
    )
    setEmployees(nextEmployees)
    setAssetRecords(nextAssetRecords)
    setPayrollInputs(nextPayrollInputs)
    setDeductions(nextDeductions)
    saveState({
      employees: nextEmployees,
      financeRecords,
      assetRecords: nextAssetRecords,
      documentRecords,
      payrollInputs: nextPayrollInputs,
      deductions: nextDeductions,
      settings,
    })
  }

  const handleEmployeeImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const fileError = validateExcelFile(file)
      if (fileError) {
        setNotice({ tone: 'error', message: fileError, details: [] })
        return
      }

      const workbook = await readWorkbook(file)
      const sheet = readFirstWorksheet(workbook)
      const rows = worksheetToJsonRows(sheet)
      const headers = rows[0] ? Object.keys(rows[0]) : []
      const missingColumns = findMissingColumns(headers, employeeTemplateColumns)

      if (!rows.length) {
        setNotice({ tone: 'error', message: 'No employee rows found.', details: [] })
        return
      }

      if (missingColumns.length > 0) {
        setNotice({
          tone: 'error',
          message: 'Employee import template columns are missing.',
          details: missingColumns.map((column) => `Missing column: ${column}`),
        })
        return
      }

      const imported: Employee[] = []
      const errors: string[] = []
      const warnings: string[] = []
      let comparisonSet = [...employees]

      rows.forEach((row, index) => {
        const employee: Employee = {
          id: crypto.randomUUID(),
          employeeCode: text(row['Employee Code']),
          payrollCode: text(row['Payroll Code']),
          wpsPersonId: normalizeWpsId(text(row['WPS Person ID']), 14),
          wpsRoutingCode: normalizeWpsRoutingCode(text(row['WPS Agent Routing Code'])),
          wpsAccountNumber: normalizeWpsAccount(text(row['WPS IBAN / Account'])),
          employeeName: text(row['Employee Name']),
          roleDesignation: text(row['Role / Designation']),
          dateOfJoining: normalizeDate(row['Date of Joining']),
          mobileNumber: text(row['Mobile Number']),
          basicSalary: numberValue(row['Basic Salary']),
          status: text(row.Status) || 'Active',
          email: text(row.Email),
          emiratesId: formatEmiratesId(text(row['Emirates ID'])),
          emiratesIdExpiry: normalizeDate(row['Emirates ID Expiry']),
          visaExpiryDate: normalizeDate(row['Visa Expiry Date']),
          passportNumber: text(row['Passport Number']),
          passportExpiry: normalizeDate(row['Passport Expiry']),
          emergencyContactName: text(row['Emergency Contact Name']),
          emergencyContactMobile: text(row['Emergency Contact Mobile']),
          emergencyContactName2: text(row['Emergency Contact Name 2']),
          emergencyContactMobile2: text(row['Emergency Contact Mobile 2']),
          licenseNumber: text(row['License Number']),
          licenseExpiry: normalizeDate(row['License Expiry']),
          permitNumber: text(row['Permit Number']),
          permitExpiry: normalizeDate(row['Permit Expiry']),
          workHours: text(row['Work Hours']),
          salaryType:
            text(row['Salary Type']) === 'Per Day' ? 'Per Day' : 'Fixed Salary',
          daysPresent: numberValue(row['Days Present']),
          workingDays: numberValue(row['Working Days']),
          accommodationAllowance: numberValue(row['Accommodation Allowance']),
          travelAllowance: numberValue(row['Travel Allowance']),
          incentive: numberValue(row.Incentive),
          dueFromLastMonth: numberValue(row['Due From Last Month']),
          allowance: numberValue(row.Allowance),
          remarks: text(row.Remarks),
        }
        const result = validateEmployee(employee, comparisonSet)
        const label = `Row ${index + 2}`

        if (result.errors.length) {
          result.errors.forEach((error) => errors.push(`${label}: ${error}`))
          return
        }

        result.warnings.forEach((warning) => warnings.push(`${label}: ${warning}`))
        imported.push(employee)
        comparisonSet = [...comparisonSet, employee]
      })

      updateEmployees([...employees, ...imported])
      setNotice({
        tone: errors.length ? 'warning' : 'success',
        message: `Imported ${imported.length} employees.`,
        details: [...errors, ...warnings],
      })
    } catch (error) {
      setNotice({
        tone: 'error',
        message: 'Could not import employee Excel file.',
        details: [error instanceof Error ? error.message : 'Unknown import error'],
      })
    } finally {
      event.target.value = ''
    }
  }

  const handleFinanceImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const fileError = validateExcelFile(file)
      if (fileError) {
        setNotice({ tone: 'error', message: fileError, details: [] })
        return
      }

      const workbook = await readWorkbook(file)
      const rows = worksheetToJsonRows(readFirstWorksheet(workbook))
      const headers = rows[0] ? Object.keys(rows[0]) : []
      const missingColumns = findMissingColumns(headers, financeTemplateColumns)

      if (!rows.length) {
        setNotice({ tone: 'error', message: 'No finance rows found.', details: [] })
        return
      }

      if (missingColumns.length > 0) {
        setNotice({
          tone: 'error',
          message: 'Finance import template columns are missing.',
          details: missingColumns.map((column) => `Missing column: ${column}`),
        })
        return
      }

      const imported = rows.map((row) =>
        normalizeStoredFinanceRecord({
          id: crypto.randomUUID(),
          date: normalizeDate(rowValue(row, 'Date')),
          type: text(rowValue(row, 'Type')) || 'Expense',
          category: text(rowValue(row, 'Category')),
          party: text(rowValue(row, 'Party')),
          referenceNo: text(rowValue(row, 'Reference No')),
          description: text(rowValue(row, 'Description')),
          amount: numberValue(rowValue(row, 'Amount')),
          vatAmount: numberValue(rowValue(row, 'VAT Amount')),
          paymentMode: text(rowValue(row, 'Payment Mode')),
          status: text(rowValue(row, 'Status')) || 'Pending',
          attachmentName: text(rowValue(row, 'Attachment Name')),
        }),
      )

      updateFinanceRecords([...financeRecords, ...imported])
      setNotice({
        tone: 'success',
        message: `Imported ${imported.length} finance rows.`,
        details: [],
      })
    } catch (error) {
      setNotice({
        tone: 'error',
        message: 'Could not import finance Excel file.',
        details: [error instanceof Error ? error.message : 'Unknown import error'],
      })
    } finally {
      event.target.value = ''
    }
  }

  const handleAssetImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const fileError = validateExcelFile(file)
      if (fileError) {
        setNotice({ tone: 'error', message: fileError, details: [] })
        return
      }

      const workbook = await readWorkbook(file)
      const rows = worksheetToJsonRows(readFirstWorksheet(workbook))
      const headers = rows[0] ? Object.keys(rows[0]) : []
      const missingColumns = findMissingColumns(headers, assetTemplateColumns)

      if (!rows.length) {
        setNotice({ tone: 'error', message: 'No asset rows found.', details: [] })
        return
      }

      if (missingColumns.length > 0) {
        setNotice({
          tone: 'error',
          message: 'Asset import template columns are missing.',
          details: missingColumns.map((column) => `Missing column: ${column}`),
        })
        return
      }

      const imported = rows.map((row) => {
        const employee = resolveEmployeeByCodeOrName(
          rowValue(row, 'Assigned Employee Code', 'Assigned Employee Name'),
          employees,
        )
        return normalizeStoredAssetRecord({
          id: crypto.randomUUID(),
          assetCode: text(rowValue(row, 'Asset Code')),
          assetName: text(rowValue(row, 'Asset Name')),
          category: text(rowValue(row, 'Category')),
          assignedToEmployeeId: employee?.id ?? '',
          purchaseDate: normalizeDate(rowValue(row, 'Purchase Date')),
          purchaseCost: numberValue(rowValue(row, 'Purchase Cost')),
          serialNumber: text(rowValue(row, 'Serial Number')),
          location: text(rowValue(row, 'Location')),
          condition: text(rowValue(row, 'Condition')) || 'Good',
          status: text(rowValue(row, 'Status')) || 'Active',
          warrantyExpiry: normalizeDate(rowValue(row, 'Warranty Expiry')),
          notes: text(rowValue(row, 'Notes')),
        })
      })

      updateAssetRecords([...assetRecords, ...imported])
      setNotice({
        tone: 'success',
        message: `Imported ${imported.length} asset rows.`,
        details: [],
      })
    } catch (error) {
      setNotice({
        tone: 'error',
        message: 'Could not import asset Excel file.',
        details: [error instanceof Error ? error.message : 'Unknown import error'],
      })
    } finally {
      event.target.value = ''
    }
  }

  const handleDocumentImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const fileError = validateExcelFile(file)
      if (fileError) {
        setNotice({ tone: 'error', message: fileError, details: [] })
        return
      }

      const workbook = await readWorkbook(file)
      const rows = worksheetToJsonRows(readFirstWorksheet(workbook))
      const headers = rows[0] ? Object.keys(rows[0]) : []
      const missingColumns = findMissingColumns(headers, documentTemplateColumns)

      if (!rows.length) {
        setNotice({ tone: 'error', message: 'No document rows found.', details: [] })
        return
      }

      if (missingColumns.length > 0) {
        setNotice({
          tone: 'error',
          message: 'Document import template columns are missing.',
          details: missingColumns.map((column) => `Missing column: ${column}`),
        })
        return
      }

      const imported = rows.map((row) =>
        normalizeStoredDocumentRecord({
          id: crypto.randomUUID(),
          documentDate: normalizeDate(rowValue(row, 'Document Date')),
          documentType: text(rowValue(row, 'Document Type')) || 'PDF',
          linkedModule: text(rowValue(row, 'Linked Module')),
          linkedReference: text(rowValue(row, 'Linked Reference')),
          title: text(rowValue(row, 'Title')),
          fileName: text(rowValue(row, 'File Name')),
          fileSizeKb: numberValue(rowValue(row, 'File Size KB')),
          notes: text(rowValue(row, 'Notes')),
        }),
      )

      updateDocumentRecords([...documentRecords, ...imported])
      setNotice({
        tone: 'success',
        message: `Imported ${imported.length} document rows.`,
        details: [],
      })
    } catch (error) {
      setNotice({
        tone: 'error',
        message: 'Could not import document Excel file.',
        details: [error instanceof Error ? error.message : 'Unknown import error'],
      })
    } finally {
      event.target.value = ''
    }
  }

  const handlePayrollInputImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const fileError = validateExcelFile(file)
      if (fileError) {
        setNotice({ tone: 'error', message: fileError, details: [] })
        return
      }

      const workbook = await readWorkbook(file)
      const sheet = readFirstWorksheet(workbook)
      const rows = worksheetToJsonRows(sheet)
      const headers = rows[0] ? Object.keys(rows[0]) : []
      const missingColumns = findMissingColumns(headers, payrollInputTemplateColumns)
      const imported: PayrollInput[] = []
      const errors: string[] = []
      const warnings: string[] = []

      if (!rows.length) {
        setNotice({ tone: 'error', message: 'No payroll input rows found.', details: [] })
        return
      }

      if (missingColumns.length > 0) {
        setNotice({
          tone: 'error',
          message: 'Payroll input template columns are missing.',
          details: missingColumns.map((column) => `Missing column: ${column}`),
        })
        return
      }

      rows.forEach((row, index) => {
        const resolvedEmployee = resolveEmployeeForDeductionRow(row, employees)
        const employee = resolvedEmployee.employee
        if (!employee) {
          errors.push(
            `Row ${index + 2}: Could not match employee from ${resolvedEmployee.matchedBy} "${resolvedEmployee.lookupValue}".`,
          )
          return
        }

        const payrollInput: PayrollInput = {
          id: crypto.randomUUID(),
          employeeId: employee.id,
          month: normalizeMonth(rowValue(row, 'Month')),
          year: normalizeYear(rowValue(row, 'Year')),
          daysPresent: numberValue(rowValue(row, 'Days Present')),
          workingDays: numberValue(rowValue(row, 'Working Days')),
          allowance: numberValue(rowValue(row, 'Allowance')),
          incentive: numberValue(rowValue(row, 'Incentive')),
          attendanceBonus: numberValue(rowValue(row, 'Attendance Bonus')),
          performanceBonus: numberValue(rowValue(row, 'Performance Bonus / Tip')),
          overtimePay: numberValue(rowValue(row, 'Overtime Pay', 'O/T')),
          otherEarning: numberValue(rowValue(row, 'Other Earnings')),
          arrears: numberValue(
            rowValue(row, 'Arrears / Due From Last Month', 'Arrears', 'Due From Last Month'),
          ),
          remarks: text(rowValue(row, 'Remarks')),
        }
        const result = validatePayrollInput(payrollInput, employees, [
          ...payrollInputs,
          ...imported,
        ])

        if (result.errors.length > 0) {
          result.errors.forEach((error) => errors.push(`Row ${index + 2}: ${error}`))
          return
        }

        result.warnings.forEach((warning) =>
          warnings.push(`Row ${index + 2}: ${warning}`),
        )
        imported.push(payrollInput)
      })

      updatePayrollInputs([...payrollInputs, ...imported])
      setNotice({
        tone: errors.length || warnings.length ? 'warning' : 'success',
        message: `Imported ${imported.length} payroll input rows.`,
        details: [...errors, ...warnings],
      })
    } catch (error) {
      setNotice({
        tone: 'error',
        message: 'Could not import payroll inputs Excel file.',
        details: [error instanceof Error ? error.message : 'Unknown import error'],
      })
    } finally {
      event.target.value = ''
    }
  }

  const handleDeductionImport = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const fileError = validateExcelFile(file)
      if (fileError) {
        setNotice({ tone: 'error', message: fileError, details: [] })
        return
      }

      const workbook = await readWorkbook(file)
      const sheet = readFirstWorksheet(workbook)
      const rows = worksheetToJsonRows(sheet)
      const headers = rows[0] ? Object.keys(rows[0]) : []
      const missingColumns = findMissingColumns(headers, deductionsTemplateColumns)
      const imported: Deduction[] = []
      const errors: string[] = []
      const warnings: string[] = []

      if (!rows.length) {
        setNotice({ tone: 'error', message: 'No deduction rows found.', details: [] })
        return
      }

      if (missingColumns.length > 0) {
        setNotice({
          tone: 'error',
          message: 'Deductions import template columns are missing.',
          details: missingColumns.map((column) => `Missing column: ${column}`),
        })
        return
      }

      rows.forEach((row, index) => {
        const resolvedEmployee = resolveEmployeeForDeductionRow(row, employees)
        const employee = resolvedEmployee.employee
        if (!employee) {
          errors.push(
            `Row ${index + 2}: Could not match employee from ${resolvedEmployee.matchedBy} "${resolvedEmployee.lookupValue}".`,
          )
          return
        }

        const rowEmployeeName = text(row['Employee Name'])
        if (
          rowEmployeeName &&
          rowEmployeeName.toLowerCase() !== employee.employeeName.toLowerCase()
        ) {
          warnings.push(
            `Row ${index + 2}: Employee Name does not match the matched employee record. Using ${resolvedEmployee.matchedBy} match.`,
          )
        }

        const deduction: Deduction = {
          id: crypto.randomUUID(),
          employeeId: employee.id,
          month: normalizeMonth(rowValue(row, 'Month', 'Date')),
          year:
            normalizeYear(rowValue(row, 'Year', 'Date')) || settings.defaultPayrollYear,
          transactionId: text(rowValue(row, 'Transaction ID', 'Invoice No')),
          advanceDeduction: numberValue(rowValue(row, 'Advance Deduction')),
          fineDeduction: numberValue(
            rowValue(row, 'Fine Deduction') || (
              text(rowValue(row, 'Invoice Type')).toLowerCase().includes('miscellaneous')
                ? rowValue(row, 'Amount')
                : ''
            ),
          ),
          fineDate: normalizeDate(rowValue(row, 'Fine Date', 'Date')),
          fineReceiptNo: text(rowValue(row, 'Fine Receipt No')),
          fineTransactionNo: text(rowValue(row, 'Fine Transaction No', 'Invoice No')),
          salikDeduction: numberValue(
            rowValue(row, 'Salik Deduction') || (
              text(rowValue(row, 'Invoice Type')).toLowerCase().includes('toll')
                ? rowValue(row, 'Amount')
                : ''
            ),
          ),
          salikDate: normalizeDate(rowValue(row, 'Salik Date', 'Date')),
          salikReceiptNo: text(rowValue(row, 'Salik Receipt No')),
          salikTransactionNo: text(rowValue(row, 'Salik Transaction No', 'Invoice No')),
          otherDeduction: numberValue(rowValue(row, 'Other Deduction')),
          reasonRemarks: text(rowValue(row, 'Reason / Remarks', 'Invoice Type')),
          referenceNo: text(rowValue(row, 'Reference No', 'Invoice No')),
          attachmentName: text(rowValue(row, 'Attachment Name')),
        }
        const result = validateDeduction(deduction, employees, [...deductions, ...imported], employee.id)
        if (result.errors.length > 0) {
          result.errors.forEach((error) => errors.push(`Row ${index + 2}: ${error}`))
          return
        }

        result.warnings.forEach((warning) =>
          warnings.push(`Row ${index + 2}: ${warning}`),
        )
        imported.push(deduction)
      })

      updateDeductions([...deductions, ...imported])
      setNotice({
        tone: errors.length || warnings.length ? 'warning' : 'success',
        message: `Imported ${imported.length} deduction rows.`,
        details: [...errors, ...warnings],
      })
    } catch (error) {
      setNotice({
        tone: 'error',
        message: 'Could not import deductions Excel file.',
        details: [error instanceof Error ? error.message : 'Unknown import error'],
      })
    } finally {
      event.target.value = ''
    }
  }

  const addDeduction = () => {
    const deduction = { ...draftDeduction, id: crypto.randomUUID() }
    const result = validateDeduction(
      deduction,
      employees,
      deductions,
      draftDeduction.employeeId,
    )

    if (result.errors.length > 0) {
      setNotice({
        tone: 'error',
        message: 'Deduction row could not be added.',
        details: result.errors,
      })
      return
    }

    updateDeductions([...deductions, deduction])
    setDraftDeduction(
      blankDeduction(settings.defaultPayrollMonth, settings.defaultPayrollYear),
    )
    setNotice({
      tone: result.warnings.length ? 'warning' : 'success',
      message: result.warnings.length ? 'Deduction row added with warnings.' : 'Deduction row added.',
      details: result.warnings,
    })
  }

  const addPayrollInput = () => {
    const payrollInput = { ...draftPayrollInput, id: crypto.randomUUID() }
    const result = validatePayrollInput(payrollInput, employees, payrollInputs)

    if (result.errors.length > 0) {
      setNotice({
        tone: 'error',
        message: 'Payroll input row could not be added.',
        details: result.errors,
      })
      return
    }

    updatePayrollInputs([...payrollInputs, payrollInput])
    setDraftPayrollInput(
      blankPayrollInput(settings.defaultPayrollMonth, settings.defaultPayrollYear),
    )
    setNotice({
      tone: result.warnings.length ? 'warning' : 'success',
      message: result.warnings.length
        ? 'Payroll input row added with warnings.'
        : 'Payroll input row added.',
      details: result.warnings,
    })
  }

  const addFinanceRecord = () => {
    const record = { ...draftFinanceRecord, id: crypto.randomUUID() }
    if (!record.date || !record.type || record.amount <= 0) {
      setNotice({
        tone: 'error',
        message: 'Finance row could not be added.',
        details: ['Date, Type, and positive Amount are required.'],
      })
      return
    }

    updateFinanceRecords([...financeRecords, record])
    setDraftFinanceRecord(blankFinanceRecord())
    setNotice({ tone: 'success', message: 'Finance row added.', details: [] })
  }

  const addAssetRecord = () => {
    const record = { ...draftAssetRecord, id: crypto.randomUUID() }
    if (!record.assetCode || !record.assetName) {
      setNotice({
        tone: 'error',
        message: 'Asset row could not be added.',
        details: ['Asset Code and Asset Name are required.'],
      })
      return
    }

    updateAssetRecords([...assetRecords, record])
    setDraftAssetRecord(blankAssetRecord())
    setNotice({ tone: 'success', message: 'Asset row added.', details: [] })
  }

  const addDocumentRecord = () => {
    const record = { ...draftDocumentRecord, id: crypto.randomUUID() }
    if (!record.title && !record.fileName) {
      setNotice({
        tone: 'error',
        message: 'Document row could not be added.',
        details: ['Document title or file name is required.'],
      })
      return
    }

    updateDocumentRecords([...documentRecords, record])
    setDraftDocumentRecord(blankDocumentRecord())
    setNotice({ tone: 'success', message: 'Document row added.', details: [] })
  }

  const exportEmployees = () => {
    writeWorkbook(
      'employees.xlsx',
      'Employees',
      employees.map(employeeToImportRow),
    )
  }

  const exportFinanceRecords = () => {
    writeWorkbook('finance-register.xlsx', 'Finance', financeRecords.map(financeRecordToExportRow))
  }

  const exportAssetRecords = () => {
    writeWorkbook(
      'asset-register.xlsx',
      'Assets',
      assetRecords.map((record) =>
        assetRecordToExportRow(
          record,
          employees.find((employee) => employee.id === record.assignedToEmployeeId),
        ),
      ),
    )
  }

  const exportDocumentRecords = () => {
    writeWorkbook(
      'document-register.xlsx',
      'Documents',
      documentRecords.map(documentRecordToExportRow),
    )
  }

  const exportPayrollInputs = () => {
    writeWorkbook(
      'payroll-inputs.xlsx',
      'Payroll Inputs',
      payrollInputs.map((input) =>
        payrollInputToExportRow(
          input,
          employees.find((employee) => employee.id === input.employeeId),
        ),
      ),
    )
  }

  const exportDeductions = () => {
    writeWorkbook(
      'deductions.xlsx',
      'Deductions',
      deductions.map((deduction) =>
        deductionToExportRow(
          deduction,
          employees.find((employee) => employee.id === deduction.employeeId),
        ),
      ),
    )
  }

  const exportPayroll = () => {
    writeWorkbook(
      `payroll-${monthName(payrollMonth)}-${payrollYear}.xlsx`,
      'Payroll',
      payrollRows.map((row, index) => payrollToExportRow(row, index)),
    )
  }

  const exportPayrollTemplate = () => {
    writeTemplate('payroll-export-template.xlsx', 'Payroll', payrollExportColumns)
  }

  const exportSifFile = () => {
    const sif = buildSifRecords(payrollRows, settings, payrollMonth, payrollYear)
    if (sif.validation.errors.length > 0) {
      setNotice({
        tone: 'error',
        message: 'SIF file could not be exported.',
        details: sif.validation.errors,
      })
      return
    }

    downloadText(sif.filename, sif.content, 'text/plain;charset=utf-8')
    setNotice({
      tone: sif.validation.warnings.length ? 'warning' : 'success',
      message: `SIF file exported: ${sif.filename}`,
      details: sif.validation.warnings,
    })
  }

  const exportSlipPdf = (row: PayrollRow) => {
    const doc = new jsPDF()
    const generatedDate = new Date().toLocaleDateString()

    doc.setFontSize(15)
    doc.text(settings.companyName || 'Logicore Excel Manager', 14, 18)
    doc.setFontSize(10)
    if (settings.companyAddress) doc.text(settings.companyAddress, 14, 25)
    if (settings.companyPhone || settings.companyEmail) {
      doc.text(
        [settings.companyPhone, settings.companyEmail].filter(Boolean).join(' | '),
        14,
        settings.companyAddress ? 32 : 25,
      )
    }
    doc.text(
      'Powered by Logicore Velocity',
      14,
      settings.companyAddress || settings.companyPhone || settings.companyEmail ? 38 : 25,
    )

    doc.setFontSize(13)
    doc.text('Salary Slip', 14, 48)
    doc.setFontSize(10)
    doc.text(`Payroll Month: ${monthName(row.month)} ${row.year}`, 14, 56)
    doc.text(`Generated Date: ${generatedDate}`, 14, 62)

    autoTable(doc, {
      startY: 70,
      head: [['Field', 'Value']],
      body: [
        ['Company Name', settings.companyName],
        ['Payroll Month', `${monthName(row.month)} ${row.year}`],
        ['Employee Code', row.employeeCode],
        ['Payroll Code', row.payrollCode || row.employeeCode],
        ['Employee Name', row.employeeName],
        ['Role / Designation', row.roleDesignation],
        ['Date of Joining', row.dateOfJoining],
        ['Mobile Number', row.mobileNumber],
        ['Emirates ID', row.emiratesId],
        ['Salary Type', row.salaryType],
        ['Basic Salary', formatMoney(row.basicSalary, settings.currency)],
        ['Accommodation Allowance', formatMoney(row.accommodationAllowance, settings.currency)],
        ['Travel Allowance', formatMoney(row.travelAllowance, settings.currency)],
        ['Allowance', formatMoney(row.allowance, settings.currency)],
        ['Incentive', formatMoney(row.incentive, settings.currency)],
        ['Attendance Bonus', formatMoney(row.attendanceBonus, settings.currency)],
        ['Performance Bonus / Tip', formatMoney(row.performanceBonus, settings.currency)],
        ['Overtime Pay', formatMoney(row.overtimePay, settings.currency)],
        ['Other Earnings', formatMoney(row.otherEarning, settings.currency)],
        ['Salary Earned', formatMoney(row.salaryEarned, settings.currency)],
        ['Gross Salary', formatMoney(row.grossSalary, settings.currency)],
        ['Arrears / Due From Last Month', formatMoney(row.dueFromLastMonth, settings.currency)],
        ['Payable Before Deductions', formatMoney(row.payableBeforeDeductions, settings.currency)],
        ['Advance Deduction', formatMoney(row.advanceDeduction, settings.currency)],
        ['Fine Deduction', formatMoney(row.fineDeduction, settings.currency)],
        ['Salik Deduction', formatMoney(row.salikDeduction, settings.currency)],
        ['Other Deduction', formatMoney(row.otherDeduction, settings.currency)],
        ['Total Deduction', formatMoney(row.totalDeduction, settings.currency)],
        ['Net Salary', formatMoney(row.netSalary, settings.currency)],
        ['Work Hours', row.workHours],
        ['Remarks', row.remarks],
      ],
    })

    doc.setFontSize(9)
    doc.text(
      'Generated locally by Logicore Excel Manager.',
      14,
      286,
    )
    doc.save(`salary-slip-${row.employeeCode}-${row.month}-${row.year}.pdf`)
  }

  const exportAllSlipPdfs = () => {
    slipRows.forEach((row) => exportSlipPdf(row))
  }

  const restoreBackup = async (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]
    if (!file) return

    try {
      const parsed = JSON.parse(await file.text()) as Partial<StoredState>
      const nextState: StoredState = {
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

      setEmployees(nextState.employees)
      setFinanceRecords(nextState.financeRecords)
      setAssetRecords(nextState.assetRecords)
      setDocumentRecords(nextState.documentRecords)
      setPayrollInputs(nextState.payrollInputs)
      setDeductions(nextState.deductions)
      setSettings(nextState.settings)
      setPayrollMonth(nextState.settings.defaultPayrollMonth)
      setPayrollYear(nextState.settings.defaultPayrollYear)
      setSlipMonth(nextState.settings.defaultPayrollMonth)
      setSlipYear(nextState.settings.defaultPayrollYear)
      saveState(nextState)
      setNotice({ tone: 'success', message: 'Backup restored.', details: [] })
    } catch (error) {
      setNotice({
        tone: 'error',
        message: 'Could not restore backup JSON.',
        details: [error instanceof Error ? error.message : 'Unknown restore error'],
      })
    } finally {
      event.target.value = ''
    }
  }

  const clearLocalData = () => {
    if (
      !window.confirm(
        'Clear all local employees, deductions, payroll settings, and backup state from this device?',
      )
    ) {
      return
    }

    localStorage.removeItem(storageKey)
    setEmployees([])
    setFinanceRecords([])
    setAssetRecords([])
    setDocumentRecords([])
    setPayrollInputs([])
    setDeductions([])
    setSettings(defaultSettings)
    setPayrollMonth(defaultSettings.defaultPayrollMonth)
    setPayrollYear(defaultSettings.defaultPayrollYear)
    setSlipMonth(defaultSettings.defaultPayrollMonth)
    setSlipYear(defaultSettings.defaultPayrollYear)
    setSelectedSlipEmployeeId('')
    setNotice({ tone: 'success', message: 'Local data cleared.', details: [] })
  }

  const renderNotice = () => {
    if (!notice) return null

    return (
      <div className={`notice ${notice.tone}`}>
        <strong>{notice.message}</strong>
        {notice.details.length > 0 && (
          <ul>
            {notice.details.slice(0, 12).map((detail) => (
              <li key={detail}>{detail}</li>
            ))}
          </ul>
        )}
      </div>
    )
  }

  return (
    <div className="app-shell">
      <aside className="sidebar">
        <div className="brand-block">
          <span className="brand-mark">LV</span>
          <div>
            <h1>Logicore Excel Manager</h1>
            <p>Powered by Logicore Velocity</p>
          </div>
        </div>

        <nav className="main-nav" aria-label="Main navigation">
          {pages.map((page) => (
            <button
              key={page.key}
              type="button"
              className={activePage === page.key ? 'active' : ''}
              onClick={() => setActivePage(page.key)}
            >
              {page.label}
            </button>
          ))}
        </nav>

        <footer className="local-footer">
          Logicore Payroll is a local payroll, WPS/SIF, and payslip PDF tool.
          Data stays on this device unless the user exports or shares it.
        </footer>
      </aside>

      <main className="workspace">
        <header className="topbar">
          <div>
            <p className="eyebrow">Powered by Logicore Velocity</p>
            <h2>{pages.find((page) => page.key === activePage)?.label}</h2>
          </div>
          <div className="topbar-actions">
            <button
              type="button"
              className="secondary"
              onClick={() =>
                writeTemplate(
                  'employee-import-template.xlsx',
                  'Employees',
                  employeeTemplateColumns,
                )
              }
            >
              Employee Template
            </button>
          </div>
        </header>

        {renderNotice()}

        {activePage === 'dashboard' && (
          <section className="page-stack">
            <section className="hero-panel">
              <div>
                <p className="eyebrow">Free local Excel payroll tool</p>
                <h3>Employee, driver, deduction, payroll, and salary slip workflows in one simple desktop app.</h3>
                <p>
                  Built to help small employers, delivery companies, manpower teams,
                  fleet operators, and 3PL companies run a simple payroll
                  workflow without cloud storage or tenant setup, while still
                  staying usable as a local web app or packaged desktop app.
                </p>
              </div>
              <button type="button" onClick={() => setActivePage('employees')}>
                Start With Employees
              </button>
            </section>

            <div className="summary-grid twelve">
              <SummaryCard label="Total Employees" value={dashboardStats.totalEmployees} />
              <SummaryCard label="Finance Rows" value={dashboardStats.financeRecords} />
              <SummaryCard label="Asset Rows" value={dashboardStats.assetRecords} />
              <SummaryCard label="Document Rows" value={dashboardStats.documentRecords} />
              <SummaryCard label="Active Employees" value={dashboardStats.activeEmployees} />
              <SummaryCard label="Inactive Employees" value={dashboardStats.inactiveEmployees} />
              <SummaryCard
                label="Finance Net"
                value={formatMoney(dashboardStats.financeNet, settings.currency)}
              />
              <SummaryCard
                label="Asset Value"
                value={formatMoney(dashboardStats.assetValue, settings.currency)}
              />
              <SummaryCard
                label="Total Gross Salary"
                value={formatMoney(dashboardStats.totalGrossSalary, settings.currency)}
              />
              <SummaryCard
                label="Total Deductions"
                value={formatMoney(dashboardStats.totalDeductions, settings.currency)}
              />
              <SummaryCard
                label="Total Net Salary"
                value={formatMoney(dashboardStats.totalNetSalary, settings.currency)}
              />
              <SummaryCard label="Missing Emirates ID" value={dashboardStats.missingEmiratesId} />
              <SummaryCard
                label="Missing Passport Number"
                value={dashboardStats.missingPassportNumber}
              />
              <SummaryCard
                label="Missing Emergency Contact"
                value={dashboardStats.missingEmergencyContact}
              />
              <SummaryCard
                label="Missing License Info"
                value={dashboardStats.missingLicenseInfo}
              />
              <SummaryCard
                label="Missing Work Hours"
                value={dashboardStats.missingWorkHours}
              />
            </div>

            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Payroll Preview</h3>
                  <p>
                    Gross Salary = Salary Earned using salary type and allowance
                    components. Payable Before Deductions = Gross Salary + Due From
                    Last Month. Net Salary = Payable Before Deductions - Advance -
                    Fine - Salik - Other Deduction.
                  </p>
                </div>
                <button type="button" onClick={exportPayroll} disabled={!payrollRows.length}>
                  Export Payroll
                </button>
              </div>
              <PayrollTable rows={dashboardRows.slice(0, 6)} currency={settings.currency} compact />
            </section>
          </section>
        )}

        {activePage === 'employees' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Employees / Drivers</h3>
                  <p>Import, add, edit, delete, search, filter, and export local employee records.</p>
                </div>
                <div className="inline-actions">
                  <label className="file-button">
                    Import Employee Excel
                    <input type="file" accept=".xlsx" onChange={handleEmployeeImport} />
                  </label>
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      writeTemplate(
                        'employee-import-template.xlsx',
                        'Employees',
                        employeeTemplateColumns,
                      )
                    }
                  >
                    Download Template
                  </button>
                  <button type="button" className="secondary" onClick={exportEmployees}>
                    Export Employee List
                  </button>
                </div>
              </div>

              <ValidationPanel
                errors={validationSummary.errors}
                warnings={validationSummary.warnings}
              />

              <div className="toolbar">
                <input
                  type="search"
                  value={employeeSearch}
                  onChange={(event) => setEmployeeSearch(event.target.value)}
                  placeholder="Search employee, role, mobile, Emirates ID, passport..."
                />
                <select
                  value={employeeStatusFilter}
                  onChange={(event) => setEmployeeStatusFilter(event.target.value)}
                >
                  <option value="All">All Status</option>
                  <option value="Active">Active</option>
                  <option value="Inactive">Inactive</option>
                </select>
              </div>

              <EditableEmployeeTable
                employees={filteredEmployees}
                onCellChange={setEmployeeField}
                onDelete={deleteEmployee}
              />
            </section>

            <section className="panel">
              <h3>Add Employee Manually</h3>
              <div className="form-grid employee-form">
                {employeeFields.map((field) => (
                  <label key={field.key}>
                    {field.label}
                    {field.type === 'select' ? (
                      <select
                        value={String(draftEmployee[field.key])}
                        onChange={(event) =>
                          setDraftEmployee({
                            ...draftEmployee,
                            [field.key]: event.target.value,
                          })
                        }
                      >
                        {selectOptions(field.key).map((option) => (
                          <option key={option} value={option}>
                            {option}
                          </option>
                        ))}
                      </select>
                    ) : (
                      <input
                        type={field.type}
                        value={String(draftEmployee[field.key])}
                        onChange={(event) =>
                          setDraftEmployee({
                            ...draftEmployee,
                            [field.key]:
                              field.key === 'emiratesId'
                                ? formatEmiratesId(event.target.value)
                                : field.key === 'wpsPersonId'
                                  ? normalizeWpsId(event.target.value, 14)
                                  : field.key === 'wpsRoutingCode'
                                    ? normalizeWpsRoutingCode(event.target.value)
                                    : field.key === 'wpsAccountNumber'
                                      ? normalizeWpsAccount(event.target.value)
                                : field.type === 'number'
                                ? numberValue(event.target.value)
                                : event.target.value,
                          })
                        }
                      />
                    )}
                  </label>
                ))}
              </div>
              <button type="button" onClick={addEmployee}>
                Add Employee
              </button>
            </section>
          </section>
        )}

        {activePage === 'finance' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Finance Register</h3>
                  <p>Track income, expenses, references, VAT, payment mode, status, and attachment names.</p>
                </div>
                <div className="inline-actions">
                  <label className="file-button">
                    Import Finance Excel
                    <input type="file" accept=".xlsx" onChange={handleFinanceImport} />
                  </label>
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      writeTemplate('finance-register-template.xlsx', 'Finance', financeTemplateColumns)
                    }
                  >
                    Download Template
                  </button>
                  <button type="button" className="secondary" onClick={exportFinanceRecords}>
                    Export Finance
                  </button>
                </div>
              </div>
              <div className="toolbar">
                <input
                  type="search"
                  value={financeSearch}
                  onChange={(event) => setFinanceSearch(event.target.value)}
                  placeholder="Search finance records..."
                />
              </div>
              <FinanceTable
                records={filteredFinanceRecords}
                currency={settings.currency}
                onDelete={(recordId) =>
                  updateFinanceRecords(financeRecords.filter((item) => item.id !== recordId))
                }
              />
            </section>

            <section className="panel">
              <h3>Add Finance Row</h3>
              <div className="form-grid deductions-form">
                <label>
                  Date
                  <input
                    type="date"
                    value={draftFinanceRecord.date}
                    onChange={(event) =>
                      setDraftFinanceRecord({ ...draftFinanceRecord, date: event.target.value })
                    }
                  />
                </label>
                <label>
                  Type
                  <select
                    value={draftFinanceRecord.type}
                    onChange={(event) =>
                      setDraftFinanceRecord({ ...draftFinanceRecord, type: event.target.value })
                    }
                  >
                    <option value="Expense">Expense</option>
                    <option value="Income">Income</option>
                    <option value="Transfer">Transfer</option>
                  </select>
                </label>
                <label>
                  Category
                  <input
                    value={draftFinanceRecord.category}
                    onChange={(event) =>
                      setDraftFinanceRecord({ ...draftFinanceRecord, category: event.target.value })
                    }
                  />
                </label>
                <label>
                  Party
                  <input
                    value={draftFinanceRecord.party}
                    onChange={(event) =>
                      setDraftFinanceRecord({ ...draftFinanceRecord, party: event.target.value })
                    }
                  />
                </label>
                <label>
                  Reference No
                  <input
                    value={draftFinanceRecord.referenceNo}
                    onChange={(event) =>
                      setDraftFinanceRecord({
                        ...draftFinanceRecord,
                        referenceNo: event.target.value,
                      })
                    }
                  />
                </label>
                <NumberField
                  label="Amount"
                  value={draftFinanceRecord.amount}
                  onChange={(value) =>
                    setDraftFinanceRecord({ ...draftFinanceRecord, amount: value })
                  }
                />
                <NumberField
                  label="VAT Amount"
                  value={draftFinanceRecord.vatAmount}
                  onChange={(value) =>
                    setDraftFinanceRecord({ ...draftFinanceRecord, vatAmount: value })
                  }
                />
                <label>
                  Payment Mode
                  <input
                    value={draftFinanceRecord.paymentMode}
                    onChange={(event) =>
                      setDraftFinanceRecord({
                        ...draftFinanceRecord,
                        paymentMode: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Status
                  <select
                    value={draftFinanceRecord.status}
                    onChange={(event) =>
                      setDraftFinanceRecord({ ...draftFinanceRecord, status: event.target.value })
                    }
                  >
                    <option value="Pending">Pending</option>
                    <option value="Paid">Paid</option>
                    <option value="Reconciled">Reconciled</option>
                  </select>
                </label>
                <label>
                  Attachment Name
                  <input
                    value={draftFinanceRecord.attachmentName}
                    onChange={(event) =>
                      setDraftFinanceRecord({
                        ...draftFinanceRecord,
                        attachmentName: event.target.value,
                      })
                    }
                  />
                </label>
                <label className="wide">
                  Description
                  <input
                    value={draftFinanceRecord.description}
                    onChange={(event) =>
                      setDraftFinanceRecord({
                        ...draftFinanceRecord,
                        description: event.target.value,
                      })
                    }
                  />
                </label>
              </div>
              <button type="button" onClick={addFinanceRecord}>
                Add Finance Row
              </button>
            </section>
          </section>
        )}

        {activePage === 'assets' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Asset Register</h3>
                  <p>Track company assets, assignment, cost, serial number, warranty, condition, and status.</p>
                </div>
                <div className="inline-actions">
                  <label className="file-button">
                    Import Assets Excel
                    <input type="file" accept=".xlsx" onChange={handleAssetImport} />
                  </label>
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      writeTemplate('asset-register-template.xlsx', 'Assets', assetTemplateColumns)
                    }
                  >
                    Download Template
                  </button>
                  <button type="button" className="secondary" onClick={exportAssetRecords}>
                    Export Assets
                  </button>
                </div>
              </div>
              <div className="toolbar">
                <input
                  type="search"
                  value={assetSearch}
                  onChange={(event) => setAssetSearch(event.target.value)}
                  placeholder="Search assets..."
                />
              </div>
              <AssetTable
                records={filteredAssetRecords}
                employees={employees}
                currency={settings.currency}
                onDelete={(recordId) =>
                  updateAssetRecords(assetRecords.filter((item) => item.id !== recordId))
                }
              />
            </section>

            <section className="panel">
              <h3>Add Asset Row</h3>
              <div className="form-grid employee-form">
                <label>
                  Asset Code
                  <input
                    value={draftAssetRecord.assetCode}
                    onChange={(event) =>
                      setDraftAssetRecord({ ...draftAssetRecord, assetCode: event.target.value })
                    }
                  />
                </label>
                <label>
                  Asset Name
                  <input
                    value={draftAssetRecord.assetName}
                    onChange={(event) =>
                      setDraftAssetRecord({ ...draftAssetRecord, assetName: event.target.value })
                    }
                  />
                </label>
                <label>
                  Category
                  <input
                    value={draftAssetRecord.category}
                    onChange={(event) =>
                      setDraftAssetRecord({ ...draftAssetRecord, category: event.target.value })
                    }
                  />
                </label>
                <label>
                  Assigned Employee
                  <select
                    value={draftAssetRecord.assignedToEmployeeId}
                    onChange={(event) =>
                      setDraftAssetRecord({
                        ...draftAssetRecord,
                        assignedToEmployeeId: event.target.value,
                      })
                    }
                  >
                    <option value="">Unassigned</option>
                    {employees.map((employee) => (
                      <option key={employee.id} value={employee.id}>
                        {employee.employeeCode} - {employee.employeeName}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  Purchase Date
                  <input
                    type="date"
                    value={draftAssetRecord.purchaseDate}
                    onChange={(event) =>
                      setDraftAssetRecord({
                        ...draftAssetRecord,
                        purchaseDate: event.target.value,
                      })
                    }
                  />
                </label>
                <NumberField
                  label="Purchase Cost"
                  value={draftAssetRecord.purchaseCost}
                  onChange={(value) =>
                    setDraftAssetRecord({ ...draftAssetRecord, purchaseCost: value })
                  }
                />
                <label>
                  Serial Number
                  <input
                    value={draftAssetRecord.serialNumber}
                    onChange={(event) =>
                      setDraftAssetRecord({
                        ...draftAssetRecord,
                        serialNumber: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Location
                  <input
                    value={draftAssetRecord.location}
                    onChange={(event) =>
                      setDraftAssetRecord({ ...draftAssetRecord, location: event.target.value })
                    }
                  />
                </label>
                <label>
                  Condition
                  <select
                    value={draftAssetRecord.condition}
                    onChange={(event) =>
                      setDraftAssetRecord({ ...draftAssetRecord, condition: event.target.value })
                    }
                  >
                    <option value="Good">Good</option>
                    <option value="Service Due">Service Due</option>
                    <option value="Damaged">Damaged</option>
                    <option value="Retired">Retired</option>
                  </select>
                </label>
                <label>
                  Status
                  <select
                    value={draftAssetRecord.status}
                    onChange={(event) =>
                      setDraftAssetRecord({ ...draftAssetRecord, status: event.target.value })
                    }
                  >
                    <option value="Active">Active</option>
                    <option value="Maintenance">Maintenance</option>
                    <option value="Disposed">Disposed</option>
                  </select>
                </label>
                <label>
                  Warranty Expiry
                  <input
                    type="date"
                    value={draftAssetRecord.warrantyExpiry}
                    onChange={(event) =>
                      setDraftAssetRecord({
                        ...draftAssetRecord,
                        warrantyExpiry: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Notes
                  <input
                    value={draftAssetRecord.notes}
                    onChange={(event) =>
                      setDraftAssetRecord({ ...draftAssetRecord, notes: event.target.value })
                    }
                  />
                </label>
              </div>
              <button type="button" onClick={addAssetRecord}>
                Add Asset
              </button>
            </section>
          </section>
        )}

        {activePage === 'documents' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Documents / PDFs</h3>
                  <p>Register PDF and document metadata alongside employee, finance, asset, and payroll records.</p>
                </div>
                <div className="inline-actions">
                  <label className="file-button">
                    Import Document Excel
                    <input type="file" accept=".xlsx" onChange={handleDocumentImport} />
                  </label>
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      writeTemplate(
                        'document-register-template.xlsx',
                        'Documents',
                        documentTemplateColumns,
                      )
                    }
                  >
                    Download Template
                  </button>
                  <button type="button" className="secondary" onClick={exportDocumentRecords}>
                    Export Documents
                  </button>
                </div>
              </div>
              <div className="toolbar">
                <input
                  type="search"
                  value={documentSearch}
                  onChange={(event) => setDocumentSearch(event.target.value)}
                  placeholder="Search documents..."
                />
              </div>
              <DocumentTable
                records={filteredDocumentRecords}
                onDelete={(recordId) =>
                  updateDocumentRecords(documentRecords.filter((item) => item.id !== recordId))
                }
              />
            </section>

            <section className="panel">
              <h3>Add Document Row</h3>
              <div className="form-grid deductions-form">
                <label>
                  Document Date
                  <input
                    type="date"
                    value={draftDocumentRecord.documentDate}
                    onChange={(event) =>
                      setDraftDocumentRecord({
                        ...draftDocumentRecord,
                        documentDate: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Document Type
                  <input
                    value={draftDocumentRecord.documentType}
                    onChange={(event) =>
                      setDraftDocumentRecord({
                        ...draftDocumentRecord,
                        documentType: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Linked Module
                  <input
                    value={draftDocumentRecord.linkedModule}
                    onChange={(event) =>
                      setDraftDocumentRecord({
                        ...draftDocumentRecord,
                        linkedModule: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Linked Reference
                  <input
                    value={draftDocumentRecord.linkedReference}
                    onChange={(event) =>
                      setDraftDocumentRecord({
                        ...draftDocumentRecord,
                        linkedReference: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Title
                  <input
                    value={draftDocumentRecord.title}
                    onChange={(event) =>
                      setDraftDocumentRecord({ ...draftDocumentRecord, title: event.target.value })
                    }
                  />
                </label>
                <label>
                  PDF / Document File
                  <input
                    type="file"
                    accept=".pdf,.png,.jpg,.jpeg,.xlsx,.doc,.docx"
                    onChange={(event) => {
                      const file = event.target.files?.[0]
                      setDraftDocumentRecord({
                        ...draftDocumentRecord,
                        fileName: file?.name ?? '',
                        fileSizeKb: file ? Math.round(file.size / 1024) : 0,
                        title: draftDocumentRecord.title || file?.name || '',
                      })
                    }}
                  />
                </label>
                <NumberField
                  label="File Size KB"
                  value={draftDocumentRecord.fileSizeKb}
                  onChange={(value) =>
                    setDraftDocumentRecord({
                      ...draftDocumentRecord,
                      fileSizeKb: value,
                    })
                  }
                />
                <label>
                  Notes
                  <input
                    value={draftDocumentRecord.notes}
                    onChange={(event) =>
                      setDraftDocumentRecord({ ...draftDocumentRecord, notes: event.target.value })
                    }
                  />
                </label>
              </div>
              <button type="button" onClick={addDocumentRecord}>
                Add Document
              </button>
            </section>
          </section>
        )}

        {activePage === 'templates' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Import Templates</h3>
                  <p>Download clean Excel templates for employee import, deductions import, and payroll export structure.</p>
                </div>
              </div>
              <div className="action-grid">
                <button
                  type="button"
                  onClick={() =>
                    writeTemplate(
                      'employee-import-template.xlsx',
                      'Employees',
                      employeeTemplateColumns,
                    )
                  }
                >
                  Employee Import Template
                </button>
                <button
                  type="button"
                  onClick={() =>
                    writeTemplate('finance-register-template.xlsx', 'Finance', financeTemplateColumns)
                  }
                >
                  Finance Register Template
                </button>
                <button
                  type="button"
                  onClick={() =>
                    writeTemplate('asset-register-template.xlsx', 'Assets', assetTemplateColumns)
                  }
                >
                  Asset Register Template
                </button>
                <button
                  type="button"
                  onClick={() =>
                    writeTemplate(
                      'document-register-template.xlsx',
                      'Documents',
                      documentTemplateColumns,
                    )
                  }
                >
                  Document Register Template
                </button>
                <button
                  type="button"
                  onClick={() =>
                    writeTemplate(
                      'deductions-import-template.xlsx',
                      'Deductions',
                      deductionsTemplateColumns,
                    )
                  }
                >
                  Deductions Import Template
                </button>
                <button
                  type="button"
                  onClick={() =>
                    writeTemplate(
                      'payroll-inputs-template.xlsx',
                      'Payroll Inputs',
                      payrollInputTemplateColumns,
                    )
                  }
                >
                  Payroll Inputs Template
                </button>
                <button type="button" onClick={exportPayrollTemplate}>
                  Payroll Export Template
                </button>
              </div>
            </section>
          </section>
        )}

        {activePage === 'inputs' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Payroll Inputs</h3>
                  <p>Import or enter monthly attendance, incentive, bonus, allowance, overtime, other earnings, and arrears before generating payroll.</p>
                </div>
                <div className="inline-actions">
                  <label className="file-button">
                    Import Payroll Inputs
                    <input type="file" accept=".xlsx" onChange={handlePayrollInputImport} />
                  </label>
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      writeTemplate(
                        'payroll-inputs-template.xlsx',
                        'Payroll Inputs',
                        payrollInputTemplateColumns,
                      )
                    }
                  >
                    Download Template
                  </button>
                  <button type="button" className="secondary" onClick={exportPayrollInputs}>
                    Export Inputs
                  </button>
                </div>
              </div>

              <div className="form-grid deductions-form">
                <label>
                  Employee
                  <select
                    value={draftPayrollInput.employeeId}
                    onChange={(event) =>
                      setDraftPayrollInput({
                        ...draftPayrollInput,
                        employeeId: event.target.value,
                      })
                    }
                  >
                    <option value="">Select employee</option>
                    {employees.map((employee) => (
                      <option key={employee.id} value={employee.id}>
                        {employee.employeeCode} - {employee.employeeName}
                      </option>
                    ))}
                  </select>
                </label>
                <MonthYearControls
                  month={draftPayrollInput.month}
                  year={draftPayrollInput.year}
                  onMonthChange={(month) =>
                    setDraftPayrollInput({ ...draftPayrollInput, month })
                  }
                  onYearChange={(year) =>
                    setDraftPayrollInput({ ...draftPayrollInput, year })
                  }
                />
                <NumberField
                  label="Days Present"
                  value={draftPayrollInput.daysPresent}
                  onChange={(value) =>
                    setDraftPayrollInput({ ...draftPayrollInput, daysPresent: value })
                  }
                />
                <NumberField
                  label="Working Days"
                  value={draftPayrollInput.workingDays}
                  onChange={(value) =>
                    setDraftPayrollInput({ ...draftPayrollInput, workingDays: value })
                  }
                />
                <NumberField
                  label="Allowance"
                  value={draftPayrollInput.allowance}
                  onChange={(value) =>
                    setDraftPayrollInput({ ...draftPayrollInput, allowance: value })
                  }
                />
                <NumberField
                  label="Incentive"
                  value={draftPayrollInput.incentive}
                  onChange={(value) =>
                    setDraftPayrollInput({ ...draftPayrollInput, incentive: value })
                  }
                />
                <NumberField
                  label="Attendance Bonus"
                  value={draftPayrollInput.attendanceBonus}
                  onChange={(value) =>
                    setDraftPayrollInput({
                      ...draftPayrollInput,
                      attendanceBonus: value,
                    })
                  }
                />
                <NumberField
                  label="Performance Bonus / Tip"
                  value={draftPayrollInput.performanceBonus}
                  onChange={(value) =>
                    setDraftPayrollInput({
                      ...draftPayrollInput,
                      performanceBonus: value,
                    })
                  }
                />
                <NumberField
                  label="Overtime Pay"
                  value={draftPayrollInput.overtimePay}
                  onChange={(value) =>
                    setDraftPayrollInput({ ...draftPayrollInput, overtimePay: value })
                  }
                />
                <NumberField
                  label="Other Earnings"
                  value={draftPayrollInput.otherEarning}
                  onChange={(value) =>
                    setDraftPayrollInput({ ...draftPayrollInput, otherEarning: value })
                  }
                />
                <NumberField
                  label="Arrears / Due From Last Month"
                  value={draftPayrollInput.arrears}
                  onChange={(value) =>
                    setDraftPayrollInput({ ...draftPayrollInput, arrears: value })
                  }
                />
                <label>
                  Remarks
                  <input
                    value={draftPayrollInput.remarks}
                    onChange={(event) =>
                      setDraftPayrollInput({
                        ...draftPayrollInput,
                        remarks: event.target.value,
                      })
                    }
                  />
                </label>
              </div>
              <button type="button" onClick={addPayrollInput}>
                Add Payroll Input
              </button>
            </section>

            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Monthly Input Rows</h3>
                  <p>One row per employee/month/year overrides the employee defaults for that payroll period.</p>
                </div>
              </div>
              <div className="toolbar">
                <input
                  type="search"
                  value={payrollInputSearch}
                  onChange={(event) => setPayrollInputSearch(event.target.value)}
                  placeholder="Search payroll inputs by employee, period, or remarks..."
                />
              </div>
              <PayrollInputsTable
                inputs={filteredPayrollInputs}
                employees={employees}
                currency={settings.currency}
                onDelete={(inputId) =>
                  updatePayrollInputs(payrollInputs.filter((item) => item.id !== inputId))
                }
              />
            </section>
          </section>
        )}

        {activePage === 'deductions' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Advances & Deductions</h3>
                  <p>Import deduction Excel rows, add manual rows, capture fine and Salik references, attach a PDF filename locally, search, summarize, and export.</p>
                </div>
                <div className="inline-actions">
                  <label className="file-button">
                    Import Deductions Excel
                    <input type="file" accept=".xlsx" onChange={handleDeductionImport} />
                  </label>
                  <button
                    type="button"
                    className="secondary"
                    onClick={() =>
                      writeTemplate(
                        'deductions-import-template.xlsx',
                        'Deductions',
                        deductionsTemplateColumns,
                      )
                    }
                  >
                    Download Template
                  </button>
                  <button type="button" className="secondary" onClick={exportDeductions}>
                    Export Deductions
                  </button>
                </div>
              </div>

              <div className="form-grid deductions-form">
                <label>
                  Employee
                  <select
                    value={draftDeduction.employeeId}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        employeeId: event.target.value,
                      })
                    }
                  >
                    <option value="">Select employee</option>
                    {employees.map((employee) => (
                      <option key={employee.id} value={employee.id}>
                        {employee.employeeCode} - {employee.employeeName}
                      </option>
                    ))}
                  </select>
                </label>
                <MonthYearControls
                  month={draftDeduction.month}
                  year={draftDeduction.year}
                  onMonthChange={(month) =>
                    setDraftDeduction({ ...draftDeduction, month })
                  }
                  onYearChange={(year) => setDraftDeduction({ ...draftDeduction, year })}
                />
                <NumberField
                  label="Advance Deduction"
                  value={draftDeduction.advanceDeduction}
                  onChange={(value) =>
                    setDraftDeduction({ ...draftDeduction, advanceDeduction: value })
                  }
                />
                <NumberField
                  label="Fine Deduction"
                  value={draftDeduction.fineDeduction}
                  onChange={(value) =>
                    setDraftDeduction({ ...draftDeduction, fineDeduction: value })
                  }
                />
                <NumberField
                  label="Salik Deduction"
                  value={draftDeduction.salikDeduction}
                  onChange={(value) =>
                    setDraftDeduction({ ...draftDeduction, salikDeduction: value })
                  }
                />
                <NumberField
                  label="Other Deduction"
                  value={draftDeduction.otherDeduction}
                  onChange={(value) =>
                    setDraftDeduction({ ...draftDeduction, otherDeduction: value })
                  }
                />
                <label>
                  Transaction ID
                  <input
                    value={draftDeduction.transactionId}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        transactionId: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Fine Date
                  <input
                    type="date"
                    value={draftDeduction.fineDate}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        fineDate: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Fine Receipt No
                  <input
                    value={draftDeduction.fineReceiptNo}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        fineReceiptNo: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Fine Transaction No
                  <input
                    value={draftDeduction.fineTransactionNo}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        fineTransactionNo: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Salik Date
                  <input
                    type="date"
                    value={draftDeduction.salikDate}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        salikDate: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Salik Receipt No
                  <input
                    value={draftDeduction.salikReceiptNo}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        salikReceiptNo: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Salik Transaction No
                  <input
                    value={draftDeduction.salikTransactionNo}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        salikTransactionNo: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Reason / Remarks
                  <input
                    value={draftDeduction.reasonRemarks}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        reasonRemarks: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  Reference No
                  <input
                    value={draftDeduction.referenceNo}
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        referenceNo: event.target.value,
                      })
                    }
                  />
                </label>
                <label>
                  PDF Attachment
                  <input
                    type="file"
                    accept=".pdf"
                    onChange={(event) =>
                      setDraftDeduction({
                        ...draftDeduction,
                        attachmentName: event.target.files?.[0]?.name ?? '',
                      })
                    }
                  />
                </label>
              </div>
              <button type="button" onClick={addDeduction}>
                Add Deduction
              </button>
            </section>

            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Deduction Rows</h3>
                  <p>Monthly summary is based on all local deduction rows.</p>
                </div>
              </div>
              <div className="summary-strip">
                {deductionMonthlySummary.length ? (
                  deductionMonthlySummary.map(([period, total]) => (
                    <span key={period}>
                      {period}: {formatMoney(total, settings.currency)}
                    </span>
                  ))
                ) : (
                  <span>No deductions entered.</span>
                )}
              </div>
              <div className="toolbar">
                <input
                  type="search"
                  value={deductionSearch}
                  onChange={(event) => setDeductionSearch(event.target.value)}
                  placeholder="Search deductions by employee, period, reason, reference..."
                />
              </div>
              <DeductionsTable
                deductions={filteredDeductions}
                employees={employees}
                currency={settings.currency}
                onDelete={(deductionId) =>
                  updateDeductions(deductions.filter((item) => item.id !== deductionId))
                }
              />
            </section>
          </section>
        )}

        {activePage === 'payroll' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Payroll</h3>
                  <p>Select month and year, calculate payroll, preview rows, and export Excel.</p>
                </div>
                <button type="button" onClick={exportPayroll} disabled={!payrollRows.length}>
                  Export Payroll Excel
                </button>
              </div>
              <div className="toolbar period-toolbar">
                <MonthYearControls
                  month={payrollMonth}
                  year={payrollYear}
                  onMonthChange={setPayrollMonth}
                  onYearChange={setPayrollYear}
                />
              </div>
              <PayrollTable rows={payrollRows} currency={settings.currency} />
            </section>
          </section>
        )}

        {activePage === 'wps' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>WPS / SIF Export</h3>
                  <p>Generate a UAE WPS salary information file from the selected payroll period after required employer and employee WPS fields are complete.</p>
                </div>
                <button
                  type="button"
                  onClick={exportSifFile}
                  disabled={!payrollRows.length || wpsPreview.validation.errors.length > 0}
                >
                  Export SIF File
                </button>
              </div>
              <div className="toolbar period-toolbar">
                <MonthYearControls
                  month={payrollMonth}
                  year={payrollYear}
                  onMonthChange={setPayrollMonth}
                  onYearChange={setPayrollYear}
                />
              </div>
              <ValidationPanel
                errors={wpsPreview.validation.errors}
                warnings={wpsPreview.validation.warnings}
              />
              <div className="summary-strip">
                <span>Employees: {wpsPreview.rows.length}</span>
                <span>
                  Total WPS Pay:{' '}
                  {formatMoney(
                    wpsPreview.rows.reduce((sum, row) => sum + row.totalPay, 0),
                    settings.currency,
                  )}
                </span>
                <span>File: {wpsPreview.filename}</span>
              </div>
              <WpsPreviewTable rows={wpsPreview.rows} currency={settings.currency} />
            </section>

            <section className="panel notice-panel">
              <h3>WPS Note</h3>
              <p>
                The SIF export uses EDR detail rows and an SCR control row. Final
                acceptance depends on the company's bank or exchange-house WPS
                agent settings, so employer ID, routing codes, employee person IDs,
                and account identifiers must be checked against that portal before
                submission.
              </p>
            </section>
          </section>
        )}

        {activePage === 'slips' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Salary Slips</h3>
                  <p>Select month, year, and employee to export salary slip PDF files.</p>
                </div>
                <button
                  type="button"
                  className="secondary"
                  onClick={exportAllSlipPdfs}
                  disabled={!slipRows.length}
                >
                  Generate All Salary Slips
                </button>
              </div>
              <div className="toolbar period-toolbar">
                <MonthYearControls
                  month={slipMonth}
                  year={slipYear}
                  onMonthChange={setSlipMonth}
                  onYearChange={setSlipYear}
                />
              </div>
              <div className="slip-layout">
                <label>
                  Employee
                  <select
                    value={selectedSlip?.id ?? ''}
                    onChange={(event) => setSelectedSlipEmployeeId(event.target.value)}
                  >
                    {slipRows.map((row) => (
                      <option key={row.id} value={row.id}>
                        {row.employeeCode} - {row.employeeName}
                      </option>
                    ))}
                  </select>
                </label>
                {selectedSlip ? (
                  <article className="salary-slip">
                    <div>
                      <span>{settings.companyName}</span>
                      <strong>{monthName(selectedSlip.month)} {selectedSlip.year}</strong>
                    </div>
                    <dl>
                      <div>
                        <dt>Employee Code</dt>
                        <dd>{selectedSlip.employeeCode}</dd>
                      </div>
                      <div>
                        <dt>Employee Name</dt>
                        <dd>{selectedSlip.employeeName}</dd>
                      </div>
                      <div>
                        <dt>Role / Designation</dt>
                        <dd>{selectedSlip.roleDesignation}</dd>
                      </div>
                      <div>
                        <dt>Net Salary</dt>
                        <dd>{formatMoney(selectedSlip.netSalary, settings.currency)}</dd>
                      </div>
                    </dl>
                    <button type="button" onClick={() => exportSlipPdf(selectedSlip)}>
                      Export Salary Slip PDF
                    </button>
                  </article>
                ) : (
                  <div className="empty-state">Add employees to generate salary slips.</div>
                )}
              </div>
            </section>
          </section>
        )}

        {activePage === 'backup' && (
          <section className="page-stack">
            <section className="panel">
              <div className="panel-heading">
                <div>
                  <h3>Backup / Restore</h3>
                  <p>Export local backup JSON, import backup JSON, or clear local data with confirmation.</p>
                </div>
              </div>
              <div className="action-grid">
                <button
                  type="button"
                  onClick={() =>
                    downloadJson('logicore-payroll-backup.json', {
                      employees,
                      financeRecords,
                      assetRecords,
                      documentRecords,
                      payrollInputs,
                      deductions,
                      settings,
                    })
                  }
                >
                  Export Local Backup JSON
                </button>
                <label className="file-button">
                  Import Backup JSON
                  <input type="file" accept=".json" onChange={restoreBackup} />
                </label>
                <button type="button" className="danger" onClick={clearLocalData}>
                  Clear Local Data
                </button>
              </div>
            </section>
          </section>
        )}

        {activePage === 'settings' && (
          <section className="page-stack">
            <section className="panel">
              <h3>Settings</h3>
              <div className="form-grid settings-form">
                <label>
                  Company Name
                  <input
                    value={settings.companyName}
                    onChange={(event) =>
                      updateSettings({ ...settings, companyName: event.target.value })
                    }
                  />
                </label>
                <label>
                  Company Address
                  <input
                    value={settings.companyAddress}
                    onChange={(event) =>
                      updateSettings({ ...settings, companyAddress: event.target.value })
                    }
                  />
                </label>
                <label>
                  Company Phone
                  <input
                    value={settings.companyPhone}
                    onChange={(event) =>
                      updateSettings({ ...settings, companyPhone: event.target.value })
                    }
                  />
                </label>
                <label>
                  Company Email
                  <input
                    type="email"
                    value={settings.companyEmail}
                    onChange={(event) =>
                      updateSettings({ ...settings, companyEmail: event.target.value })
                    }
                  />
                </label>
                <label>
                  WPS Employer ID
                  <input
                    value={settings.wpsEmployerId}
                    onChange={(event) =>
                      updateSettings({
                        ...settings,
                        wpsEmployerId: normalizeWpsId(event.target.value, 13),
                      })
                    }
                  />
                </label>
                <label>
                  WPS Employer Routing Code
                  <input
                    value={settings.wpsRoutingCode}
                    onChange={(event) =>
                      updateSettings({
                        ...settings,
                        wpsRoutingCode: normalizeWpsRoutingCode(event.target.value),
                      })
                    }
                  />
                </label>
                <label>
                  WPS Employer Reference
                  <input
                    value={settings.wpsEmployerReference}
                    onChange={(event) =>
                      updateSettings({
                        ...settings,
                        wpsEmployerReference: event.target.value,
                      })
                    }
                  />
                </label>
                <MonthYearControls
                  month={settings.defaultPayrollMonth}
                  year={settings.defaultPayrollYear}
                  onMonthChange={(month) => {
                    updateSettings({ ...settings, defaultPayrollMonth: month })
                    setPayrollMonth(month)
                    setSlipMonth(month)
                    setDraftPayrollInput({ ...draftPayrollInput, month })
                    setDraftDeduction({ ...draftDeduction, month })
                  }}
                  onYearChange={(year) => {
                    updateSettings({ ...settings, defaultPayrollYear: year })
                    setPayrollYear(year)
                    setSlipYear(year)
                    setDraftPayrollInput({ ...draftPayrollInput, year })
                    setDraftDeduction({ ...draftDeduction, year })
                  }}
                />
                <label>
                  Currency
                  <input
                    value={settings.currency}
                    onChange={(event) =>
                      updateSettings({
                        ...settings,
                        currency: event.target.value.toUpperCase(),
                      })
                    }
                  />
                </label>
              </div>
            </section>
            <section className="panel notice-panel">
              <h3>Local Data Notice</h3>
              <p>
                Logicore Excel Manager stores data locally on this device.
                No employee or payroll data is sent to a hosted service by this
                desktop app.
              </p>
              <p>
                WPS/SIF output should be checked against the employer's bank or
                exchange-house portal before submission.
              </p>
            </section>
          </section>
        )}

        {activePage === 'about' && (
          <section className="page-stack">
            <section className="hero-panel about-hero">
              <div>
                <p className="eyebrow">Notes</p>
                <h3>A local payroll and payslip generator for small company HR teams.</h3>
                <p>
                  The app helps organize employees, monthly payroll inputs,
                  deductions, payroll exports, WPS/SIF files, and salary slip PDFs
                  without requiring a cloud account.
                </p>
              </div>
            </section>
            <section className="panel about-grid">
              <article>
                <h3>What this product is</h3>
                <p>
                  A local employee and payroll workflow for small employers that
                  need Excel import, review, calculation, and export.
                </p>
              </article>
              <article>
                <h3>Compliance stance</h3>
                <p>
                  This app generates payroll records and a common WPS/SIF structure,
                  but it is not legal advice and does not guarantee bank acceptance.
                  The employer remains responsible for final submission checks.
                </p>
              </article>
              <article>
                <h3>Local-only data stance</h3>
                <p>
                  This version uses browser/Electron local storage and local JSON
                  backup. Data remains on the user's device unless the user exports or
                  shares files.
                </p>
              </article>
              <article>
                <h3>Architecture stance</h3>
                <p>
                  PostgreSQL, API, Docker, and hosted demo layers are intentionally
                  not enabled in this local-first release because they would change
                  the privacy and operating model. The Electron package is the
                  recommended way to operate locally.
                </p>
              </article>
            </section>
          </section>
        )}
      </main>
    </div>
  )
}

function SummaryCard({ label, value }: { label: string; value: string | number }) {
  return (
    <article className="summary-card">
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  )
}

function NumberField({
  label,
  value,
  onChange,
}: {
  label: string
  value: number
  onChange: (value: number) => void
}) {
  return (
    <label>
      {label}
      <input
        type="number"
        min="0"
        step="0.01"
        value={value}
        onChange={(event) => onChange(numberValue(event.target.value))}
      />
    </label>
  )
}

function MonthYearControls({
  month,
  year,
  onMonthChange,
  onYearChange,
}: {
  month: string
  year: string
  onMonthChange: (month: string) => void
  onYearChange: (year: string) => void
}) {
  return (
    <>
      <label>
        Month
        <select value={month} onChange={(event) => onMonthChange(event.target.value)}>
          {monthOptions.map((option) => (
            <option key={option.value} value={option.value}>
              {option.label}
            </option>
          ))}
        </select>
      </label>
      <label>
        Year
        <input
          type="number"
          min="2000"
          max="2100"
          value={year}
          onChange={(event) => onYearChange(event.target.value)}
        />
      </label>
    </>
  )
}

function ValidationPanel({
  errors,
  warnings,
}: {
  errors: string[]
  warnings: string[]
}) {
  if (!errors.length && !warnings.length) return null

  return (
    <div className="validation-grid">
      {errors.length > 0 && (
        <div className="notice error">
          <strong>Validation errors</strong>
          <ul>
            {errors.slice(0, 10).map((error) => (
              <li key={error}>{error}</li>
            ))}
          </ul>
        </div>
      )}
      {warnings.length > 0 && (
        <div className="notice warning">
          <strong>Validation warnings</strong>
          <ul>
            {warnings.slice(0, 10).map((warning) => (
              <li key={warning}>{warning}</li>
            ))}
          </ul>
        </div>
      )}
    </div>
  )
}

function EditableEmployeeTable({
  employees,
  onCellChange,
  onDelete,
}: {
  employees: Employee[]
  onCellChange: (employeeId: string, key: keyof Employee, value: string) => void
  onDelete: (employeeId: string) => void
}) {
  if (!employees.length) return <div className="empty-state">No employees found.</div>

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            {employeeFields.map((field) => (
              <th key={field.key}>{field.label}</th>
            ))}
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {employees.map((employee) => (
            <tr key={employee.id}>
              {employeeFields.map((field) => (
                <td key={field.key}>
                  {field.type === 'select' ? (
                    <select
                      value={String(employee[field.key])}
                      onChange={(event) =>
                        onCellChange(employee.id, field.key, event.target.value)
                      }
                    >
                      {selectOptions(field.key).map((option) => (
                        <option key={option} value={option}>
                          {option}
                        </option>
                      ))}
                    </select>
                  ) : (
                    <input
                      type={field.type}
                      value={String(employee[field.key])}
                      onChange={(event) =>
                        onCellChange(employee.id, field.key, event.target.value)
                      }
                    />
                  )}
                </td>
              ))}
              <td>
                <button
                  type="button"
                  className="link-danger"
                  onClick={() => onDelete(employee.id)}
                >
                  Delete
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function PayrollInputsTable({
  inputs,
  employees,
  currency,
  onDelete,
}: {
  inputs: PayrollInput[]
  employees: Employee[]
  currency: string
  onDelete: (inputId: string) => void
}) {
  if (!inputs.length) return <div className="empty-state">No payroll input rows found.</div>

  const employeeLabel = (employeeId: string) => {
    const employee = employees.find((item) => item.id === employeeId)
    return employee ? `${employee.employeeCode} - ${employee.employeeName}` : 'Unknown'
  }

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Employee</th>
            <th>Month</th>
            <th>Year</th>
            <th>Days Present</th>
            <th>Working Days</th>
            <th>Allowance</th>
            <th>Incentive</th>
            <th>Attendance Bonus</th>
            <th>Performance Bonus / Tip</th>
            <th>Overtime Pay</th>
            <th>Other Earnings</th>
            <th>Arrears</th>
            <th>Remarks</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {inputs.map((input) => (
            <tr key={input.id}>
              <td>{employeeLabel(input.employeeId)}</td>
              <td>{monthName(input.month)}</td>
              <td>{input.year}</td>
              <td>{input.daysPresent}</td>
              <td>{input.workingDays}</td>
              <td>{formatMoney(input.allowance, currency)}</td>
              <td>{formatMoney(input.incentive, currency)}</td>
              <td>{formatMoney(input.attendanceBonus, currency)}</td>
              <td>{formatMoney(input.performanceBonus, currency)}</td>
              <td>{formatMoney(input.overtimePay, currency)}</td>
              <td>{formatMoney(input.otherEarning, currency)}</td>
              <td>{formatMoney(input.arrears, currency)}</td>
              <td>{input.remarks}</td>
              <td>
                <button
                  type="button"
                  className="link-danger"
                  onClick={() => onDelete(input.id)}
                >
                  Delete
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function FinanceTable({
  records,
  currency,
  onDelete,
}: {
  records: FinanceRecord[]
  currency: string
  onDelete: (recordId: string) => void
}) {
  if (!records.length) return <div className="empty-state">No finance rows found.</div>

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Date</th>
            <th>Type</th>
            <th>Category</th>
            <th>Party</th>
            <th>Reference No</th>
            <th>Description</th>
            <th>Amount</th>
            <th>VAT Amount</th>
            <th>Payment Mode</th>
            <th>Status</th>
            <th>Attachment</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {records.map((record) => (
            <tr key={record.id}>
              <td>{record.date}</td>
              <td>{record.type}</td>
              <td>{record.category}</td>
              <td>{record.party}</td>
              <td>{record.referenceNo}</td>
              <td>{record.description}</td>
              <td>{formatMoney(record.amount, currency)}</td>
              <td>{formatMoney(record.vatAmount, currency)}</td>
              <td>{record.paymentMode}</td>
              <td>{record.status}</td>
              <td>{record.attachmentName}</td>
              <td>
                <button
                  type="button"
                  className="link-danger"
                  onClick={() => onDelete(record.id)}
                >
                  Delete
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function AssetTable({
  records,
  employees,
  currency,
  onDelete,
}: {
  records: AssetRecord[]
  employees: Employee[]
  currency: string
  onDelete: (recordId: string) => void
}) {
  if (!records.length) return <div className="empty-state">No asset rows found.</div>

  const employeeLabel = (employeeId: string) => {
    const employee = employees.find((item) => item.id === employeeId)
    return employee ? `${employee.employeeCode} - ${employee.employeeName}` : ''
  }

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Asset Code</th>
            <th>Asset Name</th>
            <th>Category</th>
            <th>Assigned To</th>
            <th>Purchase Date</th>
            <th>Purchase Cost</th>
            <th>Serial Number</th>
            <th>Location</th>
            <th>Condition</th>
            <th>Status</th>
            <th>Warranty Expiry</th>
            <th>Notes</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {records.map((record) => (
            <tr key={record.id}>
              <td>{record.assetCode}</td>
              <td>{record.assetName}</td>
              <td>{record.category}</td>
              <td>{employeeLabel(record.assignedToEmployeeId)}</td>
              <td>{record.purchaseDate}</td>
              <td>{formatMoney(record.purchaseCost, currency)}</td>
              <td>{record.serialNumber}</td>
              <td>{record.location}</td>
              <td>{record.condition}</td>
              <td>{record.status}</td>
              <td>{record.warrantyExpiry}</td>
              <td>{record.notes}</td>
              <td>
                <button
                  type="button"
                  className="link-danger"
                  onClick={() => onDelete(record.id)}
                >
                  Delete
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function DocumentTable({
  records,
  onDelete,
}: {
  records: DocumentRecord[]
  onDelete: (recordId: string) => void
}) {
  if (!records.length) return <div className="empty-state">No document rows found.</div>

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Document Date</th>
            <th>Document Type</th>
            <th>Linked Module</th>
            <th>Linked Reference</th>
            <th>Title</th>
            <th>File Name</th>
            <th>File Size KB</th>
            <th>Notes</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {records.map((record) => (
            <tr key={record.id}>
              <td>{record.documentDate}</td>
              <td>{record.documentType}</td>
              <td>{record.linkedModule}</td>
              <td>{record.linkedReference}</td>
              <td>{record.title}</td>
              <td>{record.fileName}</td>
              <td>{record.fileSizeKb}</td>
              <td>{record.notes}</td>
              <td>
                <button
                  type="button"
                  className="link-danger"
                  onClick={() => onDelete(record.id)}
                >
                  Delete
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function DeductionsTable({
  deductions,
  employees,
  currency,
  onDelete,
}: {
  deductions: Deduction[]
  employees: Employee[]
  currency: string
  onDelete: (deductionId: string) => void
}) {
  if (!deductions.length) return <div className="empty-state">No deduction rows found.</div>

  const employeeLabel = (employeeId: string) => {
    const employee = employees.find((item) => item.id === employeeId)
    return employee ? `${employee.employeeCode} - ${employee.employeeName}` : 'Unknown'
  }

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Employee</th>
            <th>Month</th>
            <th>Year</th>
            <th>Transaction ID</th>
            <th>Advance Deduction</th>
            <th>Fine Deduction</th>
            <th>Fine Date</th>
            <th>Fine Receipt No</th>
            <th>Salik Deduction</th>
            <th>Salik Date</th>
            <th>Salik Receipt No</th>
            <th>Other Deduction</th>
            <th>Reason / Remarks</th>
            <th>Reference No</th>
            <th>Attachment</th>
            <th>Action</th>
          </tr>
        </thead>
        <tbody>
          {deductions.map((deduction) => (
            <tr key={deduction.id}>
              <td>{employeeLabel(deduction.employeeId)}</td>
              <td>{monthName(deduction.month)}</td>
              <td>{deduction.year}</td>
              <td>{deduction.transactionId}</td>
              <td>{formatMoney(deduction.advanceDeduction, currency)}</td>
              <td>{formatMoney(deduction.fineDeduction, currency)}</td>
              <td>{deduction.fineDate}</td>
              <td>{deduction.fineReceiptNo || deduction.fineTransactionNo}</td>
              <td>{formatMoney(deduction.salikDeduction, currency)}</td>
              <td>{deduction.salikDate}</td>
              <td>{deduction.salikReceiptNo || deduction.salikTransactionNo}</td>
              <td>{formatMoney(deduction.otherDeduction, currency)}</td>
              <td>{deduction.reasonRemarks}</td>
              <td>{deduction.referenceNo}</td>
              <td>{deduction.attachmentName}</td>
              <td>
                <button
                  type="button"
                  className="link-danger"
                  onClick={() => onDelete(deduction.id)}
                >
                  Delete
                </button>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function WpsPreviewTable({
  rows,
  currency,
}: {
  rows: SifRecordSet['rows']
  currency: string
}) {
  if (!rows.length) return <div className="empty-state">No WPS rows available.</div>

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Employee Code</th>
            <th>Employee Name</th>
            <th>Fixed Income</th>
            <th>Variable Income</th>
            <th>Leave Days</th>
            <th>Total Pay</th>
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={`${row.employeeCode}-${row.employeeName}`}>
              <td>{row.employeeCode}</td>
              <td>{row.employeeName}</td>
              <td>{formatMoney(row.fixedIncome, currency)}</td>
              <td>{formatMoney(row.variableIncome, currency)}</td>
              <td>{row.leaveDays}</td>
              <td>
                <strong>{formatMoney(row.totalPay, currency)}</strong>
              </td>
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

function PayrollTable({
  rows,
  currency,
  compact = false,
}: {
  rows: PayrollRow[]
  currency: string
  compact?: boolean
}) {
  if (!rows.length) return <div className="empty-state">No payroll rows available.</div>

  return (
    <div className="table-wrap">
      <table>
        <thead>
          <tr>
            <th>Employee Code</th>
            {!compact && <th>Payroll Code</th>}
            {!compact && <th>Emirates ID</th>}
            <th>Employee Name</th>
            <th>Role / Designation</th>
            {!compact && <th>Date of Joining</th>}
            {!compact && <th>Salary Type</th>}
            <th>Basic Salary</th>
            {!compact && <th>Accommodation</th>}
            {!compact && <th>Travel</th>}
            <th>Allowance</th>
            {!compact && <th>Incentive</th>}
            {!compact && <th>Attendance Bonus</th>}
            {!compact && <th>Performance Bonus / Tip</th>}
            {!compact && <th>O/T</th>}
            {!compact && <th>Other Earnings</th>}
            {!compact && <th>Salary Earned</th>}
            <th>Gross Salary</th>
            {!compact && <th>Arrears</th>}
            {!compact && <th>Payable Before Deductions</th>}
            <th>Advance Deduction</th>
            {!compact && <th>Fine Deduction</th>}
            {!compact && <th>Salik Deduction</th>}
            <th>Other Deduction</th>
            <th>Total Deduction</th>
            <th>Net Salary</th>
            {!compact && <th>Days Present</th>}
            {!compact && <th>Working Days</th>}
            {!compact && <th>Status</th>}
          </tr>
        </thead>
        <tbody>
          {rows.map((row) => (
            <tr key={row.id}>
              <td>{row.employeeCode}</td>
              {!compact && <td>{row.payrollCode || row.employeeCode}</td>}
              {!compact && <td>{row.emiratesId}</td>}
              <td>{row.employeeName}</td>
              <td>{row.roleDesignation}</td>
              {!compact && <td>{row.dateOfJoining}</td>}
              {!compact && <td>{row.salaryType}</td>}
              <td>{formatMoney(row.basicSalary, currency)}</td>
              {!compact && <td>{formatMoney(row.accommodationAllowance, currency)}</td>}
              {!compact && <td>{formatMoney(row.travelAllowance, currency)}</td>}
              <td>{formatMoney(row.allowance, currency)}</td>
              {!compact && <td>{formatMoney(row.incentive, currency)}</td>}
              {!compact && <td>{formatMoney(row.attendanceBonus, currency)}</td>}
              {!compact && <td>{formatMoney(row.performanceBonus, currency)}</td>}
              {!compact && <td>{formatMoney(row.overtimePay, currency)}</td>}
              {!compact && <td>{formatMoney(row.otherEarning, currency)}</td>}
              {!compact && <td>{formatMoney(row.salaryEarned, currency)}</td>}
              <td>{formatMoney(row.grossSalary, currency)}</td>
              {!compact && <td>{formatMoney(row.dueFromLastMonth, currency)}</td>}
              {!compact && <td>{formatMoney(row.payableBeforeDeductions, currency)}</td>}
              <td>{formatMoney(row.advanceDeduction, currency)}</td>
              {!compact && <td>{formatMoney(row.fineDeduction, currency)}</td>}
              {!compact && <td>{formatMoney(row.salikDeduction, currency)}</td>}
              <td>{formatMoney(row.otherDeduction, currency)}</td>
              <td>{formatMoney(row.totalDeduction, currency)}</td>
              <td>
                <strong>{formatMoney(row.netSalary, currency)}</strong>
              </td>
              {!compact && <td>{row.daysPresent}</td>}
              {!compact && <td>{row.workingDays}</td>}
              {!compact && <td>{row.status}</td>}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
}

export default App
