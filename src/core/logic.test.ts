import ExcelJS from 'exceljs'
import { describe, expect, it } from 'vitest'
import {
  blankDeduction,
  blankEmployee,
  blankPayrollInput,
  buildPayrollRows,
  buildSifRecords,
  calculateWpsPayment,
  defaultSettings,
  employeeTemplateColumns,
  findMissingColumns,
  readFirstWorksheet,
  resolveEmployeeForDeductionRow,
  validateDeduction,
  validateEmployee,
  validateExcelFile,
  validatePayrollInput,
  validateWpsPayroll,
  worksheetToJsonRows,
} from './logic'
import type { Deduction, Employee, PayrollInput, Settings } from './logic'

function employee(overrides: Partial<Employee> = {}): Employee {
  return {
    ...blankEmployee(),
    id: 'emp-1',
    employeeCode: 'E001',
    payrollCode: 'P001',
    wpsPersonId: '00000000012345',
    wpsRoutingCode: '123456789',
    wpsAccountNumber: 'AE070331234567890123456',
    employeeName: 'Amina Driver',
    roleDesignation: 'Driver',
    dateOfJoining: '2026-01-01',
    mobileNumber: '0501234567',
    basicSalary: 3000,
    accommodationAllowance: 500,
    travelAllowance: 300,
    workingDays: 30,
    daysPresent: 30,
    ...overrides,
  }
}

function deduction(overrides: Partial<Deduction> = {}): Deduction {
  return {
    ...blankDeduction('06', '2026', 'emp-1'),
    id: 'ded-1',
    ...overrides,
  }
}

function payrollInput(overrides: Partial<PayrollInput> = {}): PayrollInput {
  return {
    ...blankPayrollInput('06', '2026', 'emp-1'),
    id: 'input-1',
    ...overrides,
  }
}

function wpsSettings(overrides: Partial<Settings> = {}): Settings {
  return {
    ...defaultSettings,
    companyName: 'Logicore Test Tenant',
    wpsEmployerId: '1234567890123',
    wpsRoutingCode: '123456789',
    wpsEmployerReference: 'TEST-RUN',
    ...overrides,
  }
}

function fileLike(name: string, size: number): File {
  return { name, size } as File
}

describe('payroll calculation rules', () => {
  it('calculates fixed-salary gross, arrears, deductions, and net salary', () => {
    const rows = buildPayrollRows(
      [employee()],
      [
        payrollInput({
          daysPresent: 28,
          workingDays: 30,
          allowance: 250,
          incentive: 100,
          attendanceBonus: 40,
          performanceBonus: 60,
          overtimePay: 75,
          otherEarning: 20,
          arrears: 80,
        }),
      ],
      [
        deduction({
          advanceDeduction: 100,
          fineDeduction: 50,
          salikDeduction: 25,
          otherDeduction: 10,
        }),
      ],
      '06',
      '2026',
    )

    expect(rows).toHaveLength(1)
    expect(rows[0]).toMatchObject({
      salaryEarned: 4050,
      grossSalary: 4345,
      payableBeforeDeductions: 4425,
      totalDeduction: 185,
      netSalary: 4240,
      dueFromLastMonth: 80,
    })
  })

  it('prorates per-day salary by eligible days', () => {
    const rows = buildPayrollRows(
      [
        employee({
          salaryType: 'Per Day',
          basicSalary: 3000,
          accommodationAllowance: 0,
          travelAllowance: 0,
          allowance: 0,
          daysPresent: 15,
          workingDays: 30,
        }),
      ],
      [],
      [],
      '06',
      '2026',
    )

    expect(rows[0].salaryEarned).toBe(1500)
    expect(rows[0].grossSalary).toBe(1500)
    expect(rows[0].netSalary).toBe(1500)
  })
})

describe('WPS/SIF validation rules', () => {
  it('splits WPS fixed and variable income from a calculated payroll row', () => {
    const [row] = buildPayrollRows(
      [employee()],
      [payrollInput({ incentive: 100, performanceBonus: 50 })],
      [],
      '06',
      '2026',
    )

    expect(calculateWpsPayment(row)).toMatchObject({
      fixedIncome: 3800,
      variableIncome: 150,
      leaveDays: 0,
      totalPay: 3950,
    })
  })

  it('builds SIF details and summary when WPS settings and employee data are valid', () => {
    const rows = buildPayrollRows([employee()], [], [], '06', '2026')
    const result = buildSifRecords(rows, wpsSettings(), '06', '2026')

    expect(result.validation.errors).toEqual([])
    expect(result.rows).toEqual([
      {
        employeeCode: 'E001',
        employeeName: 'Amina Driver',
        fixedIncome: 3800,
        variableIncome: 0,
        leaveDays: 0,
        totalPay: 3800,
      },
    ])
    expect(result.content).toContain('EDR,00000000012345,123456789')
    expect(result.content).toContain('SCR,1234567890123,123456789')
    expect(result.filename).toMatch(/^wps-1234567890123-\d{12}\.SIF$/)
  })

  it('returns actionable errors for missing WPS tenant and employee fields', () => {
    const rows = buildPayrollRows(
      [
        employee({
          wpsPersonId: '',
          wpsRoutingCode: '',
          wpsAccountNumber: '',
        }),
      ],
      [],
      [deduction({ advanceDeduction: 5000 })],
      '06',
      '2026',
    )
    const result = validateWpsPayroll(rows, { ...defaultSettings, wpsEmployerId: '', wpsRoutingCode: '' })

    expect(result.errors).toEqual(
      expect.arrayContaining([
        'WPS Employer ID is required in Settings.',
        'WPS employer agent routing code is required in Settings.',
        'E001 - Amina Driver: WPS Person ID is required.',
        'E001 - Amina Driver: WPS agent routing code is required.',
        'E001 - Amina Driver: WPS IBAN / account number is required.',
        'E001 - Amina Driver: net salary must be greater than zero for WPS export.',
        'E001 - Amina Driver: net salary cannot be negative.',
      ]),
    )
  })
})

describe('Excel import validation rules', () => {
  it('rejects unsupported, empty, and oversized Excel files before parsing', () => {
    expect(validateExcelFile(fileLike('employees.xls', 100))).toBe(
      'Only .xlsx files are supported.',
    )
    expect(validateExcelFile(fileLike('employees.xlsx', 0))).toBe(
      'The selected Excel file is empty.',
    )
    expect(validateExcelFile(fileLike('employees.xlsx', 10 * 1024 * 1024 + 1))).toBe(
      'The selected Excel file is too large. Keep it under 10 MB.',
    )
    expect(validateExcelFile(fileLike('employees.xlsx', 1024))).toBeNull()
  })

  it('detects missing template columns with case-insensitive header matching', () => {
    const headers = employeeTemplateColumns.filter(
      (column) => column !== 'Employee Name' && column !== 'Basic Salary',
    )
    headers[0] = ' employee code '

    expect(findMissingColumns(headers, employeeTemplateColumns)).toEqual([
      'Employee Name',
      'Basic Salary',
    ])
  })

  it('converts the first worksheet into JSON rows and ignores empty data rows', () => {
    const workbook = new ExcelJS.Workbook()
    const worksheet = workbook.addWorksheet('Employees')
    worksheet.addRow(['Employee Code', 'Employee Name', 'Basic Salary'])
    worksheet.addRow(['E001', 'Amina Driver', 3000])
    worksheet.addRow([])

    expect(readFirstWorksheet(workbook).name).toBe('Employees')
    expect(worksheetToJsonRows(worksheet)).toEqual([
      {
        'Employee Code': 'E001',
        'Employee Name': 'Amina Driver',
        'Basic Salary': 3000,
      },
    ])
  })

  it('validates employee import rows for required fields and duplicate keys', () => {
    const existing = employee({ id: 'existing', employeeCode: 'E001' })
    const imported = employee({
      id: 'imported',
      employeeCode: 'E001',
      employeeName: '',
      basicSalary: 0,
    })

    const result = validateEmployee(imported, [existing])

    expect(result.errors).toEqual(
      expect.arrayContaining([
        'Employee Name is required.',
        'Basic Salary is required and must be numeric.',
        'Employee Code must be unique.',
      ]),
    )
  })

  it('validates payroll input duplicates and negative numeric values', () => {
    const existing = payrollInput({ id: 'existing', employeeId: 'emp-1' })
    const imported = payrollInput({
      id: 'imported',
      employeeId: 'emp-1',
      overtimePay: -1,
    })

    const result = validatePayrollInput(imported, [employee()], [existing])

    expect(result.errors).toEqual(
      expect.arrayContaining([
        'Only one payroll input row is allowed per employee/month/year.',
        'Overtime Pay cannot be negative.',
      ]),
    )
  })

  it('validates deduction import rows and employee matching failures', () => {
    const resolved = resolveEmployeeForDeductionRow({ 'Agreement No': 'P001' }, [employee()])
    const missing = resolveEmployeeForDeductionRow({ 'Employee Code': 'NOPE' }, [employee()])
    const result = validateDeduction(
      deduction({ employeeId: 'missing', fineDeduction: -5 }),
      [employee()],
      [],
    )

    expect(resolved.employee?.id).toBe('emp-1')
    expect(resolved.matchedBy).toBe('Agreement No')
    expect(missing.employee).toBeUndefined()
    expect(missing.lookupValue).toBe('NOPE')
    expect(result.errors).toEqual(
      expect.arrayContaining([
        'Selected employee was not found.',
        'Fine Deduction cannot be negative.',
      ]),
    )
  })
})
