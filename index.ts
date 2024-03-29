import * as ExcelJS from 'exceljs';

// provide the path to the Excel file
const file_path = 'Assignment_Timecard.xlsx';

// create a new Excel workbook
const workbook = new ExcelJS.Workbook();

// Set to keep track of printed names
const printedNames = new Set<string>();

/**
 * convert Timecard string to Hours
 *
 * @param timecard A string in the format HH:MM
 * @returns
 */
function convertTimecardToHours(timecard: string): number {
    const [hours, minutes] = timecard.split(':').map(Number);
    return hours + minutes / 60;
}

/**
 * Read the Excel file and process the data
 * make objects of Employee class and store them in a dictionary
 */
workbook.xlsx
    .readFile(file_path)
    .then(() => {
        const worksheet = workbook.getWorksheet(1);

        const employees: { [key: string]: Employee } = {};

        worksheet?.eachRow({ includeEmpty: false }, (row, rowNumber) => {
            const employeeName = row.getCell(8).text as string;

            if (!employees[employeeName]) {
                employees[employeeName] = new Employee(employeeName);
            }

            const startTime = row.getCell(3).value as Date;
            const endTime = row.getCell(4).value as Date;

            if (!(startTime instanceof Date) || !(endTime instanceof Date)) {
                console.warn(`Invalid date values in row ${rowNumber}`);
                return;
            }

            const timecardValue = row.getCell(5).text as string;
            const timecardHours = convertTimecardToHours(timecardValue);

            employees[employeeName].addShift(startTime, endTime, timecardHours);
        });

        // Employees who have worked for 7 consecutive days
        console.log('Employees who have worked for 7 consecutive days:');
        printedNames.clear();
        for (const employeeId in employees) {
            if (
                employees[employeeId].hasConsecutiveDays(7) &&
                !printedNames.has(employeeId)
            ) {
                printedNames.add(employeeId);
                console.log(`Employee name: ${employeeId}`);
            }
        }

        // Employees with less than 10 hours between shifts but greater than 1 hour
        console.log(
            '\nEmployees with less than 10 hours between shifts but greater than 1 hour:',
        );
        printedNames.clear();
        for (const employeeId in employees) {
            if (
                employees[employeeId].hasShortBreaks(1, 10) &&
                !printedNames.has(employeeId)
            ) {
                printedNames.add(employeeId);
                console.log(`Employee name: ${employeeId}`);
            }
        }

        // Employees who have worked for more than 14 hours in a single shift
        console.log(
            '\nEmployees who have worked for more than 14 hours in a single shift:',
        );
        printedNames.clear();
        for (const employeeId in employees) {
            if (
                employees[employeeId].hasLongShifts(14) &&
                !printedNames.has(employeeId)
            ) {
                printedNames.add(employeeId);
                console.log(`Employee name: ${employeeId}`);
            }
        }
    })
    .catch((error) => {
        console.error('Error reading Excel file:', error);
    });

// class to represent an Employee
class Employee {
    public name: string;
    private shifts: { startTime: Date; endTime: Date; timecardHours: number }[];

    constructor(name: string) {
        this.name = name;
        this.shifts = [];
    }

    public addShift(startTime: Date, endTime: Date, timecard: number): void {
        this.shifts.push({ startTime, endTime, timecardHours: timecard });
    }

    /**
     * method to find if the employee has worked for consecutive days
     *
     * @param days Number of consecutive days
     * @returns
     */
    public hasConsecutiveDays(days: number): boolean {
        if (this.shifts.length < days) {
            return false;
        }

        for (let i = 0; i < this.shifts.length - days + 1; i++) {
            const consecutiveShifts = this.shifts.slice(i, i + days);
            const consecutiveDays = consecutiveShifts.map(
                (shift) => shift.startTime,
            );

            if (new Set(consecutiveDays).size === days) {
                return true;
            }
        }

        return false;
    }

    /**
     * method to find if the employee has short breaks between shifts
     *
     * assumption - Time is the time of entering the office, endTime is the time of leaving the office
     *              and timecardHours is the time the employee has worked
     *              So, the break duration is calculated as (endTime - startTime) - timecardHours
     *
     * @param minimumBreak
     * @param maximumBreak
     * @returns
     */
    public hasShortBreaks(minimumBreak: number, maximumBreak: number): boolean {
        for (let i = 0; i < this.shifts.length - 1; i++) {
            const totalDuration =
                (this.shifts[i + 1].endTime.getTime() -
                    this.shifts[i].startTime.getTime()) /
                (1000 * 60 * 60);
            const breakDuration = totalDuration - this.shifts[i].timecardHours;
            if (breakDuration > minimumBreak && breakDuration < maximumBreak) {
                return true;
            }
        }

        return false;
    }

    /**
     * method to find if the employee has worked for long shifts
     *
     * @param maximumHours
     * @returns
     */
    public hasLongShifts(maximumHours: number): boolean {
        for (const shift of this.shifts) {
            const shiftDuration =
                (shift.endTime.getTime() - shift.startTime.getTime()) /
                (1000 * 60 * 60);

            if (shiftDuration > maximumHours) {
                return true;
            }
        }

        return false;
    }
}
