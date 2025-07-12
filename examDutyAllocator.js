/**
 * Basic Exam Duty Allocation System
 * 
 * This script allocates teachers to exam rooms ensuring that each room has at least one lady teacher
 * depending on the number of students in the room.
 */

// Sample data for teachers
const teachers = [
    { id: 1, name: "Alice", gender: "female" },
    { id: 2, name: "Bob", gender: "male" },
    { id: 3, name: "Carol", gender: "female" },
    { id: 4, name: "David", gender: "male" },
    { id: 5, name: "Eve", gender: "female" },
];

// Sample data for rooms with names
const rooms = [
    { id: 101, name: "LT 401", students: 30 },
    { id: 102, name: "LT 402", students: 120 },
    { id: 103, name: "LT 403", students: 50 },
];

// Allocation function with examType parameter
function allocateDuties(rooms, teachers, examType) {
    const allocations = {};
    const shifts = examType === "mid-sem" ? 4 : examType === "end-sem" ? 2 : 1;

    // Filter lady teachers
    const ladyTeachers = teachers.filter(t => t.gender === "female");
    // Filter male teachers
    const maleTeachers = teachers.filter(t => t.gender === "male");

    // Track last shift assigned to each teacher by their id
    const lastShiftAssigned = new Map();

    // Helper function to check if teacher can be assigned to current shift
    function canAssign(teacherId, currentShift) {
        return lastShiftAssigned.get(teacherId) !== currentShift - 1;
    }

    // Initialize allocations for each shift and room
    for (let shift = 1; shift <= shifts; shift++) {
        allocations[shift] = {};
        rooms.forEach(room => {
            allocations[shift][room.id] = [];
        });
    }

    // Assign teachers shift by shift
    for (let shift = 1; shift <= shifts; shift++) {
        let ladyIndex = 0;
        let maleIndex = 0;

        rooms.forEach(room => {
            let requiredTeachers = 1;
            if (examType === "mid-sem") {
                if (room.name === "LT 402") {
                    requiredTeachers = room.students > 100 ? 4 : 3;
                } else {
                    requiredTeachers = room.students > 20 ? 2 : 1;
                }
            } else if (examType === "end-sem") {
                const agricultureRooms = ["Agriculture 101", "Agriculture 102", "Agriculture 103"];
                if (room.name === "LT 201") {
                    requiredTeachers = 2;
                } else if (agricultureRooms.includes(room.name)) {
                    requiredTeachers = 2;
                } else if (room.name === "CR 205" || room.name === "CR 604") {
                    requiredTeachers = 3;
                } else if (room.name === "New Classroom") {
                    requiredTeachers = room.students > 140 ? 4 : 2;
                } else {
                    requiredTeachers = room.students > 80 ? 3 : 2;
                }
            }

            // Allocate required number of teachers ensuring at least one lady teacher
            let allocatedLady = false;
            let allocatedCount = 0;

            while (allocatedCount < requiredTeachers) {
                let assigned = false;

                // Try to assign lady teacher first if not allocated yet
                if (!allocatedLady) {
                    while (ladyIndex < ladyTeachers.length) {
                        const teacher = ladyTeachers[ladyIndex];
                        if (canAssign(teacher.id, shift)) {
                            allocations[shift][room.id].push(teacher);
                            lastShiftAssigned.set(teacher.id, shift);
                            ladyIndex++;
                            allocatedLady = true;
                            allocatedCount++;
                            assigned = true;
                            break;
                        }
                        ladyIndex++;
                    }
                }

                if (!assigned) {
                    // Assign male teacher if lady teacher not assigned or more teachers needed
                    while (maleIndex < maleTeachers.length) {
                        const teacher = maleTeachers[maleIndex];
                        if (canAssign(teacher.id, shift)) {
                            allocations[shift][room.id].push(teacher);
                            lastShiftAssigned.set(teacher.id, shift);
                            maleIndex++;
                            allocatedCount++;
                            assigned = true;
                            break;
                        }
                        maleIndex++;
                    }
                }

                if (!assigned) {
                    // If no teacher can be assigned, push error and break
                    allocations[shift][room.id].push({ error: "Not enough teachers available" });
                    break;
                }
            }
        });
    }

    return allocations;
}

// Function to allocate duties based on duty schedule and teacher list
function allocateDutiesFromSchedule(rooms, teachers, dutySchedule) {
    const allocations = {};

    // Map teacher names to teacher objects for quick lookup
    const teacherMap = new Map();
    teachers.forEach(t => {
        teacherMap.set(t.name, t);
    });

    // Initialize allocations per shift and room
    Object.keys(dutySchedule).forEach(shift => {
        allocations[shift] = {};
        rooms.forEach(room => {
            allocations[shift][room.id] = [];
        });
    });

    // Track last shift assigned to each teacher by their id to prevent consecutive shifts
    const lastShiftAssigned = new Map();

    // Helper function to check if teacher can be assigned to current shift
    function canAssign(teacherId, currentShift) {
        return lastShiftAssigned.get(teacherId) !== currentShift - 1;
    }

    // Assign teachers to rooms per shift based on dutySchedule
    Object.entries(dutySchedule).forEach(([shift, teacherNames]) => {
        let shiftNum = parseInt(shift.replace('S', ''), 10);

        // For simplicity, assign teachers to rooms in round-robin fashion
        let roomIndex = 0;
        teacherNames.forEach(teacherName => {
            const teacher = teacherMap.get(teacherName);
            if (!teacher) {
                // Teacher not found in list, skip or log error
                return;
            }

            // Check if teacher can be assigned this shift (no consecutive shifts)
            if (!canAssign(teacher.id, shiftNum)) {
                // Skip assignment to avoid consecutive shifts
                return;
            }

            // Assign teacher to current room
            const room = rooms[roomIndex % rooms.length];
            allocations[shift][room.id].push(teacher);
            lastShiftAssigned.set(teacher.id, shiftNum);

            roomIndex++;
        });
    });

    return allocations;
}

const fs = require('fs');
const path = require('path');
const csv = require('csv-parse/lib/sync');

// Function to parse duty chart CSV and return structured data
function parseDutyChartCSV(filePath) {
    const content = fs.readFileSync(filePath, 'utf-8');
    const records = csv(content, {
        columns: true,
        skip_empty_lines: true,
        from_line: 12 // Skip header lines before actual data
    });

    // Process records to map teacher duties per shift
    const dutySchedule = {};

    records.forEach(record => {
        const teacherName = record['Name'];
        ['S1', 'S2', 'S3', 'S4'].forEach(shift => {
            if (record[shift] && record[shift].trim() === '1') {
                if (!dutySchedule[shift]) {
                    dutySchedule[shift] = [];
                }
                dutySchedule[shift].push(teacherName);
            }
        });
    });

    return dutySchedule;
}

// Function to generate duty chart CSV output
function generateDutyChartCSV(allocations, teachers, examDate) {
    // Create header information
    const header = [
        '"GRAPHIC ERA HILL UNIVERSITY, DEHRADUN"',
        '"REVISED MID SEMESTER REGULAR EXAMINATION, ' + examDate + '"',
        'Faculty Duty Chart',
        ' Invigilators must report 30 mins before the commencement of the Examination in the Exam cell and be present in the allotted room 15 mins prior to the commencement of the exam.',
        'Duty allotment and distribution of answer sheets in the Main building will be done from the Exam control room on 2nd Floor Pharmacy Lab behind Student Lift',
        '"In case of duty swapping, bring to the notice of the C.O.E a day prior duly approved by the C.O.E and signed application by both the faculties."',
        '',
        '"Reporting Time for Shift S1 is 9:00 AM Shift S2 is 11:00 AM, S3 is 1:00 PM and S4 is 3:00 PM. Invigilators to kindly frisk the students  if need be to check for any mobile phones and smart gadgets or any other material corroborating to Unfair means. Kindly permit students only with hard copy of admit cards in the Examination room. Teachers are required to be present in the designated room for the entire duration of the examination. Internal swapping or exchange of duties to be completely avoided . Shift Timings . S1 :- 9:30 AM - 11:00 AM     S2 :- 11:30 AM - 1:00 PM       S3 :- 1:30 PM - 3:00 PM          S4 :- 3:30 PM- 5:00 PM"',
        '',
        '',
        '',
        '',
        '',
        '',
        '',
        'S.No.,Name,Department,Designation,' + examDate + ',,,,,,,,,,,,,,,,,,,,,,,',
        ',,,,S1,S2,S3,S4,,,,,,,,,,,,,,,,,,,,'
    ];

    // Create teacher duty assignments
    const teacherDuties = [];
    let serialNumber = 1;

    teachers.forEach(teacher => {
        const dutyRow = {
            serialNo: serialNumber,
            name: teacher.name,
            department: teacher.department || 'N/A',
            designation: teacher.designation || 'N/A',
            s1: '0',
            s2: '0',
            s3: '0',
            s4: '0'
        };

        // Check which shifts this teacher is assigned to
        Object.keys(allocations).forEach(shift => {
            Object.keys(allocations[shift]).forEach(roomId => {
                const roomTeachers = allocations[shift][roomId];
                const isAssigned = roomTeachers.some(t => t.id === teacher.id);
                if (isAssigned) {
                    dutyRow[shift.toLowerCase()] = '1';
                }
            });
        });

        teacherDuties.push(dutyRow);
        serialNumber++;
    });

    // Convert to CSV format
    const csvRows = header.map(row => row);
    
    teacherDuties.forEach(duty => {
        const row = [
            duty.serialNo,
            duty.name,
            duty.department,
            duty.designation,
            duty.s1,
            duty.s2,
            duty.s3,
            duty.s4
        ];
        csvRows.push(row.join(','));
    });

    return csvRows.join('\n');
}

// Function to save duty chart to file
function saveDutyChartToFile(csvContent, filename) {
    const fs = require('fs');
    fs.writeFileSync(filename, csvContent, 'utf-8');
    console.log(`Duty chart saved to: ${filename}`);
}

// Example usage - Generate duty chart using current allocation logic
console.log('=== Generating Duty Chart using Current Allocation Logic ===');
const result = allocateDuties(rooms, teachers, "mid-sem");

// Generate duty chart output
const examDate = '03-04-2025';
const dutyChartCSV = generateDutyChartCSV(result, teachers, examDate);

// Save to file
const outputFilename = `Generated_Duty_Chart_${examDate}.csv`;
saveDutyChartToFile(dutyChartCSV, outputFilename);

console.log('Generated Duty Chart:');
console.log(dutyChartCSV);

// Also demonstrate the Excel-based allocation
console.log('\n=== Generating Duty Chart using Excel Schedule ===');
const dutyChartPath = path.join(__dirname, 'Duty Chart for 3rd April.csv');
const dutySchedule = parseDutyChartCSV(dutyChartPath);
const excelResult = allocateDutiesFromSchedule(rooms, teachers, dutySchedule);

const excelDutyChartCSV = generateDutyChartCSV(excelResult, teachers, examDate);
const excelOutputFilename = `Generated_Duty_Chart_Excel_${examDate}.csv`;
saveDutyChartToFile(excelDutyChartCSV, excelOutputFilename);

console.log('Generated Duty Chart (Excel-based):');
console.log(excelDutyChartCSV);
