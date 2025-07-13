/**
 * Script for Exam Duty Allocation UI
 * Integrates with the existing allocation logic.
 * Updated to parse CSV files and run allocation in browser.
 */

// Function to parse CSV file using PapaParse and return a Promise with data array
function parseCSVFile(file) {
    return new Promise((resolve, reject) => {
        Papa.parse(file, {
            header: true,
            skipEmptyLines: true,
            complete: function(results) {
                resolve(results.data);
            },
            error: function(err) {
                reject(err);
            }
        });
    });
}

// Allocation function ported from examDutyAllocator.js for browser use
function allocateDuties(rooms, teachers, examType) {
    const allocations = {};
    const shifts = examType === "mid-sem" ? 4 : examType === "end-sem" ? 2 : 1;

    // Filter lady teachers
    const ladyTeachers = teachers.filter(t => t.gender && t.gender.toLowerCase() === "female");
    // Filter male teachers
    const maleTeachers = teachers.filter(t => t.gender && t.gender.toLowerCase() === "male");

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
            if (!room.id) {
                console.warn("Room is missing an 'id' property:", room);
            }
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

// Store latest allocations and output for export
let latestAllocations = null;
let latestOutput = '';

// Event listener for allocation button with CSV file uploads
document.getElementById("allocateBtn").addEventListener("click", async () => {
    const teachersFileInput = document.getElementById("teachersFileInput");
    const roomsFileInput = document.getElementById("roomsFileInput");
    const examDatesFileInput = document.getElementById("examDatesInput");
    const examType = document.getElementById("examTypeSelect").value;

    if (teachersFileInput.files.length === 0) {
        alert("Please upload a teachers CSV file.");
        return;
    }
    if (roomsFileInput.files.length === 0) {
        alert("Please upload a rooms CSV file.");
        return;
    }
    if (examDatesFileInput.files.length === 0) {
        alert("Please upload an examination dates CSV file.");
        return;
    }

    try {
        const teachers = await parseCSVFile(teachersFileInput.files[0]);
        const rooms = await parseCSVFile(roomsFileInput.files[0]);
        const examDatesData = await parseCSVFile(examDatesFileInput.files[0]);

        // Extract exam dates from examDatesData by finding a column with name containing 'date' (case-insensitive)
        let dateColumn = null;
        if (examDatesData.length > 0) {
            const columns = Object.keys(examDatesData[0]);
            for (const col of columns) {
                if (col.toLowerCase().includes('date')) {
                    dateColumn = col;
                    break;
                }
            }
        }

        if (!dateColumn) {
            alert("No date column found in the uploaded exam dates file.");
            return;
        }

        const examDates = examDatesData.map(row => row[dateColumn]).filter(d => d && d.trim());

        console.log("Parsed exam dates:", examDates);

        if (examDates.length === 0) {
            alert("No exam dates found in the uploaded file.");
            return;
        }

        if (teachers.length === 0) {
            alert("No teachers found in the uploaded file. Please check your CSV.");
            return;
        }

        // Run allocation
        const allocations = allocateDuties(rooms, teachers, examType);
        latestAllocations = allocations; // Store for export

        // Format output for display
        let output = "";
        Object.keys(allocations).forEach(shift => {
            output += `Shift ${shift}:\n`;
            Object.keys(allocations[shift]).forEach(roomId => {
                output += `  Room ${roomId}:\n`;
                allocations[shift][roomId].forEach(teacher => {
                    if (teacher.error) {
                        output += `    Error: ${teacher.error}\n`;
                    } else {
                        output += `    Teacher: ${teacher.name} (${teacher.gender}, ${teacher.designation || 'N/A'})\n`;
                    }
                });
            });
            output += "\n";
        });
        latestOutput = output; // Store for export
        document.getElementById("resultsOutput").textContent = output;

    } catch (error) {
        alert("Error parsing files: " + error.message);
    }
});

// Excel export function
function exportAllocationsToExcel() {
    if (!latestAllocations) {
        alert("No allocation results to export. Please run the allocation first.");
        return;
    }
    const rows = [];
    Object.keys(latestAllocations).forEach(shift => {
        Object.keys(latestAllocations[shift]).forEach(roomId => {
            latestAllocations[shift][roomId].forEach(teacher => {
                if (teacher.error) {
                    rows.push({
                        Shift: shift,
                        Room: roomId,
                        Name: '',
                        Gender: '',
                        Designation: '',
                        Error: teacher.error
                    });
                } else {
                    rows.push({
                        Shift: shift,
                        Room: roomId,
                        Name: teacher.name,
                        Gender: teacher.gender,
                        Designation: teacher.designation || '',
                        Error: ''
                    });
                }
            });
        });
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Allocations");
    XLSX.writeFile(wb, "exam_duty_allocations.xlsx");
}
// PDF export function
function exportAllocationsToPDF() {
    if (!latestOutput) {
        alert("No allocation results to export. Please run the allocation first.");
        return;
    }
    const doc = new window.jspdf.jsPDF();
    const lines = latestOutput.split('\n');
    let y = 10;
    lines.forEach(line => {
        doc.text(line, 10, y);
        y += 7;
        if (y > 280) {
            doc.addPage();
            y = 10;
        }
    });
    doc.save("exam_duty_allocations.pdf");
}
// Add event listeners for export buttons
document.getElementById("downloadExcelBtn").addEventListener("click", exportAllocationsToExcel);
document.getElementById("downloadPdfBtn").addEventListener("click", exportAllocationsToPDF);
