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

// Utility function to normalize room names
function normalizeRoomName(name) {
    return name.replace(/\s+/g, ' ').replace(/[-_,.()]/g, '').trim().toUpperCase();
}

// Utility function to parse teacher availability from faculty CSV
function parseTeachersWithAvailability(teachersRaw, dateShiftKeys) {
    return teachersRaw.map((row, idx) => {
        const teacher = {
            id: row['FacultyID'] || row['facultyid'] || (idx + 1), // Use FacultyID as unique ID
            name: row['Name'] || row['name'] || '',
            gender: row['Gender'] || row['gender'] || '',
            department: row['Department'] || row['department'] || '',
            designation: row['Designation'] || row['designation'] || '',
            availability: {}, // { '2025-06-02_S1': true, ... }
        };
        if (dateShiftKeys.length === 0) {
            // No per-shift columns: treat as available for all shifts (will be set later)
            teacher.availableForAll = true;
        } else {
            dateShiftKeys.forEach(key => {
                teacher.availability[key] = row[key] && row[key].trim() === '1';
            });
        }
        return teacher;
    });
}

// Utility function to parse rooms from CSV
function parseRooms(roomsRaw) {
    // Group by normalized room name and sum students
    const roomMap = {};
    roomsRaw.forEach(row => {
        const rawName = row['Room No.'] || row['Room'] || row['room'] || row['room no.'] || '';
        const name = normalizeRoomName(rawName);
        const students = parseInt(row['STR'] || row['Students'] || row['students'] || '0', 10) || 0;
        if (!roomMap[name]) {
            roomMap[name] = { id: name, name: rawName, students: 0 };
        }
        roomMap[name].students += students;
    });
    return Object.values(roomMap);
}

// Modified allocation logic to use teacher availability and avoid repeated CR assignments
function allocateDutiesWithAvailability(rooms, teachers, examType, date, shiftKey, allDateAllocations) {
    const allocations = {};
    const shifts = examType === "mid-sem" ? 4 : examType === "end-sem" ? 2 : 1;
    // Ensure all teacher ids are strings
    const ladyTeachers = shuffleArray(teachers.filter(t => t.gender && t.gender.toLowerCase() === "female" && t.availability[shiftKey])).map(t => latestTeacherMap[String(t.id)] || t);
    const maleTeachers = shuffleArray(teachers.filter(t => t.gender && t.gender.toLowerCase() === "male" && t.availability[shiftKey])).map(t => latestTeacherMap[String(t.id)] || t);
    const lastShiftAssigned = new Map();
    function canAssign(teacherId) {
        return !lastShiftAssigned.get(String(teacherId));
    }
    allocations[shiftKey] = {};
    // --- Declare avoidTeacherIds at the top so it is always defined ---
    let avoidTeacherIds = new Set();
    // --- Alternate shift logic ---
    let prevShiftKey = null;
    let shiftNumMatch = shiftKey.match(/^(\d{4}-\d{2}-\d{2})_S(\d)$/);
    if (shiftNumMatch) {
        const [_, datePart, shiftNum] = shiftNumMatch;
        const prevShiftNum = parseInt(shiftNum) - 1;
        if (prevShiftNum >= 1) {
            prevShiftKey = `${datePart}_S${prevShiftNum}`;
        }
    }
    // Find previous shift for this date using allDateAllocations
    let prevShiftTeachers = new Set();
    if (prevShiftKey && allDateAllocations && allDateAllocations[date]) {
        // prevShiftKey is like 'YYYY-MM-DD_S1', but in allDateAllocations[date] the keys are 'Shift 1', 'Shift 2', etc.
        // Extract the shift number from prevShiftKey
        const prevShiftNumMatch = prevShiftKey.match(/_S(\d)$/);
        if (prevShiftNumMatch) {
            const prevShiftNum = prevShiftNumMatch[1];
            const prevShiftObj = allDateAllocations[date][`Shift ${prevShiftNum}`];
            if (prevShiftObj) {
                Object.values(prevShiftObj).forEach(roomArr => {
                    roomArr.forEach(t => {
                        if (t && t.id) prevShiftTeachers.add(String(t.id));
                    });
                });
            }
        }
    }
    // Add these to avoidTeacherIds
    prevShiftTeachers.forEach(id => avoidTeacherIds.add(String(id)));
    // --- End alternate shift logic ---

    // NEW: Avoid assigning a teacher to consecutive shifts on the same day
    if (shiftNumMatch) {
        const [_, datePart, shiftNum] = shiftNumMatch;
        const prevShiftNum = parseInt(shiftNum) - 1;
        if (prevShiftNum >= 1 && allDateAllocations && allDateAllocations[date]) {
            const prevShiftObj = allDateAllocations[date][`Shift ${prevShiftNum}`];
            if (prevShiftObj) {
                Object.values(prevShiftObj).forEach(roomArr => {
                    roomArr.forEach(t => {
                        if (t && t.id) avoidTeacherIds.add(String(t.id));
                    });
                });
            }
        }
    }
    rooms.forEach(room => {
        allocations[shiftKey][room.id] = [];
        let requiredTeachers = 1;
        if (examType === "mid-sem") {
            if (room.name === "LT 402") {
                requiredTeachers = room.students > 100 ? 4 : 3;
            } else {
                requiredTeachers = room.students > 20 ? 2 : 1;
            }
        } else if (examType === "end-sem") {
            if (/^AGRI/i.test(room.name)) {
                requiredTeachers = 2;
            } else if (room.name === "LT 201") {
                requiredTeachers = 2;
            } else if (room.name === "CR 205" || room.name === "CR 604") {
                requiredTeachers = 3;
            } else if (/^CR/i.test(room.name)) {
                requiredTeachers = room.students <= 80 ? 2 : 3;
            } else if (room.name === "NEW CLASSROOM") {
                requiredTeachers = room.students > 140 ? 4 : 2;
            } else {
                requiredTeachers = room.students > 80 ? 3 : 2;
            }
        }
        let allocatedCount = 0;
        let ladyIndex = 0;
        let maleIndex = 0;
        // For CR rooms, avoid assigning the same teacher as previous shift/day
        let avoidTeacherIds = new Set();
        if (/^CR/i.test(room.name)) {
            Object.keys(teacherRoomHistory).forEach(teacherId => {
                const history = teacherRoomHistory[teacherId][room.id] || [];
                if (history.length > 0 && history[history.length - 1] === shiftKey) {
                    avoidTeacherIds.add(String(teacherId));
                }
            });
        }
        // --- Add alternate shift constraint to avoidTeacherIds ---
        if (prevShiftTeachers.size > 0) {
            prevShiftTeachers.forEach(id => avoidTeacherIds.add(String(id)));
        }
        // --- End alternate shift constraint ---

        // --- NEW: Ensure at least 1 male and 1 female if possible ---
        let assignedMale = false;
        let assignedFemale = false;
        let assignedTeachers = new Set();
        // Only do this if requiredTeachers >= 2 and both genders are available
        if (requiredTeachers >= 2 && ladyTeachers.length > 0 && maleTeachers.length > 0) {
            // Assign one female
            while (ladyIndex < ladyTeachers.length) {
                const teacher = ladyTeachers[ladyIndex];
                if (canAssign(teacher.id) && !avoidTeacherIds.has(String(teacher.id)) && !assignedTeachers.has(String(teacher.id))) {
                    const origTeacher = latestTeacherMap[String(teacher.id)] || teacher;
                    allocations[shiftKey][room.id].push(origTeacher);
                    lastShiftAssigned.set(String(teacher.id), true);
                    if (!teacherRoomHistory[teacher.id]) teacherRoomHistory[teacher.id] = {};
                    if (!teacherRoomHistory[teacher.id][room.id]) teacherRoomHistory[teacher.id][room.id] = [];
                    teacherRoomHistory[teacher.id][room.id].push(shiftKey);
                    teacherDutyCount[teacher.id] = (teacherDutyCount[teacher.id] || 0) + 1;
                    ladyIndex++;
                    allocatedCount++;
                    assignedFemale = true;
                    assignedTeachers.add(String(teacher.id));
                    break;
                }
                ladyIndex++;
            }
            // Assign one male
            while (maleIndex < maleTeachers.length) {
                const teacher = maleTeachers[maleIndex];
                if (canAssign(teacher.id) && !avoidTeacherIds.has(String(teacher.id)) && !assignedTeachers.has(String(teacher.id))) {
                    const origTeacher = latestTeacherMap[String(teacher.id)] || teacher;
                    allocations[shiftKey][room.id].push(origTeacher);
                    lastShiftAssigned.set(String(teacher.id), true);
                    if (!teacherRoomHistory[teacher.id]) teacherRoomHistory[teacher.id] = {};
                    if (!teacherRoomHistory[teacher.id][room.id]) teacherRoomHistory[teacher.id][room.id] = [];
                    teacherRoomHistory[teacher.id][room.id].push(shiftKey);
                    teacherDutyCount[teacher.id] = (teacherDutyCount[teacher.id] || 0) + 1;
                    maleIndex++;
                    allocatedCount++;
                    assignedMale = true;
                    assignedTeachers.add(String(teacher.id));
                    break;
                }
                maleIndex++;
            }
        }
        // --- END NEW ---

        while (allocatedCount < requiredTeachers) {
            let assigned = false;
            // Assign a lady teacher if not already assigned and one is available
            if (allocatedCount === 0 && ladyTeachers.length > 0 && !assignedFemale) {
                while (ladyIndex < ladyTeachers.length) {
                    const teacher = ladyTeachers[ladyIndex];
                    if (canAssign(teacher.id) && !avoidTeacherIds.has(String(teacher.id)) && !assignedTeachers.has(String(teacher.id))) {
                        const origTeacher = latestTeacherMap[String(teacher.id)] || teacher;
                        allocations[shiftKey][room.id].push(origTeacher);
                        lastShiftAssigned.set(String(teacher.id), true);
                        if (!teacherRoomHistory[teacher.id]) teacherRoomHistory[teacher.id] = {};
                        if (!teacherRoomHistory[teacher.id][room.id]) teacherRoomHistory[teacher.id][room.id] = [];
                        teacherRoomHistory[teacher.id][room.id].push(shiftKey);
                        teacherDutyCount[teacher.id] = (teacherDutyCount[teacher.id] || 0) + 1;
                        ladyIndex++;
                        allocatedCount++;
                        assigned = true;
                        assignedFemale = true;
                        assignedTeachers.add(String(teacher.id));
                        console.log('ALLOCATE: Assigning teacher', teacher.id, teacher.name, 'to', shiftKey, room.id);
                        break;
                    }
                    ladyIndex++;
                }
            } else if (allocatedCount === 0 && maleTeachers.length > 0 && !assignedMale) {
                while (maleIndex < maleTeachers.length) {
                    const teacher = maleTeachers[maleIndex];
                    if (canAssign(teacher.id) && !avoidTeacherIds.has(String(teacher.id)) && !assignedTeachers.has(String(teacher.id))) {
                        const origTeacher = latestTeacherMap[String(teacher.id)] || teacher;
                        allocations[shiftKey][room.id].push(origTeacher);
                        lastShiftAssigned.set(String(teacher.id), true);
                        if (!teacherRoomHistory[teacher.id]) teacherRoomHistory[teacher.id] = {};
                        if (!teacherRoomHistory[teacher.id][room.id]) teacherRoomHistory[teacher.id][room.id] = [];
                        teacherRoomHistory[teacher.id][room.id].push(shiftKey);
                        teacherDutyCount[teacher.id] = (teacherDutyCount[teacher.id] || 0) + 1;
                        maleIndex++;
                        allocatedCount++;
                        assigned = true;
                        assignedMale = true;
                        assignedTeachers.add(String(teacher.id));
                        console.log('ALLOCATE: Assigning teacher', teacher.id, teacher.name, 'to', shiftKey, room.id);
                        break;
                    }
                    maleIndex++;
                }
            } else {
                // Fill remaining slots with any available teacher (lady or male)
                while (ladyIndex < ladyTeachers.length && allocatedCount < requiredTeachers) {
                    const teacher = ladyTeachers[ladyIndex];
                    if (canAssign(teacher.id) && !avoidTeacherIds.has(String(teacher.id)) && !assignedTeachers.has(String(teacher.id))) {
                        const origTeacher = latestTeacherMap[String(teacher.id)] || teacher;
                        allocations[shiftKey][room.id].push(origTeacher);
                        lastShiftAssigned.set(String(teacher.id), true);
                        if (!teacherRoomHistory[teacher.id]) teacherRoomHistory[teacher.id] = {};
                        if (!teacherRoomHistory[teacher.id][room.id]) teacherRoomHistory[teacher.id][room.id] = [];
                        teacherRoomHistory[teacher.id][room.id].push(shiftKey);
                        teacherDutyCount[teacher.id] = (teacherDutyCount[teacher.id] || 0) + 1;
                        ladyIndex++;
                        allocatedCount++;
                        assigned = true;
                        assignedTeachers.add(String(teacher.id));
                        console.log('ALLOCATE: Assigning teacher', teacher.id, teacher.name, 'to', shiftKey, room.id);
                        break;
                    }
                    ladyIndex++;
                }
                while (!assigned && maleIndex < maleTeachers.length && allocatedCount < requiredTeachers) {
                    const teacher = maleTeachers[maleIndex];
                    if (canAssign(teacher.id) && !avoidTeacherIds.has(String(teacher.id)) && !assignedTeachers.has(String(teacher.id))) {
                        const origTeacher = latestTeacherMap[String(teacher.id)] || teacher;
                        allocations[shiftKey][room.id].push(origTeacher);
                        lastShiftAssigned.set(String(teacher.id), true);
                        if (!teacherRoomHistory[teacher.id]) teacherRoomHistory[teacher.id] = {};
                        if (!teacherRoomHistory[teacher.id][room.id]) teacherRoomHistory[teacher.id][room.id] = [];
                        teacherRoomHistory[teacher.id][room.id].push(shiftKey);
                        teacherDutyCount[teacher.id] = (teacherDutyCount[teacher.id] || 0) + 1;
                        maleIndex++;
                        allocatedCount++;
                        assigned = true;
                        assignedTeachers.add(String(teacher.id));
                        console.log('ALLOCATE: Assigning teacher', teacher.id, teacher.name, 'to', shiftKey, room.id);
                        break;
                    }
                    maleIndex++;
                }
            }
            // If not enough teachers, allow assignment even if repeated, but pick those with lowest total duties
            if (!assigned) {
                let allAvailable = ladyTeachers.concat(maleTeachers).filter(t => canAssign(t.id) && !assignedTeachers.has(String(t.id)));
                if (allAvailable.length > 0) {
                    allAvailable.sort((a, b) => (teacherDutyCount[a.id] || 0) - (teacherDutyCount[b.id] || 0));
                    const teacher = allAvailable[0];
                    const origTeacher = latestTeacherMap[String(teacher.id)] || teacher;
                    allocations[shiftKey][room.id].push(origTeacher);
                    lastShiftAssigned.set(String(teacher.id), true);
                    if (!teacherRoomHistory[teacher.id]) teacherRoomHistory[teacher.id] = {};
                    if (!teacherRoomHistory[teacher.id][room.id]) teacherRoomHistory[teacher.id][room.id] = [];
                    teacherRoomHistory[teacher.id][room.id].push(shiftKey);
                    teacherDutyCount[teacher.id] = (teacherDutyCount[teacher.id] || 0) + 1;
                    allocatedCount++;
                    assignedTeachers.add(String(teacher.id));
                    assigned = true;
                    console.log('ALLOCATE: Assigning teacher', teacher.id, teacher.name, 'to', shiftKey, room.id);
                } else {
                    allocations[shiftKey][room.id].push({ error: "Not enough teachers available" });
                    break;
                }
            }
        }
    });
    return allocations;
}

// Utility function to shuffle an array (Fisher-Yates)
function shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
        const j = Math.floor(Math.random() * (i + 1));
        [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
}

// Store latest allocations and output for export
let latestAllocations = null;
let latestOutput = '';
// Track teacher-room assignment history
let teacherRoomHistory = {};
// Track total duties per teacher
let teacherDutyCount = {};
// Store latest teachers for export
let latestTeachers = [];
let latestTeacherMap = {};
// Store latest exam type for export
let latestExamType = '';

// Event listener for allocation button with CSV file uploads
document.getElementById("allocateBtn").addEventListener("click", async () => {
    const teachersFileInput = document.getElementById("teachersFileInput");
    const roomsFileInput = document.getElementById("roomsFileInput");
    const examDatesFileInput = document.getElementById("examDatesInput");
    const examType = document.getElementById("examTypeSelect").value;
    latestExamType = examType; // Store for export
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
        const teachersRaw = await parseCSVFile(teachersFileInput.files[0]);
        const roomsRaw = await parseCSVFile(roomsFileInput.files[0]);
        const examDatesRaw = await parseCSVFile(examDatesFileInput.files[0]);
        // Get all date_shift keys from teachersRaw
        const dateShiftKeys = Object.keys(teachersRaw[0] || {}).filter(k => /\d{4}-\d{2}-\d{2}_S\d/.test(k));
        const teachers = parseTeachersWithAvailability(teachersRaw, dateShiftKeys);
        latestTeachers = teachers; // <-- Store for export
        latestTeacherMap = {};
        teachers.forEach(t => { latestTeacherMap[String(t.id)] = t; });
        const rooms = parseRooms(roomsRaw);
        // Parse dates from examDatesRaw (support both DD-MM-YYYY and YYYY-MM-DD)
        const examDates = examDatesRaw.map(row => Object.values(row)[0]).filter(d => d && d.trim());
        // For each date and shift, allocate
        const allDateAllocations = {};
        const numShifts = examType === "mid-sem" ? 4 : 2;
        examDates.forEach(date => {
            if (!allDateAllocations[date]) allDateAllocations[date] = {};
            for (let shift = 1; shift <= numShifts; shift++) {
                const shiftKey = `${date.replace(/\//g, '-').replace(/(\d{2})-(\d{2})-(\d{4})/, '$3-$2-$1')}_S${shift}`;
                // Only consider teachers available for this date/shift
                let availableTeachers;
                if (dateShiftKeys.length === 0) {
                    availableTeachers = teachers.map(t => ({ ...t, availability: { [shiftKey]: true } }));
                    teachers.forEach(t => { t.availability[shiftKey] = true; });
                } else {
                    availableTeachers = teachers.filter(t => t.availability[shiftKey]);
                }
                shuffleArray(availableTeachers);
                // Pass the current state of allDateAllocations to the allocation function
                const allocations = allocateDutiesWithAvailability(
                    rooms, teachers, examType, date, shiftKey, allDateAllocations
                );
                allDateAllocations[date][`Shift ${shift}`] = allocations[shiftKey];
            }
        });
        latestAllocations = allDateAllocations; // Store for export
        // Format output for display (HTML for highlighting)
        let output = "";
        Object.keys(allDateAllocations).forEach(date => {
            output += `<div><strong>Date: ${date}</strong></div>`;
            const allocationsByShift = allDateAllocations[date];
            Object.keys(allocationsByShift).forEach(shift => {
                output += `<div style='margin-left:1em;'><strong>${shift}:</strong></div>`;
                const roomsObj = allocationsByShift[shift];
                Object.keys(roomsObj).forEach(roomId => {
                    output += `<div style='margin-left:2em;'>Room ${roomId}:</div>`;
                    roomsObj[roomId].forEach(teacher => {
                        if (teacher.error) {
                            output += `<div style='margin-left:3em;'><span class='error-highlight'>Error: ${teacher.error}</span></div>`;
                        } else {
                            output += `<div style='margin-left:3em;'>Teacher: ${teacher.name} (${teacher.gender}, ${teacher.designation || 'N/A'})</div>`;
                        }
                    });
                });
                output += "<br/>";
            });
            output += "<br/>";
        });
        latestOutput = output; // Store for export
        document.getElementById("resultsOutput").innerHTML = output;
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
    Object.keys(latestAllocations).forEach(date => {
        const allocations = latestAllocations[date];
        Object.keys(allocations).forEach(shift => {
            Object.keys(allocations[shift]).forEach(roomId => {
                allocations[shift][roomId].forEach(teacher => {
                    if (teacher.error) {
                        rows.push({
                            Date: date,
                            Shift: shift,
                            Room: roomId,
                            Name: '',
                            Gender: '',
                            Designation: '',
                            Error: teacher.error
                        });
                    } else {
                        rows.push({
                            Date: date,
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
    });
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Allocations");
    XLSX.writeFile(wb, "exam_duty_allocations.xlsx");
}
// PDF export function
function exportAllocationsToPDF() {
    if (!latestAllocations) {
        alert("No allocation results to export. Please run the allocation first.");
        return;
    }
    const doc = new window.jspdf.jsPDF();
    let y = 10;
    const pageWidth = doc.internal.pageSize.getWidth();
    Object.keys(latestAllocations).forEach(date => {
        doc.setFont(undefined, 'bold');
        doc.text(`Date: ${date}`, 10, y);
        y += 8;
        const allocationsByShift = latestAllocations[date];
        Object.keys(allocationsByShift).forEach(shift => {
            doc.setFont(undefined, 'normal');
            doc.text(`  ${shift}:`, 12, y);
            y += 7;
            const roomsObj = allocationsByShift[shift];
            Object.keys(roomsObj).forEach(roomId => {
                doc.setFont(undefined, 'bold');
                doc.text(`    Room ${roomId}`, 16, y);
                y += 7;
                doc.setFont(undefined, 'normal');
                // Table header
                doc.text('Name', 22, y, { align: 'left' });
                doc.text('Designation', pageWidth - 22, y, { align: 'right' });
                y += 6;
                // Table rows
                roomsObj[roomId].forEach(teacher => {
                    if (teacher.error) {
                        doc.setTextColor(200, 0, 0);
                        doc.text(`Error: ${teacher.error}`, 22, y);
                        doc.setTextColor(0, 0, 0);
                        y += 6;
                    } else {
                        doc.text(teacher.name, 22, y, { align: 'left' });
                        doc.text(teacher.designation || 'N/A', pageWidth - 22, y, { align: 'right' });
                        y += 6;
                    }
                    if (y > 280) {
                        doc.addPage();
                        y = 10;
                    }
                });
                y += 3;
            });
            y += 2;
        });
        y += 4;
    });
    doc.save("exam_duty_allocations.pdf");
}
// Faculty-wise Duty Schedule Export
function exportFacultyDutyScheduleToExcel() {
    if (!latestAllocations) {
        alert("No allocation results to export. Please run the allocation first.");
        return;
    }
    if (!latestTeachers || latestTeachers.length === 0) {
        alert("No teacher list found. Please allocate duties first.");
        return;
    }
    // Debug: print all FacultyIDs in latestTeachers
    console.log('DEBUG: FacultyIDs in latestTeachers:', latestTeachers.map(t => t.id));
    // Use teachers from the original CSV
    const facultyList = latestTeachers.map(t => ({
        ...t,
        originalName: t.name,
        normalizedName: t.name.trim().toLowerCase(),
        id: t.id // ensure id is present
    }));
    // Set number of shifts per day based on exam type
    let numShifts = 2; // default to end-sem
    if (latestExamType === 'mid-sem') numShifts = 4;
    if (latestExamType === 'end-sem') numShifts = 2;
    // Gather all date/shift keys from latestAllocations structure
    const dateShiftKeysRaw = [];
    Object.keys(latestAllocations).forEach(date => {
        Object.keys(latestAllocations[date]).forEach(shift => {
            dateShiftKeysRaw.push({ raw: `${date} ${shift}`, date, shift });
        });
    });
    // Format headers for display
    const dateShiftKeys = dateShiftKeysRaw.map(({ date, shift }) => {
        // Convert YYYY-MM-DD to DD-MM-YYYY if possible
        let displayDate = date;
        const match = date.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (match) {
            displayDate = `${match[3]}-${match[2]}-${match[1]}`;
        }
        return `${displayDate} ${shift}`;
    });
    console.log('DEBUG: dateShiftKeys in export:', dateShiftKeys);
    // Build matrix
    const staticHeaders = ['S.No.', 'Name', 'Gender', 'Department', 'Designation'];
    const rows = facultyList.map((teacher, idx) => {
        const row = {
            'S.No.': idx + 1,
            'Name': teacher.originalName,
            'Gender': teacher.gender,
            'Department': teacher.department || '',
            'Designation': teacher.designation || ''
        };
        dateShiftKeys.forEach(key => {
            row[key] = '';
        });
        return row;
    });
    // Fill matrix: mark '1' if assigned (not just available)
    facultyList.forEach((teacher, idx) => {
        dateShiftKeysRaw.forEach(({ raw }, colIdx) => {
            let assigned = false;
            const [date, ...shiftParts] = raw.split(' ');
            const shift = shiftParts.join(' ');
            if (latestAllocations[date] && latestAllocations[date][shift]) {
                for (const roomArr of Object.values(latestAllocations[date][shift])) {
                    if (roomArr.some(t => t && t.id && String(t.id) === String(teacher.id))) {
                        assigned = true;
                    }
                }
            }
            if (assigned) {
                rows[idx][dateShiftKeys[colIdx]] = '1';
            }
        });
    });
    // Warn if any assigned teacher is not matched in export
    Object.keys(latestAllocations).forEach(date => {
        Object.keys(latestAllocations[date]).forEach(shiftKey => {
            Object.values(latestAllocations[date][shiftKey]).forEach(roomArr => {
                roomArr.forEach(t => {
                    if (t && t.name && !t.error) {
                        const normalizedName = t.name.trim().toLowerCase();
                        const match = facultyList.find(f => f.normalizedName === normalizedName);
                        if (!match) {
                            console.warn(`WARNING: Assigned teacher '${t.name}' not found in faculty export list.`);
                        }
                    }
                });
            });
        });
    });
    // Sheet with explicit header order
    const ws = XLSX.utils.json_to_sheet(rows, { header: staticHeaders.concat(dateShiftKeys) });
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Faculty Duty Schedule");
    XLSX.writeFile(wb, "faculty_duty_schedule.xlsx");
}
// Room-wise Allocation Export
function exportRoomWiseAllocationToExcel() {
    if (!latestAllocations) {
        alert("No allocation results to export. Please run the allocation first.");
        return;
    }
    // Gather all unique rooms
    const roomSet = new Set();
    Object.values(latestAllocations).forEach(dateObj => {
        Object.values(dateObj).forEach(shiftObj => {
            Object.keys(shiftObj).forEach(roomId => {
                roomSet.add(roomId);
            });
        });
    });
    const roomList = Array.from(roomSet);
    // Gather all date/shift keys
    const dateShiftKeys = [];
    Object.keys(latestAllocations).forEach(date => {
        Object.keys(latestAllocations[date]).forEach(shift => {
            dateShiftKeys.push(`${date} ${shift}`);
        });
    });
    // Build matrix
    const rows = roomList.map(roomId => {
        const row = { 'Room': roomId };
        dateShiftKeys.forEach(key => {
            row[key] = '';
        });
        return row;
    });
    // Fill matrix: list teacher names
    roomList.forEach((roomId, idx) => {
        let colIdx = 0;
        Object.keys(latestAllocations).forEach(date => {
            Object.keys(latestAllocations[date]).forEach(shift => {
                const key = `${date} ${shift}`;
                const shiftObj = latestAllocations[date][shift];
                if (shiftObj[roomId]) {
                    const names = shiftObj[roomId]
                        .filter(t => t && t.name && !t.error)
                        .map(t => t.name)
                        .join(', ');
                    rows[idx][key] = names;
                }
                colIdx++;
            });
        });
    });
    // Sheet
    const ws = XLSX.utils.json_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Room-wise Allocation");
    XLSX.writeFile(wb, "room_wise_allocation.xlsx");
}
// Add event listeners for export buttons
document.getElementById("downloadExcelBtn").addEventListener("click", exportAllocationsToExcel);
document.getElementById("downloadPdfBtn").addEventListener("click", exportAllocationsToPDF);
document.getElementById("downloadFacultyScheduleBtn").addEventListener("click", exportFacultyDutyScheduleToExcel);
document.getElementById("downloadRoomWiseBtn").addEventListener("click", exportRoomWiseAllocationToExcel);
