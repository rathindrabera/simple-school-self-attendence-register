let currentDate = new Date();
document.addEventListener('DOMContentLoaded', () => {
  updateCalendar();
});

function addAttendance() {
  const today = new Date().toISOString().split('T')[0];
  const location = "School";
  
  const attendance = JSON.parse(localStorage.getItem('attendance')) || [];
  
  // Check if today's attendance already exists
  if (attendance.some(entry => entry.date === today)) {
    alert("Attendance has already been added for today.");
    return;
  }

  attendance.push({ date: today, location });
  localStorage.setItem('attendance', JSON.stringify(attendance));

  updateCalendar();
}

function updateCalendar() {
  const year = currentDate.getFullYear();
  const month = currentDate.getMonth();

  document.getElementById('currentMonth').textContent = currentDate.toLocaleString('default', { month: 'long', year: 'numeric' });

  const attendance = JSON.parse(localStorage.getItem('attendance')) || [];
  const daysInMonth = new Date(year, month + 1, 0).getDate();

  const monthlyAttendance = attendance.filter(entry => {
    const entryDate = new Date(entry.date);
    return entryDate.getFullYear() === year && entryDate.getMonth() === month;
  });

  document.getElementById('attendanceLabel').textContent = `Total attendance this month: ${monthlyAttendance.length}`;

  const dayNames = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  
  const calendar = document.getElementById('calendar');
  calendar.innerHTML = '';

  dayNames.forEach(day => {
    const dayHeader = document.createElement('div');
    dayHeader.className = 'weekday-header';
    dayHeader.textContent = day;
    calendar.appendChild(dayHeader);
  });

  const firstDayOfMonth = new Date(year, month, 1).getDay();

  for (let i = 0; i < firstDayOfMonth; i++) {
    const emptyCell = document.createElement('div');
    calendar.appendChild(emptyCell);
  }

  for (let day = 1; day <= daysInMonth; day++) {
    const dayDate = new Date(year, month, day).toISOString().split('T')[0];
    const dayOfWeek = new Date(year, month, day).getDay();

    const dayElement = document.createElement('div');
    dayElement.className = 'day';
    dayElement.textContent = day;

    if (dayOfWeek === 0) {
      dayElement.classList.add('sunday');
    }

    if (monthlyAttendance.some(entry => entry.date === dayDate)) {
      dayElement.classList.add('attended');
    }

    calendar.appendChild(dayElement);
  }
}

function changeMonth(direction) {
  currentDate.setMonth(currentDate.getMonth() + direction);
  updateCalendar();
}

// Function to export attendance data to Excel
function exportToExcel() {
  const attendance = JSON.parse(localStorage.getItem('attendance')) || [];

  if (attendance.length === 0) {
    alert("No attendance data available to export.");
    return;
  }

  // Prepare data for Excel
  const worksheetData = attendance.map((entry, index) => ({
    "S.No": index + 1,
    "Date": entry.date,
    "Location": entry.location
  }));

  const worksheet = XLSX.utils.json_to_sheet(worksheetData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Attendance");

  // Export the Excel file
  XLSX.writeFile(workbook, "Attendance.xlsx");
}