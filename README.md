<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Aniruddha's Business Portal</title>
  <script src="https://cdn.tailwindcss.com"></script>
  <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css" rel="stylesheet">
</head>
<body class="bg-slate-100 font-sans text-slate-800 min-h-screen flex flex-col">

  <!-- Header / Navigation -->
  <header class="bg-indigo-700 text-white shadow-lg sticky top-0 z-50">
    <div class="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center">
      <div class="flex items-center space-x-3">
        <i class="fa-solid me-2 fa-hotel text-2xl text-indigo-200"></i>
        <h1 class="text-xl font-bold tracking-wide">Aniruddha Business Portal</h1>
      </div>
      <nav class="flex space-x-2">
        <button onclick="switchTab('dashboard')" id="tab-dashboard" class="px-4 py-2 rounded-lg text-sm font-medium transition bg-indigo-900 text-white">
          <i class="fa-solid fa-gauge me-1"></i> Dashboard
        </button>
        <button onclick="switchTab('bookings')" id="tab-bookings" class="px-4 py-2 rounded-lg text-sm font-medium transition hover:bg-indigo-600 text-indigo-100">
          <i class="fa-solid fa-book-bookmark me-1"></i> Bookings
        </button>
        <button onclick="switchTab('agents')" id="tab-agents" class="px-4 py-2 rounded-lg text-sm font-medium transition hover:bg-indigo-600 text-indigo-100">
          <i class="fa-solid fa-user-tie me-1"></i> Agents & Rooms
        </button>
        <button onclick="switchTab('calendar')" id="tab-calendar" class="px-4 py-2 rounded-lg text-sm font-medium transition hover:bg-indigo-600 text-indigo-100">
          <i class="fa-solid fa-calendar-days me-1"></i> Calendar
        </button>
      </nav>
    </div>
  </header>

  <!-- Main Content Container -->
  <main class="flex-grow max-w-7xl w-full mx-auto p-4 sm:p-6">

    <!-- 1. DASHBOARD VIEW -->
    <section id="view-dashboard" class="space-y-6">
      <div class="bg-gradient-to-r from-indigo-600 to-blue-600 text-white rounded-2xl p-6 shadow-md">
        <h2 class="text-3xl font-extrabold mb-2">Hi Aniruddha 👋</h2>
        <p class="text-indigo-100 text-lg">Welcome to your business portal! Manage room bookings, view agents, and track schedules seamlessly.</p>
      </div>

      <!-- Quick Metrics -->
      <div class="grid grid-cols-1 md:grid-cols-4 gap-4">
        <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Total Rooms</p>
          <p id="stat-total-rooms" class="text-2xl font-bold text-slate-800 mt-1">5</p>
        </div>
        <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Active Agents</p>
          <p id="stat-total-agents" class="text-2xl font-bold text-indigo-600 mt-1">3</p>
        </div>
        <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Total Bookings</p>
          <p id="stat-total-bookings" class="text-2xl font-bold text-emerald-600 mt-1">0</p>
        </div>
        <div class="bg-white p-5 rounded-xl shadow-sm border border-slate-200">
          <p class="text-xs font-semibold text-slate-500 uppercase">Pending Due</p>
          <p id="stat-total-due" class="text-2xl font-bold text-amber-600 mt-1">₹0</p>
        </div>
      </div>

      <!-- Year Selector Grid (2026 - 2085) -->
      <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
        <h3 class="text-lg font-bold mb-4 text-slate-800 flex items-center">
          <i class="fa-solid fa-calendar text-indigo-600 me-2"></i> Fast Navigation Years (2026 - 2085)
        </h3>
        <div class="grid grid-cols-4 sm:grid-cols-6 md:grid-cols-10 gap-2 max-h-64 overflow-y-auto p-1" id="years-grid"></div>
      </div>
    </section>

    <!-- 2. BOOKINGS VIEW -->
    <section id="view-bookings" class="hidden space-y-6">
      <div class="grid grid-cols-1 lg:grid-cols-3 gap-6">
        <!-- New Booking Form -->
        <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200 lg:col-span-1">
          <h3 class="text-lg font-bold text-slate-800 mb-4 pb-2 border-b">
            <i class="fa-solid fa-user-plus text-indigo-600 me-2"></i> New Customer Booking
          </h3>
          <form id="bookingForm" onsubmit="saveBooking(event)" class="space-y-3">
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">Customer Name</label>
              <input type="text" id="custName" required class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">Address</label>
              <input type="text" id="custAdd" required class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
            </div>
            <div class="grid grid-cols-2 gap-2">
              <div>
                <label class="block text-xs font-semibold text-slate-600 mb-1">ID Number</label>
                <input type="text" id="custId" required class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
              </div>
              <div>
                <label class="block text-xs font-semibold text-slate-600 mb-1">Contact No</label>
                <input type="tel" id="custContact" required class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
              </div>
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">Select Room & Agent</label>
              <select id="roomAgentSelect" onchange="autoFillAgentDetails()" required class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
                <option value="">-- Choose Room --</option>
              </select>
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">Agent Phone</label>
              <input type="text" id="agentPhone" readonly class="w-full bg-slate-100 border rounded-lg p-2 text-sm text-slate-500">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">Booking Date</label>
              <input type="date" id="bookingDate" required class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
            </div>
            <div class="grid grid-cols-2 gap-2">
              <div>
                <label class="block text-xs font-semibold text-slate-600 mb-1">Advance Paid (₹)</label>
                <input type="number" id="advancePaid" value="0" min="0" oninput="calculateDue()" class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
              </div>
              <div>
                <label class="block text-xs font-semibold text-slate-600 mb-1">Total Rate (₹)</label>
                <input type="number" id="totalRate" value="0" min="0" oninput="calculateDue()" class="w-full border rounded-lg p-2 text-sm focus:ring-2 focus:ring-indigo-500 outline-none">
              </div>
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-600 mb-1">Total Due (₹)</label>
              <input type="number" id="totalDue" readonly class="w-full bg-amber-50 border border-amber-200 text-amber-700 font-bold rounded-lg p-2 text-sm">
            </div>
            <button type="submit" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-2 rounded-lg shadow-sm transition">
              Create Booking
            </button>
          </form>
        </div>

        <!-- Bookings List -->
        <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200 lg:col-span-2">
          <h3 class="text-lg font-bold text-slate-800 mb-4 pb-2 border-b">
            <i class="fa-solid fa-list text-indigo-600 me-2"></i> Booking Records
          </h3>
          <div class="overflow-x-auto">
            <table class="w-full text-left text-sm border-collapse">
              <thead>
                <tr class="bg-slate-100 text-slate-600 font-semibold border-b">
                  <th class="p-3">Customer</th>
                  <th class="p-3">Room</th>
                  <th class="p-3">Agent</th>
                  <th class="p-3">Date</th>
                  <th class="p-3">Advance</th>
                  <th class="p-3">Due</th>
                  <th class="p-3 text-center">Action</th>
                </tr>
              </thead>
              <tbody id="bookingsTableBody" class="divide-y"></tbody>
            </table>
          </div>
        </div>
      </div>
    </section>

    <!-- 3. AGENTS & ROOMS VIEW -->
    <section id="view-agents" class="hidden space-y-6">
      <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
        <div class="flex justify-between items-center mb-4">
          <h3 class="text-lg font-bold text-slate-800">
            <i class="fa-solid fa-door-open text-indigo-600 me-2"></i> Agent & Room Master Details
          </h3>
          <button onclick="addMasterRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded-lg text-sm font-medium transition">
            + Add Room / Agent
          </button>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-left text-sm border-collapse">
            <thead>
              <tr class="bg-slate-100 text-slate-600 font-semibold border-b">
                <th class="p-3">Agent Name</th>
                <th class="p-3">Agent Ph No</th>
                <th class="p-3">Room No</th>
                <th class="p-3">Capacity</th>
                <th class="p-3 text-center">Action</th>
              </tr>
            </thead>
            <tbody id="masterTableBody" class="divide-y"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- 4. CALENDAR VIEW -->
    <section id="view-calendar" class="hidden space-y-6">
      <div class="bg-white rounded-2xl p-6 shadow-sm border border-slate-200">
        <div class="flex flex-col sm:flex-row justify-between items-center gap-4 mb-6">
          <h3 class="text-lg font-bold text-slate-800">
            <i class="fa-regular fa-calendar-check text-indigo-600 me-2"></i> Interactive Availability Calendar
          </h3>
          <div class="flex items-center space-x-3">
            <select id="calYearSelect" onchange="renderCalendar()" class="border rounded-lg p-2 text-sm font-semibold bg-slate-50"></select>
            <select id="calMonthSelect" onchange="renderCalendar()" class="border rounded-lg p-2 text-sm font-semibold bg-slate-50">
              <option value="0">January</option>
              <option value="1">February</option>
              <option value="2">March</option>
              <option value="3">April</option>
              <option value="4">May</option>
              <option value="5">June</option>
              <option value="6">July</option>
              <option value="7">August</option>
              <option value="8">September</option>
              <option value="9">October</option>
              <option value="10">November</option>
              <option value="11">December</option>
            </select>
          </div>
        </div>

        <div class="grid grid-cols-7 gap-2 text-center font-bold text-slate-600 text-sm mb-2">
          <div class="p-2 text-rose-500">Sun</div>
          <div class="p-2">Mon</div>
          <div class="p-2">Tue</div>
          <div class="p-2">Wed</div>
          <div class="p-2">Thu</div>
          <div class="p-2">Fri</div>
          <div class="p-2">Sat</div>
        </div>
        <div id="calendar-grid" class="grid grid-cols-7 gap-2"></div>
      </div>
    </section>

  </main>

  <script>
    // Initial Master Data extracted from Excel Workbook
    let masterData = JSON.parse(localStorage.getItem('app_masterData')) || [
      { agentName: 'A1', agentPhone: '1234567890', roomNo: 1, capacity: 4 },
      { agentName: 'A2', agentPhone: '1234567890', roomNo: 2, capacity: 2 },
      { agentName: 'A3', agentPhone: '1234567890', roomNo: 3, capacity: 4 },
      { agentName: 'A1', agentPhone: '1234567890', roomNo: 4, capacity: 4 },
      { agentName: 'A3', agentPhone: '1234567890', roomNo: 5, capacity: 4 }
    ];

    let bookings = JSON.parse(localStorage.getItem('app_bookings')) || [];

    // Switch Application Views
    function switchTab(tab) {
      ['dashboard', 'bookings', 'agents', 'calendar'].forEach(t => {
        document.getElementById(`view-${t}`).classList.add('hidden');
        document.getElementById(`tab-${t}`).className = 'px-4 py-2 rounded-lg text-sm font-medium transition hover:bg-indigo-600 text-indigo-100';
      });
      document.getElementById(`view-${tab}`).classList.remove('hidden');
      document.getElementById(`tab-${tab}`).className = 'px-4 py-2 rounded-lg text-sm font-medium transition bg-indigo-900 text-white';

      if (tab === 'calendar') renderCalendar();
      if (tab === 'dashboard') updateDashboardStats();
    }

    // Populate Year Navigation (2026 - 2085)
    const yearsGrid = document.getElementById('years-grid');
    const calYearSelect = document.getElementById('calYearSelect');
    for (let yr = 2026; yr <= 2085; yr++) {
      const btn = document.createElement('button');
      btn.className = 'p-2 border rounded-lg text-xs font-semibold hover:bg-indigo-600 hover:text-white transition bg-slate-50 text-slate-700';
      btn.innerText = yr;
      btn.onclick = () => {
        calYearSelect.value = yr;
        switchTab('calendar');
      };
      yearsGrid.appendChild(btn);

      const opt = document.createElement('option');
      opt.value = yr;
      opt.innerText = yr;
      calYearSelect.appendChild(opt);
    }

    // Render Master Table
    function renderMasterTable() {
      const tbody = document.getElementById('masterTableBody');
      tbody.innerHTML = '';
      masterData.forEach((item, index) => {
        tbody.innerHTML += `
          <tr class="hover:bg-slate-50">
            <td class="p-3"><input type="text" value="${item.agentName}" onchange="updateMaster(${index}, 'agentName', this.value)" class="border rounded p-1 text-sm w-full"></td>
            <td class="p-3"><input type="text" value="${item.agentPhone}" onchange="updateMaster(${index}, 'agentPhone', this.value)" class="border rounded p-1 text-sm w-full"></td>
            <td class="p-3"><input type="number" value="${item.roomNo}" onchange="updateMaster(${index}, 'roomNo', this.value)" class="border rounded p-1 text-sm w-20"></td>
            <td class="p-3"><input type="number" value="${item.capacity}" onchange="updateMaster(${index}, 'capacity', this.value)" class="border rounded p-1 text-sm w-20"></td>
            <td class="p-3 text-center">
              <button onclick="deleteMaster(${index})" class="text-rose-600 hover:text-rose-800"><i class="fa-solid fa-trash"></i></button>
            </td>
          </tr>
        `;
      });
      populateRoomDropdown();
      updateDashboardStats();
      localStorage.setItem('app_masterData', JSON.stringify(masterData));
    }

    function addMasterRow() {
      masterData.push({ agentName: 'New Agent', agentPhone: '0000000000', roomNo: masterData.length + 1, capacity: 2 });
      renderMasterTable();
    }

    function updateMaster(index, field, value) {
      masterData[index][field] = field === 'roomNo' || field === 'capacity' ? Number(value) : value;
      renderMasterTable();
    }

    function deleteMaster(index) {
      masterData.splice(index, 1);
      renderMasterTable();
    }

    // Dynamic Select Populate
    function populateRoomDropdown() {
      const select = document.getElementById('roomAgentSelect');
      select.innerHTML = '<option value="">-- Choose Room --</option>';
      masterData.forEach(m => {
        select.innerHTML += `<option value="${m.roomNo}">Room ${m.roomNo} (Agent ${m.agentName}, Cap: ${m.capacity})</option>`;
      });
    }

    function autoFillAgentDetails() {
      const roomNo = Number(document.getElementById('roomAgentSelect').value);
      const match = masterData.find(m => m.roomNo === roomNo);
      document.getElementById('agentPhone').value = match ? match.agentPhone : '';
    }

    function calculateDue() {
      const adv = Number(document.getElementById('advancePaid').value) || 0;
      const rate = Number(document.getElementById('totalRate').value) || 0;
      document.getElementById('totalDue').value = Math.max(0, rate - adv);
    }

    // Save Booking
    function saveBooking(e) {
      e.preventDefault();
      const roomNo = Number(document.getElementById('roomAgentSelect').value);
      const agent = masterData.find(m => m.roomNo === roomNo);

      const booking = {
        id: Date.now(),
        custName: document.getElementById('custName').value,
        custAdd: document.getElementById('custAdd').value,
        custId: document.getElementById('custId').value,
        custContact: document.getElementById('custContact').value,
        roomNo: roomNo,
        agentName: agent ? agent.agentName : 'N/A',
        agentPhone: document.getElementById('agentPhone').value,
        date: document.getElementById('bookingDate').value,
        advance: Number(document.getElementById('advancePaid').value) || 0,
        rate: Number(document.getElementById('totalRate').value) || 0,
        due: Number(document.getElementById('totalDue').value) || 0
      };

      bookings.push(booking);
      localStorage.setItem('app_bookings', JSON.stringify(bookings));
      document.getElementById('bookingForm').reset();
      renderBookings();
      updateDashboardStats();
    }

    function renderBookings() {
      const tbody = document.getElementById('bookingsTableBody');
      tbody.innerHTML = '';
      if (bookings.length === 0) {
        tbody.innerHTML = '<tr><td colspan="7" class="p-4 text-center text-slate-400">No bookings recorded yet.</td></tr>';
        return;
      }
      bookings.forEach((b, index) => {
        tbody.innerHTML += `
          <tr class="hover:bg-slate-50">
            <td class="p-3 font-medium">${b.custName}<br><span class="text-xs text-slate-400">${b.custContact}</span></td>
            <td class="p-3"><span class="bg-indigo-100 text-indigo-800 text-xs px-2 py-1 rounded-full font-bold">Room ${b.roomNo}</span></td>
            <td class="p-3">${b.agentName}</td>
            <td class="p-3 text-xs">${b.date}</td>
            <td class="p-3 text-emerald-600 font-semibold">₹${b.advance}</td>
            <td class="p-3 text-amber-600 font-semibold">₹${b.due}</td>
            <td class="p-3 text-center">
              <button onclick="deleteBooking(${index})" class="text-rose-600 hover:text-rose-800"><i class="fa-solid fa-trash"></i></button>
            </td>
          </tr>
        `;
      });
    }

    function deleteBooking(index) {
      bookings.splice(index, 1);
      localStorage.setItem('app_bookings', JSON.stringify(bookings));
      renderBookings();
      updateDashboardStats();
    }

    // Render Calendar
    function renderCalendar() {
      const year = Number(calYearSelect.value) || 2026;
      const month = Number(document.getElementById('calMonthSelect').value) || 0;
      const grid = document.getElementById('calendar-grid');
      grid.innerHTML = '';

      const firstDay = new Date(year, month, 1).getDay();
      const daysInMonth = new Date(year, month + 1, 0).getDate();

      for (let i = 0; i < firstDay; i++) {
        grid.innerHTML += `<div class="p-3 border rounded-xl bg-slate-50"></div>`;
      }

      for (let day = 1; day <= daysInMonth; day++) {
        const dateStr = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
        const dayBookings = bookings.filter(b => b.date === dateStr);
        const isBooked = dayBookings.length > 0;

        grid.innerHTML += `
          <div class="p-2 border rounded-xl ${isBooked ? 'bg-indigo-50 border-indigo-300' : 'bg-white'} min-h-[70px] flex flex-col justify-between">
            <span class="font-bold text-sm ${isBooked ? 'text-indigo-600' : 'text-slate-700'}">${day}</span>
            ${isBooked ? `<span class="text-[10px] bg-indigo-600 text-white p-1 rounded font-semibold truncate">${dayBookings.length} Room(s)</span>` : ''}
          </div>
        `;
      }
    }

    // Dashboard Statistics
    function updateDashboardStats() {
      document.getElementById('stat-total-rooms').innerText = masterData.length;
      document.getElementById('stat-total-agents').innerText = new Set(masterData.map(m => m.agentName)).size;
      document.getElementById('stat-total-bookings').innerText = bookings.length;
      const totalDue = bookings.reduce((sum, b) => sum + (b.due || 0), 0);
      document.getElementById('stat-total-due').innerText = `₹${totalDue}`;
    }

    // Initial Setup
    renderMasterTable();
    renderBookings();
    updateDashboardStats();
  </script>
</body>
</html>
