<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Business Portal - Web Application</title>
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  <!-- SheetJS for Exporting to Excel -->
  <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
  <!-- FontAwesome Icons -->
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">
  <style>
    /* Compact Scrollbar */
    ::-webkit-scrollbar {
      width: 5px;
      height: 5px;
    }
    ::-webkit-scrollbar-track {
      background: #f1f5f9;
    }
    ::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 3px;
    }
    ::-webkit-scrollbar-thumb:hover {
      background: #94a3b8;
    }

    /* Print-specific Styles */
    @media print {
      body * {
        visibility: hidden;
      }
      #printable-invoice, #printable-invoice * {
        visibility: visible;
      }
      #printable-invoice {
        position: absolute;
        left: 0;
        top: 0;
        width: 100%;
        margin: 0;
        padding: 15px;
      }
      .no-print {
        display: none !important;
      }
    }

    /* Excel Comment Box Arrow */
    .excel-comment-box::before {
      content: '';
      position: absolute;
      top: -8px;
      left: 16px;
      border-width: 0 8px 8px 8px;
      border-style: solid;
      border-color: transparent transparent #0f172a transparent;
    }
  </style>
</head>
<body class="bg-slate-100 text-slate-800 font-sans min-h-screen flex flex-col relative antialiased text-xs" onclick="closeCommentBox()">

  <!-- LOGIN MODAL OVERLAY -->
  <div id="login-overlay" class="fixed inset-0 z-50 bg-slate-900/90 backdrop-blur-md flex items-center justify-center p-4">
    <div class="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-sm w-full p-6 space-y-4 text-left">
      <div class="text-center space-y-1">
        <div class="bg-indigo-100 text-indigo-700 w-12 h-12 rounded-full flex items-center justify-center mx-auto text-xl shadow-inner">
          <i class="fa-solid fa-lock"></i>
        </div>
        <h2 class="text-base font-bold text-slate-800">Homestay Business Portal</h2>
        <p class="text-[11px] text-slate-500">Please enter your credentials to access the system</p>
      </div>

      <form onsubmit="handleLogin(event)" class="space-y-3">
        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">User ID</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-2.5 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-user"></i>
            </span>
            <input type="text" id="login-userid" required placeholder="Enter User ID" class="w-full bg-slate-50 border border-slate-300 rounded-md pl-8 pr-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs">
          </div>
        </div>

        <div>
          <label class="block text-[11px] font-semibold text-slate-700 mb-1">Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-2.5 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="login-password" required placeholder="Enter Password" class="w-full bg-slate-50 border border-slate-300 rounded-md pl-8 pr-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-xs">
          </div>
        </div>

        <div id="login-error" class="hidden bg-rose-50 border border-rose-200 text-rose-600 text-[10px] p-2 rounded text-center font-medium">
          Invalid User ID or Password!
        </div>

        <button type="submit" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2 rounded-md shadow transition text-xs flex items-center justify-center gap-1.5">
          <i class="fa-solid fa-right-to-bracket"></i> Login
        </button>
      </form>
    </div>
  </div>

  <!-- MASTER DATA ACCESS PASSWORD MODAL -->
  <div id="master-auth-modal" class="hidden fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-xs w-full p-5 space-y-3 text-left">
      <div class="text-center space-y-1">
        <div class="bg-rose-100 text-rose-600 w-10 h-10 rounded-full flex items-center justify-center mx-auto text-lg shadow-inner">
          <i class="fa-solid fa-shield-halved"></i>
        </div>
        <h3 class="text-xs font-bold text-slate-800">Master Data Protected</h3>
        <p class="text-[10px] text-slate-500">Enter master password to access configuration and deletion tools.</p>
      </div>

      <form onsubmit="handleMasterAuth(event)" class="space-y-2.5">
        <div>
          <label class="block text-[10px] font-semibold text-slate-700 mb-1">Master Password</label>
          <div class="relative">
            <span class="absolute inset-y-0 left-0 pl-2.5 flex items-center text-slate-400 text-xs">
              <i class="fa-solid fa-key"></i>
            </span>
            <input type="password" id="master-password-input" required placeholder="Enter Master Password" class="w-full bg-slate-50 border border-slate-300 rounded-md pl-8 pr-3 py-1.5 focus:outline-none focus:ring-2 focus:ring-rose-500 text-xs">
          </div>
        </div>

        <div id="master-auth-error" class="hidden bg-rose-50 border border-rose-200 text-rose-600 text-[10px] p-1.5 rounded text-center font-medium">
          Incorrect Master Password!
        </div>

        <div class="flex space-x-2 pt-1">
          <button type="button" onclick="closeMasterAuthModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-1.5 rounded text-[11px] transition">
            Cancel
          </button>
          <button type="submit" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-1.5 rounded shadow transition text-[11px] flex items-center justify-center gap-1">
            <i class="fa-solid fa-unlock text-[10px]"></i> Unlock
          </button>
        </div>
      </form>
    </div>
  </div>

  <!-- MASTER DATA PERMANENT DELETION RECONFIRMATION POPUP MODAL -->
  <div id="master-delete-confirm-modal" class="hidden fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center p-4 no-print">
    <div class="bg-white rounded-xl shadow-2xl border border-rose-200 max-w-sm w-full p-5 space-y-3 text-center">
      <div class="bg-rose-100 text-rose-600 w-12 h-12 rounded-full flex items-center justify-center mx-auto text-xl shadow-inner">
        <i class="fa-solid fa-triangle-exclamation"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-800">Confirm Permanent Deletion</h3>
        <p id="master-delete-modal-msg" class="text-[11px] text-slate-600 mt-1">Are you sure you want to permanently delete this data from Master Tab? This action cannot be undone.</p>
      </div>
      <div class="flex space-x-2 pt-2">
        <button type="button" onclick="closeMasterDeleteModal()" class="w-1/2 bg-slate-100 hover:bg-slate-200 text-slate-700 font-semibold py-1.5 rounded text-[11px] transition">
          Cancel
        </button>
        <button type="button" onclick="confirmMasterDeletion()" class="w-1/2 bg-rose-600 hover:bg-rose-700 text-white font-bold py-1.5 rounded shadow transition text-[11px] flex items-center justify-center gap-1">
          <i class="fa-solid fa-trash-can text-[10px]"></i> Delete Permanently
        </button>
      </div>
    </div>
  </div>

  <!-- SESSION AUTO LOGOUT WARNING MODAL -->
  <div id="logout-warning-modal" class="hidden fixed inset-0 z-50 bg-slate-900/80 backdrop-blur-sm flex items-center justify-center p-4">
    <div class="bg-white rounded-lg shadow-xl border border-slate-200 max-w-xs w-full p-4 space-y-3 text-center">
      <div class="bg-amber-100 text-amber-600 w-10 h-10 rounded-full flex items-center justify-center mx-auto text-lg">
        <i class="fa-solid fa-hourglass-half"></i>
      </div>
      <div>
        <h3 class="text-xs font-bold text-slate-800">Inactivity Timeout Warning</h3>
        <p class="text-[10px] text-slate-500 mt-1">You will be logged out automatically in <strong id="logout-countdown-seconds" class="text-rose-600">60</strong> seconds due to inactivity.</p>
      </div>
      <button onclick="resetInactivityTimer()" class="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-semibold py-1.5 rounded text-[11px] transition shadow">
        Stay Logged In
      </button>
    </div>
  </div>

  <!-- Excel Comment Box Popout -->
  <div id="excel-comment-box" onclick="event.stopPropagation()" class="excel-comment-box hidden absolute z-50 bg-slate-900 text-white text-[11px] rounded-lg p-2.5 shadow-2xl border border-amber-400 space-y-2 w-64 transition-all duration-150">
    <div class="font-bold text-amber-300 border-b border-slate-700 pb-1 flex justify-between items-center text-[10px]">
      <span class="flex items-center gap-1">
        <i class="fa-solid fa-comment-dots text-amber-400"></i>
        <span id="comm-date-header">Date Overview</span>
      </span>
      <button onclick="closeCommentBox()" class="text-slate-400 hover:text-white px-1 py-0.5 rounded text-[10px]">
        <i class="fa-solid fa-xmark"></i>
      </button>
    </div>
    <div id="comm-booking-list" class="space-y-1.5 max-h-56 overflow-y-auto pr-0.5"></div>
  </div>

  <!-- Compact Header Navigation -->
  <header class="bg-indigo-700 text-white shadow sticky top-0 z-40 no-print">
    <div class="max-w-7xl mx-auto px-3 py-2 flex flex-col md:flex-row justify-between items-center gap-2">
      <div class="flex items-center space-x-2">
        <div class="bg-indigo-600 p-1.5 rounded-md border border-indigo-500 shadow-inner">
          <i class="fa-solid fa-hotel text-base text-indigo-100"></i>
        </div>
        <div>
          <h1 class="text-sm font-bold tracking-wide leading-none">Homestay Business Portal 🏠</h1>
          <p class="text-[10px] text-indigo-200">Management & Booking Control System</p>
        </div>
      </div>
      
      <!-- Compact Tab Navigation -->
      <nav class="flex space-x-1 bg-indigo-800/80 p-0.5 rounded-md text-[11px] font-medium border border-indigo-600/50">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-2.5 py-1 rounded transition-all active-tab">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-2.5 py-1 rounded transition-all text-indigo-100 hover:bg-indigo-600/50">Booking Details</button>
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-2.5 py-1 rounded transition-all text-indigo-100 hover:bg-indigo-600/50 flex items-center gap-1">
          <i class="fa-solid fa-lock text-[9px] text-amber-300"></i> Master Data
        </button>
        <button onclick="switchTab('calendar')" id="btn-calendar" class="tab-btn px-2.5 py-1 rounded transition-all text-indigo-100 hover:bg-indigo-600/50">Calendar</button>
      </nav>

      <!-- Action Buttons -->
      <div class="flex items-center space-x-1.5">
        <button onclick="openAlertModal()" title="View Alerts" class="relative bg-amber-500 hover:bg-amber-600 text-white px-2.5 py-1 rounded text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-bell text-[10px]"></i> Alerts
          <span id="alert-badge" class="hidden absolute -top-1.5 -right-1.5 bg-rose-600 text-white text-[9px] font-black px-1.5 py-0.2 rounded-full border border-white animate-bounce">0</span>
        </button>
        <button onclick="saveChanges()" class="bg-emerald-500 hover:bg-emerald-600 text-white px-2.5 py-1 rounded text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-floppy-disk text-[10px]"></i> Save changes
        </button>
        <button onclick="exportToExcel()" class="bg-indigo-600 hover:bg-indigo-800 text-white px-2.5 py-1 rounded text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-file-excel text-[10px]"></i> Export
        </button>
        <button onclick="logoutUser()" title="Logout" class="bg-rose-600 hover:bg-rose-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold flex items-center gap-1 transition">
          <i class="fa-solid fa-right-from-bracket text-[10px]"></i> Logout
        </button>
      </div>
    </div>
  </header>

  <!-- Notification Toast -->
  <div id="toast" class="hidden fixed bottom-4 right-4 bg-slate-900 text-white px-3 py-2 rounded-md shadow-lg z-50 flex items-center gap-2 no-print border border-slate-700 text-[11px]">
    <i class="fa-solid fa-circle-check text-emerald-400 text-sm"></i>
    <span id="toast-message" class="font-medium">Changes Auto save successfully!</span>
  </div>

  <!-- Main Content Area -->
  <main class="max-w-7xl mx-auto px-3 py-3 flex-1 w-full no-print space-y-3">

    <!-- DASHBOARD TAB -->
    <section id="tab-dashboard" class="tab-content space-y-3">
      <div class="bg-gradient-to-r from-indigo-700 via-indigo-600 to-blue-600 rounded-lg p-3 text-white shadow-sm flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2">
        <div>
          <h2 class="text-base font-bold tracking-tight">Hi Aniruddha, Welcome to dashboard 🏠</h2>
          <p class="text-indigo-100 text-[10px] mt-0.5">Quickly view, schedule, and manage room allocations and orders.</p>
        </div>
        <div class="flex items-center bg-indigo-900/60 p-1.5 rounded-lg border border-indigo-400/40 space-x-2">
          <label for="dash-year-select" class="text-[10px] font-bold text-indigo-100 uppercase flex items-center gap-1">
            <i class="fa-solid fa-filter text-amber-300"></i> Filter Year:
          </label>
          <select id="dash-year-select" onchange="handleDashboardYearChange(this.value)" class="bg-white text-indigo-900 text-[11px] font-bold rounded px-2 py-1 focus:outline-none focus:ring-2 focus:ring-amber-400 cursor-pointer shadow">
          </select>
        </div>
      </div>

      <!-- Summary Filter Banner Indicator -->
      <div class="flex items-center justify-between bg-white px-3 py-1.5 rounded-md border border-slate-200 shadow-xs">
        <span class="text-[11px] font-semibold text-slate-600 flex items-center gap-1.5">
          <i class="fa-solid fa-chart-line text-indigo-600"></i>
          Showing Summary For: <strong id="dash-filter-label" class="text-indigo-700 font-bold">Consolidated (All Years)</strong>
        </span>
        <button onclick="handleDashboardYearChange('CURRENT')" class="text-[10px] bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold px-2 py-0.5 rounded transition border border-slate-300">
          Reset to Current Year
        </button>
      </div>

      <!-- Compact Summary Cards -->
      <div class="grid grid-cols-2 lg:grid-cols-4 gap-2.5">
        <div class="bg-white p-2.5 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Total Bookings</p>
            <p id="dash-total-bookings" class="text-lg font-black text-slate-800">0</p>
          </div>
          <div class="p-2 bg-blue-50 text-blue-600 rounded-md"><i class="fa-solid fa-bookmark text-sm"></i></div>
        </div>
        <div class="bg-white p-2.5 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Booking Amount</p>
            <p id="dash-total-amount" class="text-lg font-black text-slate-800">₹0</p>
          </div>
          <div class="p-2 bg-indigo-50 text-indigo-600 rounded-md"><i class="fa-solid fa-receipt text-sm"></i></div>
        </div>
        <div class="bg-white p-2.5 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Advance Received</p>
            <p id="dash-advanced" class="text-lg font-black text-emerald-600">₹0</p>
          </div>
          <div class="p-2 bg-emerald-50 text-emerald-600 rounded-md"><i class="fa-solid fa-wallet text-sm"></i></div>
        </div>
        <div class="bg-white p-2.5 rounded-lg shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[9px] uppercase font-bold text-slate-400 tracking-wider">Total Due Amount</p>
            <p id="dash-due" class="text-lg font-black text-rose-600">₹0</p>
          </div>
          <div class="p-2 bg-rose-50 text-rose-600 rounded-md"><i class="fa-solid fa-hand-holding-dollar text-sm"></i></div>
        </div>
      </div>

      <!-- Active years Directory Table Hidden -->
      <div class="hidden bg-white rounded-lg shadow-sm border border-slate-200 p-3">
        <div class="mb-2 flex justify-between items-center">
          <h3 class="text-xs font-bold text-slate-800 flex items-center gap-1">
            <i class="fa-solid fa-calendar-days text-indigo-600"></i> Active Years Directory (2026 – 2085)
          </h3>
          <span class="text-[10px] text-slate-400 font-medium">Click any year to filter dashboard & open year calendar</span>
        </div>
        <div id="years-grid" class="grid grid-cols-6 sm:grid-cols-10 md:grid-cols-12 gap-1.5"></div>
      </div>
    </section>

    <!-- BOOKING DETAILS TAB -->
    <section id="tab-booking" class="tab-content hidden space-y-3">
      <div class="bg-white rounded-lg shadow-sm border border-slate-200 p-3">
        <div class="flex flex-col md:flex-row justify-between items-start md:items-center gap-2 mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-800 flex items-center gap-1">
              <i class="fa-solid fa-address-card text-indigo-600"></i> Guest Information & Reservation Directory
            </h2>
            <div class="flex items-center gap-3 mt-1.5 text-[10px]">
              <span class="flex items-center gap-1 font-semibold text-amber-800">
                <span class="relative flex h-2.5 w-2.5">
                  <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
                  <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
                </span> Live Booking
              </span>
              <span class="flex items-center gap-1 font-semibold text-blue-800">
                <span class="w-2.5 h-2.5 bg-blue-500 rounded-full inline-block"></span> Upcoming Booking
              </span>
              <span class="flex items-center gap-1 font-semibold text-emerald-800">
                <span class="w-2.5 h-2.5 bg-emerald-500 rounded-full inline-block"></span> Closed Booking
              </span>
              <span class="flex items-center gap-1 font-semibold text-slate-700">
                <span class="w-2 h-2 bg-rose-600 rounded-full inline-block"></span> Inactive Booking
              </span>
            </div>
          </div>
          
          <div class="flex items-center space-x-2 w-full md:w-auto">
            <!-- Search by Date -->
            <div class="flex items-center bg-slate-50 border border-slate-300 rounded p-1 space-x-1.5">
              <label for="booking-date-search" class="text-[10px] font-bold text-slate-500 uppercase flex items-center gap-1 pl-1">
                <i class="fa-solid fa-calendar-day text-indigo-600"></i> Search Date:
              </label>
              <input type="date" id="booking-date-search" onchange="searchBookingByDate()" class="bg-white text-[11px] border border-slate-300 rounded px-2 py-0.5 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-bold text-indigo-700 cursor-pointer">
              <button onclick="clearDateSearchBooking()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Reset Filter">
                <i class="fa-solid fa-rotate-left"></i> Reset
              </button>
            </div>

            <button onclick="openBookingModal()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded text-[11px] font-semibold flex items-center gap-1 transition whitespace-nowrap">
              <i class="fa-solid fa-plus text-[10px]"></i> Add Booking
            </button>
          </div>
        </div>

        <!-- Bookings Table View -->
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                <th class="py-2 px-2">Booking ID</th>
                <th class="py-2 px-2">Guest Name</th>
                <th class="py-2 px-2">Contact No</th>
                <th class="py-2 px-2">ID No</th>
                <th class="py-2 px-2">Attached ID</th>
                <th class="py-2 px-2">Room</th>
                <th class="py-2 px-2">Capacity</th>
                <th class="py-2 px-2">Agent Info</th>
                <th class="py-2 px-2 min-w-[150px]">Stay Window</th>
                <th class="py-2 px-2">Tariff & Extras</th>
                <th class="py-2 px-2">Payment/Adv</th>
                <th class="py-2 px-2">Due</th>
                <th class="py-2 px-2 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="bookings-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- MASTER DATA TAB -->
    <section id="tab-master" class="tab-content hidden space-y-3">
      
      <!-- Room Capacity Table (Defaults to 5 Rooms: 1, 2, 3, 4, 5) -->
      <div class="bg-white rounded-lg shadow-sm border border-slate-200 p-3">
        <div class="flex justify-between items-center mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-800 flex items-center gap-1">
              <i class="fa-solid fa-door-open text-indigo-600"></i> Room Capacity Configuration
            </h2>
            <p class="text-[10px] text-slate-500">Default rooms 1 to 5. Click Add Room to append new rooms anytime.</p>
          </div>
          <button type="button" onclick="addRoomCapacityRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded text-[11px] font-medium flex items-center gap-1 transition shadow-sm cursor-pointer">
            <i class="fa-solid fa-plus text-[10px]"></i> Add Room
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                <th class="py-2 px-2">Room No</th>
                <th class="py-2 px-2">Room Capacity (Person)</th>
                <th class="py-2 px-2 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="room-capacity-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>

      <!-- Agent Information Table -->
      <div class="bg-white rounded-lg shadow-sm border border-slate-200 p-3">
        <div class="flex justify-between items-center mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-800 flex items-center gap-1">
              <i class="fa-solid fa-users-gear text-indigo-600"></i> Master Agent Directory
            </h2>
            <p class="text-[10px] text-slate-500">Manage Agents linked with room allocations.</p>
          </div>
          <button type="button" onclick="addAgentRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded text-[11px] font-medium flex items-center gap-1 transition shadow-sm cursor-pointer">
            <i class="fa-solid fa-plus text-[10px]"></i> Add Agent Entry
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                <th class="py-2 px-2">Agent Name</th>
                <th class="py-2 px-2">Agent Contact</th>
                <th class="py-2 px-2">Linked Room No</th>
                <th class="py-2 px-2 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="agent-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
          </table>
        </div>
      </div>

      <!-- BOOKING ID TYPING SEARCH & DELETION CONTROL -->
      <div class="bg-white rounded-lg shadow-sm border border-rose-200/80 p-3 space-y-2.5">
        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-2 border-b border-slate-100 pb-2">
          <div>
            <h2 class="text-xs font-bold text-slate-800 flex items-center gap-1">
              <i class="fa-solid fa-trash-can text-rose-600"></i>Booking Deletion Manager
            </h2>
            <p class="text-[10px] text-slate-500">Type a Booking ID directly to safely locate and remove it from the system.</p>
          </div>
          
          <div class="flex items-center bg-slate-50 border border-slate-300 rounded p-1 space-x-1.5">
            <label for="master-booking-search-input" class="text-[10px] font-bold text-slate-600 uppercase flex items-center gap-1 pl-1">
              <i class="fa-solid fa-magnifying-glass text-indigo-600"></i> Type Booking ID:
            </label>
            <input type="text" id="master-booking-search-input" oninput="searchMasterBookingById()" placeholder="e.g. BKG-2026-0000001" class="bg-white text-[11px] border border-slate-300 rounded px-2 py-0.5 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-mono font-bold text-indigo-700 uppercase w-48">
            <button onclick="clearMasterBookingSearch()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Clear Search">
              <i class="fa-solid fa-xmark"></i>
            </button>
          </div>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-rose-50/60 border-b border-rose-100 text-[10px] font-bold text-rose-800 uppercase tracking-wider">
                <th class="py-2 px-2">Booking ID</th>
                <th class="py-2 px-2">Guest Name</th>
                <th class="py-2 px-2">Room No</th>
                <th class="py-2 px-2">Stay Window</th>
                <th class="py-2 px-2">Total Amount</th>
                <th class="py-2 px-2">Due Amount</th>
                <th class="py-2 px-2 text-center">Delete Linked Booking</th>
              </tr>
            </thead>
            <tbody id="master-delete-tbody" class="divide-y divide-slate-100 text-[11px]">
              <tr>
                <td colspan="7" class="text-center py-4 text-slate-400">Please type a Booking ID into the search field above to view and delete details.</td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- CALENDAR TAB -->
    <section id="tab-calendar" class="tab-content hidden space-y-3">
      <div class="bg-white rounded-lg shadow-sm border border-slate-200 p-3">
        <div class="flex justify-between items-center mb-3">
          <div>
            <h2 class="text-xs font-bold text-slate-800 flex items-center gap-1">
              <i class="fa-regular fa-calendar-check text-indigo-600"></i> Year Overview Calendar
            </h2>
          </div>
          
          <div class="flex items-center bg-indigo-900/60 p-1.5 rounded-lg border border-indigo-400/40 space-x-2">
            <label for="cal-year-select" class="text-[10px] font-bold text-indigo-100 uppercase flex items-center gap-1">
              <i class="fa-solid fa-filter text-amber-300"></i> Filter Year:
            </label>
            <select id="cal-year-select" onchange="renderCalendar(parseInt(this.value))" class="bg-white text-indigo-900 text-[11px] font-bold rounded px-2 py-1 focus:outline-none focus:ring-2 focus:ring-amber-400 cursor-pointer shadow"></select>
          </div>
        </div>

        <div id="calendar-container" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-2.5"></div>
      </div>
    </section>

  </main>

  <!-- POPUP MODAL: CHECK-OUT ALERT LIST -->
  <div id="alert-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-3 no-print">
    <div class="bg-white rounded-lg shadow-xl border border-slate-200 max-w-lg w-full flex flex-col max-h-[85vh] overflow-hidden">
      <div class="bg-amber-500 p-3 text-white flex justify-between items-center">
        <div class="flex items-center space-x-2">
          <i class="fa-solid fa-bell text-base"></i>
          <h3 class="text-xs font-bold">Check-out Alert</h3>
        </div>
        <button onclick="closeAlertModal()" class="text-amber-100 hover:text-white px-1 text-base">
          <i class="fa-solid fa-xmark"></i>
        </button>
      </div>

      <div id="alert-list-container" class="p-3 overflow-y-auto space-y-2 flex-1 text-[11px]"></div>

      <div class="bg-slate-50 border-t border-slate-200 p-2.5 flex justify-between items-center text-[11px]">
        <span id="alert-list-count-text" class="text-slate-500 font-medium">0 active warnings found</span>
        <button onclick="closeAlertModal()" class="px-3 py-1 bg-slate-800 text-white rounded font-semibold text-[10px]">Dismiss</button>
      </div>
    </div>
  </div>

  <!-- COMPACT ADD / EDIT BOOKING MODAL -->
  <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-2 overflow-y-auto no-print">
    <div class="bg-white rounded-lg shadow-xl border border-slate-200 max-w-xl w-full p-4 space-y-2.5 my-4">
      <div class="flex justify-between items-center pb-2 border-b border-slate-100">
        <div>
          <h3 id="modal-title" class="text-xs font-bold text-slate-800 flex items-center gap-1">
            <i class="fa-solid fa-calendar-plus text-indigo-600"></i> Add New Booking
          </h3>
        </div>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 p-0.5 text-base"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="booking-form" onsubmit="handleSaveBooking(event)" class="space-y-2.5 text-[11px]">
        <input type="hidden" id="modal-booking-id">

        <!-- GUEST DETAILS -->
        <div class="bg-slate-50 p-2.5 rounded-md border border-slate-200 space-y-2">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1">
            <i class="fa-solid fa-user-tag text-indigo-500"></i> Guest Information
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Guest Name <span class="text-rose-500">*</span></label>
              <input type="text" id="cust-name" required pattern="[A-Za-z\s]+" oninput="this.value = formatTitleCase(this.value.replace(/[^A-Za-z\s]/g, ''))" title="Please enter Guest Name using characters only" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Address</label>
              <input type="text" id="cust-address" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">City</label>
              <input type="text" id="cust-city" placeholder="City" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">State</label>
              <input type="text" id="cust-state" oninput="this.value = formatTitleCase(this.value); handleStateChange(this.value)" placeholder="State" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Country</label>
              <input type="text" id="cust-country" placeholder="Country" oninput="this.value = formatTitleCase(this.value)" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Pin/Zip Code</label>
              <input type="text" id="cust-zip" placeholder="Pin/Zip Code" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">ID Number</label>
              <input type="text" id="cust-id" maxlength="16" pattern="[A-Za-z0-9\s]*" oninput="this.value = this.value.replace(/[^A-Za-z0-9\s]/g, '')" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Contact No</label>
              <input type="text" id="cust-contact" maxlength="10" pattern="[0-9]*" oninput="this.value = this.value.replace(/[^0-9]/g, '')" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div class="sm:col-span-2">
              <label class="block font-semibold text-slate-600 mb-0.5 flex justify-between items-center">
                <span>Attached ID Proof <span class="text-[9px] text-indigo-500 font-normal">(PDF, 10KB - 900KB)</span></span>
                <button type="button" id="cust-id-file-remove" onclick="removeAttachedIdProof()" class="hidden text-rose-500 hover:text-rose-700 text-[9px] font-bold">Remove</button>
              </label>
              <div class="flex items-center gap-1.5">
                <input type="file" id="cust-id-file" accept="application/pdf" onchange="handleIdProofUpload(event)" class="w-full text-[10px] text-slate-500 file:mr-2 file:py-1 file:px-2 file:rounded file:border-0 file:text-[10px] file:font-semibold file:bg-indigo-50 file:text-indigo-700 hover:file:bg-indigo-100 cursor-pointer bg-white border border-slate-300 rounded py-0.5">
                <input type="hidden" id="cust-id-file-base64">
                <input type="hidden" id="cust-id-file-name">
              </div>
              <p id="cust-id-file-status" class="text-[9px] text-slate-400 mt-0.5 italic">No PDF document attached.</p>
            </div>
          </div>
        </div>

        <!-- Room & Stay Schedule Box -->
        <div class="bg-slate-50 p-2.5 rounded-md border border-slate-200 space-y-2">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1">
            <i class="fa-solid fa-bed text-indigo-500"></i> Room Selection & Stay Dates
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-3 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Room No</label>
              <select id="cust-room" onchange="autoCaptureRoomDetails()" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-bold text-indigo-700"></select>
            </div>
            
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Agent Info</label>
              <select id="cust-agent" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-bold text-slate-700"></select>
            </div>

            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Capacity (Person)</label>
              <input type="number" id="cust-capacity" min="1" value="1" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-bold text-slate-700">
            </div>

            <div class="sm:col-span-3 grid grid-cols-2 gap-2 pt-1 border-t border-slate-200/60">
              <div>
                <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-solid fa-plane-arrival text-emerald-600 mr-1"></i> Check In</label>
                <div class="flex gap-1">
                  <input type="date" id="cust-checkin-date" onchange="handleStayDatesChange()" required class="w-2/3 bg-white border border-slate-300 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-medium">
                  <input type="time" id="cust-checkin-time" value="12:00" onchange="handleStayDatesChange()" required class="w-1/3 bg-white border border-slate-300 rounded px-1 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-medium">
                </div>
              </div>

              <div>
                <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-solid fa-plane-departure text-rose-500 mr-1"></i> Check Out</label>
                <div class="flex gap-1">
                  <input type="date" id="cust-checkout-date" onchange="handleStayDatesChange()" required class="w-2/3 bg-white border border-slate-300 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-medium">
                  <input type="time" id="cust-checkout-time" value="11:00" onchange="handleStayDatesChange()" required class="w-1/3 bg-white border border-slate-300 rounded px-1 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-medium">
                </div>
              </div>
            </div>
          </div>
        </div>

        <!-- EXTRA FOOD SECTION WITH DATE & TIME -->
        <div class="bg-amber-50/70 p-2.5 rounded-md border border-amber-200 space-y-2">
          <div class="flex justify-between items-center">
            <h4 class="text-[9px] font-bold uppercase tracking-wider text-amber-800 flex items-center gap-1">
              <i class="fa-solid fa-utensils text-amber-600"></i> Extra Food / Drink Orders List
            </h4>
            <button type="button" id="btn-add-food-order" onclick="addFoodOrderItem()" class="bg-amber-600 hover:bg-amber-700 text-white px-2 py-0.5 rounded text-[10px] font-semibold flex items-center gap-1 transition">
              <i class="fa-solid fa-plus text-[9px]"></i> Add Food Order
            </button>
          </div>
          
          <div id="food-orders-container" class="space-y-2 max-h-40 overflow-y-auto pr-1"></div>
        </div>

        <!-- Billing Calculation Box -->
        <div class="bg-indigo-50/50 p-2.5 rounded-md border border-indigo-100 space-y-2">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-indigo-700 flex items-center gap-1">
            <i class="fa-solid fa-calculator text-indigo-600"></i> Billing Summary
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-6 gap-1.5">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Days</label>
              <input type="number" id="cust-days" readonly class="w-full bg-slate-200/80 font-bold text-slate-700 border border-slate-300 rounded px-1.5 py-1 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Price/Day (₹)</label>
              <input type="number" id="cust-price" value="1200" oninput="calculateModalBilling()" class="w-full bg-white font-bold text-slate-700 border border-slate-300 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Total (₹)</label>
              <input type="number" id="cust-total" readonly class="w-full bg-slate-200/80 text-indigo-700 font-bold border border-slate-300 rounded px-1.5 py-1 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Advance (₹)</label>
              <input type="number" id="cust-advance" value="0" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-300 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-semibold text-emerald-600">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Due (₹)</label>
              <input type="number" id="cust-due" readonly class="w-full bg-slate-200/80 text-rose-700 font-bold border border-slate-300 rounded px-1.5 py-1 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5 text-[10px] text-emerald-700">Clear Bill (₹)</label>
              <input type="number" id="cust-clear-bill" value="0" placeholder="0" oninput="handleClearBillPayment(this.value)" class="w-full bg-emerald-50 border border-emerald-400 font-bold text-emerald-800 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-emerald-500" title="Put payment amount to clear due bill">
            </div>
          </div>
        </div>

        <div class="flex justify-end space-x-2 pt-1">
          <button type="button" onclick="closeBookingModal()" class="px-3 py-1 bg-slate-100 text-slate-700 rounded font-semibold transition">Cancel</button>
          <button type="submit" id="btn-save-booking" class="px-4 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded font-semibold shadow transition">Save Booking</button>
        </div>
      </form>
    </div>
  </div>

  <!-- PRINTABLE INVOICE / BOOKING RECEIPT MODAL -->
  <div id="invoice-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-3 overflow-y-auto">
    <div class="bg-white rounded-lg shadow-xl border border-slate-200 max-w-xl w-full p-6 space-y-4 relative" id="printable-invoice">
      
      <!-- Read-Only Notice Bar (Shown for Closed/Inactive Bookings) -->
      <div id="inv-readonly-notice" class="hidden bg-slate-800 text-amber-300 text-[10px] font-bold px-3 py-1.5 rounded-md flex items-center justify-between border border-amber-400/40">
        <span class="flex items-center gap-1.5">
          <i class="fa-solid fa-lock text-amber-400"></i> Read-Only View Mode (Editing Disabled)
        </span>
        <span class="text-[9px] text-slate-300 font-normal">System Protected</span>
      </div>

      <div class="flex justify-between items-start border-b border-slate-200 pb-3">
        <div>
          <h2 class="text-lg font-black text-indigo-700 uppercase tracking-wide">Aniruddha Homestay</h2>
          <p class="text-[10px] text-slate-500 mt-0.5">Sittong, Village in West Bengal</p>
          <p class="text-[10px] text-slate-500">Phone: +91 9804396541 | Email: info@businessportal.com</p>
        </div>
        <div class="text-right">
          <span id="inv-badge" class="inline-block bg-indigo-100 text-indigo-800 text-[9px] font-bold px-2 py-0.5 rounded-full uppercase mb-0.5">e-Invoice</span>
          <p id="inv-id-container" class="text-[10px] text-slate-500">Invoice ID: <strong id="inv-id" class="text-slate-800 font-mono">INV-2026-0000001</strong></p>
          <p class="text-[10px] text-slate-500">Booking ID: <strong id="inv-booking-id" class="text-indigo-700 font-mono">BKG-2026-0000001</strong></p>
          <p class="text-[10px] text-slate-500">Issued On: <strong id="inv-date" class="text-slate-800"></strong></p>
        </div>
      </div>

      <div class="grid grid-cols-2 gap-4 bg-slate-50 p-3 rounded-lg border border-slate-100 text-[11px]">
        <div>
          <h4 class="font-bold text-slate-400 uppercase text-[9px] tracking-wider mb-0.5">Guest Information</h4>
          <p class="text-slate-800 font-semibold" id="inv-guest-name">-</p>
          <p class="text-slate-600" id="inv-guest-address">Address: -</p>
          <p class="text-slate-600" id="inv-guest-contact">Contact: -</p>
          <p class="text-slate-600" id="inv-guest-id">ID No: -</p>
        </div>
        <div>
          <h4 class="font-bold text-slate-400 uppercase text-[9px] tracking-wider mb-0.5">Reservation Info</h4>
          <p class="text-slate-800 font-semibold" id="inv-room">Room No: -</p>
          <p class="text-slate-600" id="inv-checkin">Check-in: -</p>
          <p class="text-slate-600" id="inv-checkout">Check-out: -</p>
        </div>
      </div>

      <div class="overflow-x-auto">
        <table class="w-full text-left text-[11px]">
          <thead>
            <tr class="bg-indigo-50 text-indigo-900 border-b border-indigo-100">
              <th class="p-2">Description</th>
              <th class="p-2 text-center">Qty / Duration</th>
              <th class="p-2 text-right">Rate / Day</th>
              <th class="p-2 text-right">Total Amount</th>
            </tr>
          </thead>
          <tbody id="inv-items-tbody" class="divide-y divide-slate-100"></tbody>
        </table>
      </div>

      <div class="flex justify-end pt-1 border-t border-slate-200">
        <div class="w-1/2 space-y-1 text-[11px]">
          <div class="flex justify-between text-slate-600">
            <span>Total Amount:</span>
            <strong id="inv-sum-total" class="text-slate-800">₹0</strong>
          </div>
          <div class="flex justify-between text-emerald-600">
            <span>Advance Payment:</span>
            <strong id="inv-sum-advance">₹0</strong>
          </div>
          <div class="flex justify-between text-rose-600 font-bold border-t border-slate-200 pt-1">
            <span>Balance Due:</span>
            <span id="inv-sum-due">₹0</span>
          </div>
          <!-- DYNAMICALLY DISPLAYED CLEAR DUE ROW -->
          <div id="inv-clear-due-row" class="hidden flex justify-between text-emerald-700 font-bold border-t border-slate-100 pt-1">
            <span>Clear Due:</span>
            <strong id="inv-sum-clear-due">₹0</strong>
          </div>
        </div>
      </div>

      <div class="pt-4 border-t border-slate-200 flex justify-between items-end text-[10px] text-slate-400">
        <div>
          <p class="font-bold text-slate-600">Thank you for staying with us!</p>
          <p>For inquiries, please contact hotel management.</p>
        </div>
        <div class="text-center border-t border-slate-300 pt-1 w-28">
          <p class="font-semibold text-slate-600">Authorized Signature</p>
        </div>
      </div>

      <div class="flex justify-end space-x-2 pt-2 no-print border-t border-slate-100">
        <button type="button" onclick="closeInvoiceModal()" class="px-3 py-1 bg-slate-100 text-slate-700 rounded font-semibold transition">Close</button>
        <button type="button" onclick="window.print()" id="inv-print-btn" class="px-4 py-1 bg-indigo-600 text-white rounded font-semibold shadow flex items-center gap-1 transition">
          <i class="fa-solid fa-print"></i> Print Invoice
        </button>
      </div>
    </div>
  </div>

  <script>
    window.addEventListener('beforeunload', function (e) {
      if (isLoggedIn) {
        e.preventDefault();
        e.returnValue = 'Please click "Save Changes" button to save the history.'; 
        return e.returnValue;
      }
    });

    function formatTitleCase(text) {
      if (!text) return '';
      return text.replace(/\w\S*/g, function(txt) {
        return txt.charAt(0).toUpperCase() + txt.substr(1).toLowerCase();
      });
    }

    function handleStateChange(stateValue) {
      if (stateValue && stateValue.trim().toLowerCase() === 'west bengal') {
        const countryInput = document.getElementById('cust-country');
        if (countryInput) countryInput.value = 'India';
      }
    }

    function handleIdProofUpload(e) {
      const fileInput = e.target;
      const file = fileInput.files[0];
      const statusText = document.getElementById('cust-id-file-status');
      const base64Input = document.getElementById('cust-id-file-base64');
      const fileNameInput = document.getElementById('cust-id-file-name');
      const removeBtn = document.getElementById('cust-id-file-remove');

      if (!file) return;

      if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith('.pdf')) {
        alert("⚠️ Invalid file format! Only PDF files are allowed.");
        fileInput.value = '';
        return;
      }

      const minSize = 10 * 1024;  
      const maxSize = 900 * 1024; 

      if (file.size < minSize || file.size > maxSize) {
        const fileSizeKB = (file.size / 1024).toFixed(1);
        alert(`⚠️ Invalid file size (${fileSizeKB} KB)!\n\nThe attached ID proof PDF must be between 10 KB and 900 KB.`);
        fileInput.value = '';
        return;
      }

      const reader = new FileReader();
      reader.onload = function(evt) {
        base64Input.value = evt.target.result;
        fileNameInput.value = file.name;
        statusText.innerHTML = `<span class="text-emerald-600 font-semibold"><i class="fa-solid fa-circle-check"></i> Attached: ${file.name} (${(file.size / 1024).toFixed(1)} KB)</span>`;
        removeBtn.classList.remove('hidden');
      };
      reader.readAsDataURL(file);
    }

    function removeAttachedIdProof() {
      document.getElementById('cust-id-file').value = '';
      document.getElementById('cust-id-file-base64').value = '';
      document.getElementById('cust-id-file-name').value = '';
      document.getElementById('cust-id-file-status').innerText = 'No PDF document attached.';
      document.getElementById('cust-id-file-remove').classList.add('hidden');
    }

    function openPdfAttachment(base64Data) {
      if (!base64Data) {
        alert("No ID Proof attached!");
        return;
      }
      const win = window.open();
      if (win) {
        win.document.write(`<iframe src="${base64Data}" frameborder="0" style="border:0; top:0px; left:0px; bottom:0px; right:0px; width:100%; height:100%;" allowfullscreen></iframe>`);
      } else {
        alert("Please allow popups to view attached PDF document.");
      }
    }

    let isLoggedIn = false;
    let isMasterUnlocked = false; 
    let inactivityTimer = null;
    let warningTimer = null;
    let countdownInterval = null;
    const INACTIVITY_LIMIT_MS = 10 * 60 * 1000; 
    const WARNING_BUFFER_MS = 1 * 60 * 1000;   

    const DEFAULT_USER_ID = "Admin";
    const DEFAULT_PASSWORD = "Aadmin123";

    let pendingMasterDeleteType = null; 
    let pendingMasterDeleteTarget = null; 

    function openMasterDeleteModal(type, target) {
      pendingMasterDeleteType = type;
      pendingMasterDeleteTarget = target;

      const msgElem = document.getElementById('master-delete-modal-msg');
      if (type === 'booking') {
        const b = state.bookings.find(item => item.id === target);
        const bCode = b ? b.bookingCode : 'this booking';
        msgElem.innerText = `Are you sure you want to permanently delete booking ${bCode} from the Master Tab? This action cannot be undone.`;
      } else if (type === 'room') {
        msgElem.innerText = `Are you sure you want to permanently delete this Room Capacity record? This action cannot be undone.`;
      } else if (type === 'agent') {
        msgElem.innerText = `Are you sure you want to permanently delete this Agent record? This action cannot be undone.`;
      }

      document.getElementById('master-delete-confirm-modal').classList.remove('hidden');
    }

    function closeMasterDeleteModal() {
      pendingMasterDeleteType = null;
      pendingMasterDeleteTarget = null;
      document.getElementById('master-delete-confirm-modal').classList.add('hidden');
    }

    function confirmMasterDeletion() {
      if (pendingMasterDeleteType === 'booking') {
        const id = pendingMasterDeleteTarget;
        const idx = state.bookings.findIndex(b => b.id === id);
        if (idx !== -1) {
          state.bookings[idx].inactive = true;
        }
        searchMasterBookingById();
        renderBookingsTable();
        updateDashboardCards();
        renderCalendar(defaultAppYear);
        checkUpcomingCheckoutsWithDue();
        saveChanges(false, false);
      } else if (pendingMasterDeleteType === 'room') {
        const index = pendingMasterDeleteTarget;
        state.roomsCapacity.splice(index, 1);
        renderRoomCapacityTable();
        populateRoomDropdown();
        populateAgentDropdown();
        saveChanges(false, false);
      } else if (pendingMasterDeleteType === 'agent') {
        const index = pendingMasterDeleteTarget;
        state.masterAgents.splice(index, 1);
        renderMasterAgentTable();
        populateAgentDropdown();
        saveChanges(false, false);
      }
      closeMasterDeleteModal();
    }

    function checkAuthStatus() {
      const sessionAuth = sessionStorage.getItem('app_authenticated');
      if (sessionAuth === 'true') {
        isLoggedIn = true;
        document.getElementById('login-overlay').classList.add('hidden');
        startInactivityMonitoring();
      } else {
        isLoggedIn = false;
        document.getElementById('login-overlay').classList.remove('hidden');
      }
    }

    function handleLogin(e) {
      e.preventDefault();
      const user = document.getElementById('login-userid').value.trim();
      const pass = document.getElementById('login-password').value.trim();

      if (user === DEFAULT_USER_ID && pass === DEFAULT_PASSWORD) {
        isLoggedIn = true;
        sessionStorage.setItem('app_authenticated', 'true');
        document.getElementById('login-overlay').classList.add('hidden');
        document.getElementById('login-error').classList.remove('hidden');
        startInactivityMonitoring();
      } else {
        document.getElementById('login-error').classList.remove('hidden');
      }
    }

    function logoutUser() {
      isLoggedIn = false;
      isMasterUnlocked = false;
      sessionStorage.removeItem('app_authenticated');
      stopInactivityMonitoring();
      document.getElementById('logout-warning-modal').classList.add('hidden');
      document.getElementById('login-password').value = '';
      document.getElementById('login-overlay').classList.remove('hidden');
    }

    function openMasterAuthModal() {
      document.getElementById('master-password-input').value = '';
      document.getElementById('master-auth-error').classList.add('hidden');
      document.getElementById('master-auth-modal').classList.remove('hidden');
    }

    function closeMasterAuthModal() {
      document.getElementById('master-auth-modal').classList.add('hidden');
    }

    function handleMasterAuth(e) {
      e.preventDefault();
      const enteredPass = document.getElementById('master-password-input').value.trim();

      if (enteredPass === DEFAULT_PASSWORD) {
        isMasterUnlocked = true;
        closeMasterAuthModal();
        performSwitchTab('master');
      } else {
        document.getElementById('master-auth-error').classList.remove('hidden');
      }
    }

    function startInactivityMonitoring() {
      stopInactivityMonitoring();
      
      const activityEvents = ['mousemove', 'keydown', 'mousedown', 'touchstart', 'scroll'];
      activityEvents.forEach(evt => {
        window.addEventListener(evt, resetInactivityTimer);
      });

      resetInactivityTimer();
    }

    function stopInactivityMonitoring() {
      if (inactivityTimer) clearTimeout(inactivityTimer);
      if (warningTimer) clearTimeout(warningTimer);
      if (countdownInterval) clearInterval(countdownInterval);
      
      const activityEvents = ['mousemove', 'keydown', 'mousedown', 'touchstart', 'scroll'];
      activityEvents.forEach(evt => {
        window.removeEventListener(evt, resetInactivityTimer);
      });
    }

    function resetInactivityTimer() {
      if (!isLoggedIn) return;

      if (inactivityTimer) clearTimeout(inactivityTimer);
      if (warningTimer) clearTimeout(warningTimer);
      if (countdownInterval) clearInterval(countdownInterval);

      document.getElementById('logout-warning-modal').classList.add('hidden');

      warningTimer = setTimeout(showInactivityWarning, INACTIVITY_LIMIT_MS - WARNING_BUFFER_MS);
      inactivityTimer = setTimeout(logoutUser, INACTIVITY_LIMIT_MS);
    }

    function showInactivityWarning() {
      if (!isLoggedIn) return;

      let secondsLeft = 60;
      document.getElementById('logout-countdown-seconds').innerText = secondsLeft;
      document.getElementById('logout-warning-modal').classList.remove('hidden');

      countdownInterval = setInterval(() => {
        secondsLeft--;
        if (secondsLeft >= 0) {
          document.getElementById('logout-countdown-seconds').innerText = secondsLeft;
        } else {
          clearInterval(countdownInterval);
        }
      }, 1000);
    }

    function formatDateTime(dtStr) {
      if (!dtStr) return '-';
      const d = new Date(dtStr);
      if (isNaN(d.getTime())) {
        const parts = dtStr.split('T');
        if (parts.length === 2) {
          const dateParts = parts[0].split('-');
          if (dateParts.length === 3) {
            return `${dateParts[2]}-${dateParts[1]}-${dateParts[0]} ${parts[1]}`;
          }
        }
        return dtStr.replace('T', ' ');
      }
      const day = String(d.getDate()).padStart(2, '0');
      const month = String(d.getMonth() + 1).padStart(2, '0');
      const year = d.getFullYear();
      const hours = String(d.getHours()).padStart(2, '0');
      const minutes = String(d.getMinutes()).padStart(2, '0');
      return `${day}-${month}-${year} ${hours}:${minutes}`;
    }

    function formatDate(d) {
      if (!d) return '-';
      const dateObj = typeof d === 'string' ? new Date(d) : d;
      if (isNaN(dateObj.getTime())) return d;
      const day = String(dateObj.getDate()).padStart(2, '0');
      const month = String(dateObj.getMonth() + 1).padStart(2, '0');
      const year = dateObj.getFullYear();
      return `${day}-${month}-${year}`;
    }

    const currentRealYear = new Date().getFullYear();
    const defaultAppYear = currentRealYear >= 2026 && currentRealYear <= 2085 ? currentRealYear : 2026;

    let state = {
      yearlyCounters: { [defaultAppYear]: 0 },
      bookings: [],
      roomsCapacity: [
        { roomNo: 1, capacity: 4 },
        { roomNo: 2, capacity: 2 },
        { roomNo: 3, capacity: 4 },
        { roomNo: 4, capacity: 4 },
        { roomNo: 5, capacity: 4 }
      ],
      masterAgents: [
        { agentName: "Self", phone: "Direct", roomNo: "All Rooms" },
        { agentName: "A1", phone: "1234567890", roomNo: "All Rooms" },
        { agentName: "A2", phone: "1234567890", roomNo: "All Rooms" },
        { agentName: "A3", phone: "1234567890", roomNo: "All Rooms" },
        { agentName: "A4", phone: "1234567890", roomNo: "All Rooms" }
      ],
      selectedYear: defaultAppYear,
      dashSelectedYear: defaultAppYear
    };

    function isRoomInMaster(roomNo) {
      return state.roomsCapacity.some(m => parseInt(m.roomNo) === parseInt(roomNo));
    }

    function exportToExcel() {
      if (!state.bookings || state.bookings.length === 0) {
        alert("No booking records available to export!");
        return;
      }

      const now = new Date().getTime();

      const exportData = state.bookings.map(b => {
        const cIn = new Date(b.checkIn).getTime();
        const cOut = new Date(b.checkOut).getTime();

        let statusStr = "Upcoming";
        if (b.inactive) {
          statusStr = "Inactive"; 
        } else if (now >= cIn && now <= cOut) {
          statusStr = "Live";     
        } else if (now < cIn) {
          statusStr = "Upcoming"; 
        } else if (now > cOut) {
          statusStr = "Closed";   
        }

        let foodSummary = "";
        if (b.foodOrders && b.foodOrders.length > 0) {
          foodSummary = b.foodOrders.map(f => `${f.foodDesc || 'Food'} (${f.plates || 1} Plates - ₹${f.foodCharge || 0})`).join(", ");
        }

        return {
          "Booking ID": b.bookingCode || "-",
          "Invoice ID": b.invoiceNo || "-",
          "Status": statusStr,
          "Guest Name": b.name || "-",
          "Contact No": b.contactNo || "-",
          "ID Number": b.idNo || "-",
          "Address": b.address || "-",
          "City": b.city || "-",
          "State": b.state || "-",
          "Country": b.country || "-",
          "Pin/Zip Code": b.zipCode || "-",
          "Room No": b.roomNo || "-",
          "Capacity": b.capacity || 1,
          "Agent Info": b.agentInfo || "-",
          "Check-In": formatDateTime(b.checkIn),
          "Check-Out": formatDateTime(b.checkOut),
          "Stay Days": b.noOfDays || 0,
          "Price / Day": b.perDayPrice || 0,
          "Food Orders": foodSummary || "None",
          "Total Amount": b.totalAmount || 0,
          "Advance Paid": b.advanced || 0,
          "Cleared Due": b.clearedDue || 0,
          "Balance Due": b.totalDue || 0
        };
      });

      const worksheet = XLSX.utils.json_to_sheet(exportData);
      const workbook = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(workbook, worksheet, "Bookings");

      XLSX.writeFile(workbook, `Booking_Report_${formatDate(new Date())}.xlsx`);
    }

    function searchBookingByDate() {
      const dateVal = document.getElementById('booking-date-search').value;
      renderBookingsTable(dateVal);
    }

    function clearDateSearchBooking() {
      const input = document.getElementById('booking-date-search');
      if (input) input.value = "";
      renderBookingsTable();
    }

    /* POINT 2: MASTER DATA BOOKING SEARCH SHOWS ONLY NON-DELETED BOOKINGS */
    function searchMasterBookingById() {
      const inputElem = document.getElementById('master-booking-search-input');
      if (!inputElem) return;

      const query = inputElem.value.trim().toUpperCase();
      const tbody = document.getElementById('master-delete-tbody');
      if (!tbody) return;

      tbody.innerHTML = '';

      if (!query) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-slate-400">Please type a Booking ID into the search field above to view and delete details.</td></tr>`;
        return;
      }

      // Filter out deleted/inactive bookings from the Master Search list
      const matchedBookings = state.bookings.filter(item => 
        !item.inactive && (item.bookingCode || '').toUpperCase().includes(query)
      );

      if (matchedBookings.length === 0) {
        tbody.innerHTML = `<tr><td colspan="7" class="text-center py-4 text-rose-500 font-semibold">No active booking found matching "${query}".</td></tr>`;
        return;
      }

      matchedBookings.forEach(b => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2 px-2 font-mono font-bold text-indigo-700">${b.bookingCode}</td>
          <td class="py-2 px-2 font-bold text-slate-800">${b.name}</td>
          <td class="py-2 px-2"><span class="bg-indigo-50 text-indigo-700 font-bold px-1.5 py-0.2 rounded text-[10px]">Room ${b.roomNo}</span></td>
          <td class="py-2 px-2 text-[10px] text-slate-600">${formatDateTime(b.checkIn)} to ${formatDateTime(b.checkOut)}</td>
          <td class="py-2 px-2 font-semibold text-slate-800">₹${b.totalAmount}</td>
          <td class="py-2 px-2 font-bold text-rose-600">₹${b.totalDue}</td>
          <td class="py-2 px-2 text-center">
            <button onclick="deleteBooking('${b.id}')" class="bg-rose-600 hover:bg-rose-700 text-white px-2.5 py-1 rounded text-[10px] font-semibold flex items-center gap-1 mx-auto transition shadow-xs">
              <i class="fa-solid fa-trash-can text-[9px]"></i> Delete Linked Booking
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function clearMasterBookingSearch() {
      const input = document.getElementById('master-booking-search-input');
      if (input) input.value = '';
      searchMasterBookingById();
    }

    function generateIDsForYear(checkInDateStr) {
      let targetYear = defaultAppYear;
      if (checkInDateStr) {
        targetYear = new Date(checkInDateStr).getFullYear() || defaultAppYear;
      }

      if (!state.yearlyCounters) state.yearlyCounters = {};

      if (!state.yearlyCounters[targetYear]) {
        const countForYear = state.bookings.filter(b => {
          return b.checkIn && new Date(b.checkIn).getFullYear() === targetYear;
        }).length;
        state.yearlyCounters[targetYear] = countForYear;
      }

      state.yearlyCounters[targetYear] += 1;
      const seq = state.yearlyCounters[targetYear];
      const paddedSeq = String(seq).padStart(7, '0');

      return {
        bookingCode: `BKG-${targetYear}-${paddedSeq}`,
        invoiceNo: `INV-${targetYear}-${paddedSeq}`
      };
    }

    function loadSavedData() {
      const saved = localStorage.getItem('webapp_data');
      if (saved) {
        try { 
          const parsed = JSON.parse(saved);
          if (parsed.bookings) {
            state = parsed;
            if (!state.roomsCapacity || state.roomsCapacity.length === 0) {
              state.roomsCapacity = [
                { roomNo: 1, capacity: 4 },
                { roomNo: 2, capacity: 2 },
                { roomNo: 3, capacity: 4 },
                { roomNo: 4, capacity: 4 },
                { roomNo: 5, capacity: 4 }
              ];
            }
            if (!state.masterAgents) {
              state.masterAgents = [
                { agentName: "Self", phone: "Direct", roomNo: "All Rooms" },
                { agentName: "A1", phone: "1234567890", roomNo: "All Rooms" },
                { agentName: "A2", phone: "1234567890", roomNo: "All Rooms" },
                { agentName: "A3", phone: "1234567890", roomNo: "All Rooms" },
                { agentName: "A4", phone: "1234567890", roomNo: "All Rooms" }
              ];
            }
            state.dashSelectedYear = defaultAppYear;
            state.selectedYear = defaultAppYear;
            if (!state.yearlyCounters) {
              state.yearlyCounters = { [defaultAppYear]: state.bookings.length || 0 };
            }
          }
        } catch(e){}
      }
      state.selectedYear = defaultAppYear;
    }

    function setMinBookingDates() {
      const today = new Date().toISOString().split('T')[0];
      const checkInInput = document.getElementById('cust-checkin-date');
      const checkOutInput = document.getElementById('cust-checkout-date');
      if (checkInInput) checkInInput.min = today;
      if (checkOutInput) checkOutInput.min = today;
    }

    document.addEventListener("DOMContentLoaded", () => {
      checkAuthStatus();
      loadSavedData();
      setMinBookingDates();
      populateDashboardYearDropdown();
      initDashboard();
      populateRoomDropdown();
      populateAgentDropdown();
      populateCalendarYearDropdown();
      searchMasterBookingById();
      renderBookingsTable();
      renderRoomCapacityTable();
      renderMasterAgentTable();
      renderCalendar(defaultAppYear);

      checkUpcomingCheckoutsWithDue();
      setInterval(checkUpcomingCheckoutsWithDue, 60000);
      setInterval(triggerPeriodicAutoSave, 300000);
    });

    function triggerPeriodicAutoSave() {
      saveChanges(true, true);
    }

    function saveChanges(isAutoSave = false, quiet = false) {
      localStorage.setItem('webapp_data', JSON.stringify(state));
      if (!quiet) {
        const toast = document.getElementById('toast');
        const msg = document.getElementById('toast-message');
        msg.innerText = isAutoSave ? 'Changes Auto save successfully!' : 'Data saved successfully!';
        toast.classList.remove('hidden');
        setTimeout(() => toast.classList.add('hidden'), 3000);
      }
    }

    function populateDashboardYearDropdown() {
      const yearSelect = document.getElementById('dash-year-select');
      if (!yearSelect) return;
      yearSelect.innerHTML = '';

      const optConsolidated = document.createElement('option');
      optConsolidated.value = "ALL";
      optConsolidated.text = "All Years (Consolidated)";
      yearSelect.appendChild(optConsolidated);

      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = y === defaultAppYear ? `${y} (Current Year)` : `Year ${y}`;
        yearSelect.appendChild(opt);
      }

      yearSelect.value = defaultAppYear;
      state.dashSelectedYear = defaultAppYear;
    }

    function handleDashboardYearChange(val) {
      if (val === 'CURRENT') {
        val = defaultAppYear;
      }
      
      const select = document.getElementById('dash-year-select');
      if (select) select.value = val;

      if (val === 'ALL') {
        state.dashSelectedYear = 'ALL';
      } else {
        state.dashSelectedYear = parseInt(val);
      }

      initDashboard();
    }

    /* POINT 3: ALERT MESSAGE LOGIC IMPLEMENTATION */
    function checkUpcomingCheckoutsWithDue() {
      const now = new Date().getTime();
      const twoHoursMs = 2 * 60 * 60 * 1000;

      const alertBookings = state.bookings.filter(b => {
        if (!b.checkOut || !isRoomInMaster(b.roomNo) || b.inactive) return false;
        
        const checkOutTime = new Date(b.checkOut).getTime();
        const diff = checkOutTime - now;
        
        // Logic 1: Any booking that has due, notify until due clear
        const hasDue = (b.totalDue || 0) > 0;

        // Logic 2: Notify before 2 hours whenever any booking check out date and time is nearby
        // (diff > 0 ensures it triggers in the 2-hour window leading up to checkout)
        const isNearbyCheckOut = (diff >= 0 && diff <= twoHoursMs);

        return (hasDue || isNearbyCheckOut);
      });

      const badge = document.getElementById('alert-badge');
      if (alertBookings.length > 0) {
        badge.innerText = alertBookings.length;
        badge.classList.remove('hidden');
      } else {
        badge.classList.add('hidden');
      }

      renderAlertModalList(alertBookings);
    }

    function renderAlertModalList(alertList) {
      const container = document.getElementById('alert-list-container');
      const textCount = document.getElementById('alert-list-count-text');
      container.innerHTML = '';

      textCount.innerText = `${alertList.length} active warnings found`;

      if (alertList.length === 0) {
        container.innerHTML = `
          <div class="text-center py-8 space-y-1">
            <div class="bg-emerald-50 text-emerald-600 w-10 h-10 rounded-full flex items-center justify-center mx-auto text-base">
              <i class="fa-solid fa-circle-check"></i>
            </div>
            <p class="font-bold text-slate-700">No Check-out Alerts</p>
            <p class="text-slate-400 text-[10px]">No upcoming check-outs within 2 hours or pending dues found.</p>
          </div>
        `;
        return;
      }

      alertList.forEach((b, i) => {
        const timeFormatted = formatDateTime(b.checkOut);
        const hasDue = (b.totalDue || 0) > 0;

        const card = document.createElement('div');
        card.className = "bg-amber-50/60 border border-amber-200 rounded-md overflow-hidden shadow-sm";
        
        let alertMessageText = "";
        let alertBadgeHtml = "";

        if (!hasDue) {
          alertMessageText = `Checkout: <strong>${timeFormatted}</strong> | Room ${b.roomNo} | Guest: <strong>${b.name}</strong>`;
          alertBadgeHtml = `<span class="text-[11px] font-bold text-emerald-600 bg-emerald-50 border border-emerald-200 px-2 py-0.5 rounded">No Due (Paid)</span>`;
        } else {
          alertMessageText = `Checkout: <strong>${timeFormatted}</strong> | Room ${b.roomNo} | Guest: <strong>${b.name}</strong> | Total: ₹${b.totalAmount} | Due: ₹${b.totalDue}`;
          alertBadgeHtml = `<span class="text-[11px] font-black text-rose-600 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">₹${b.totalDue.toLocaleString('en-IN')} Due</span>`;
        }

        card.innerHTML = `
          <div class="p-2.5 flex justify-between items-center cursor-pointer hover:bg-amber-100/50 transition" onclick="toggleAlertDetails('alert-details-${i}')">
            <div class="flex items-center space-x-2">
              <span class="bg-amber-500 text-white p-1.5 rounded text-[10px] font-bold"><i class="fa-solid fa-clock"></i></span>
              <div>
                <h4 class="font-bold text-slate-800 text-[11px] flex items-center gap-1.5">
                  ${b.name} <span class="bg-indigo-100 text-indigo-800 text-[9px] px-1.5 py-0.2 rounded font-mono">${b.bookingCode || 'N/A'}</span>
                  <span class="bg-slate-100 text-slate-700 text-[9px] px-1.5 py-0.2 rounded font-medium">Room ${b.roomNo}</span>
                </h4>
                <p class="text-[10px] text-slate-600 mt-0.5">${alertMessageText}</p>
              </div>
            </div>
            <div class="flex items-center space-x-1.5">
              ${alertBadgeHtml}
              <i class="fa-solid fa-chevron-down text-slate-400 text-[10px]"></i>
            </div>
          </div>

          <div id="alert-details-${i}" class="hidden bg-white border-t border-amber-200/60 p-2.5 space-y-2 text-[10px]">
            <div class="grid grid-cols-2 gap-1 text-slate-600">
              <div>Total Charges: <strong>₹${b.totalAmount}</strong></div>
              <div>Advance Paid: <strong class="text-emerald-600">₹${b.advanced}</strong></div>
            </div>
            <div class="flex justify-end pt-1 border-t border-slate-100">
              <button onclick="closeAlertModal(); openBookingModal('${b.id}')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded font-bold text-[10px] flex items-center gap-1 transition">
                <i class="fa-solid fa-wallet"></i> View / Edit Booking
              </button>
            </div>
          </div>
        `;
        container.appendChild(card);
      });
    }

    function toggleAlertDetails(elemId) {
      const detailsBox = document.getElementById(elemId);
      if (detailsBox) detailsBox.classList.toggle('hidden');
    }

    function openAlertModal() {
      checkUpcomingCheckoutsWithDue();
      document.getElementById('alert-modal').classList.remove('hidden');
    }

    function closeAlertModal() {
      document.getElementById('alert-modal').classList.add('hidden');
    }

    function populateCalendarYearDropdown() {
      const yearSelect = document.getElementById('cal-year-select');
      if (!yearSelect) return;
      yearSelect.innerHTML = '';
      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = y === defaultAppYear ? `${y} (Current Year)` : `Year ${y}`;
        if (y === defaultAppYear) opt.selected = true;
        yearSelect.appendChild(opt);
      }
    }

    function populateRoomDropdown(selectedRoomNo = 1) {
      const roomSelect = document.getElementById('cust-room');
      if (!roomSelect) return;
      roomSelect.innerHTML = '';

      state.roomsCapacity.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.roomNo;
        opt.text = `Room ${m.roomNo}`;
        if (parseInt(m.roomNo) === parseInt(selectedRoomNo)) {
          opt.selected = true;
        }
        roomSelect.appendChild(opt);
      });
      autoCaptureRoomDetails();
    }

    function populateAgentDropdown(selectedAgentName = "") {
      const agentSelect = document.getElementById('cust-agent');
      if (!agentSelect) return;
      agentSelect.innerHTML = '';

      state.masterAgents.forEach(a => {
        const opt = document.createElement('option');
        opt.value = `${a.agentName} (${a.phone})`;
        opt.text = `${a.agentName} (${a.phone})`;
        if (selectedAgentName && opt.value.includes(selectedAgentName)) {
          opt.selected = true;
        }
        agentSelect.appendChild(opt);
      });
    }

    function autoCaptureRoomDetails() {
      const selectedRoom = document.getElementById('cust-room').value;
      const matched = state.roomsCapacity.find(m => parseInt(m.roomNo) === parseInt(selectedRoom));

      if (matched) {
        document.getElementById('cust-capacity').value = matched.capacity || 1;
      }
      calculateModalBilling();
    }

    function switchTab(tabId) {
      if (tabId === 'master' && !isMasterUnlocked) {
        openMasterAuthModal();
        return;
      }
      performSwitchTab(tabId);
    }

    function performSwitchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active-tab', 'bg-indigo-600', 'text-white');
        btn.classList.add('text-indigo-100', 'hover:bg-indigo-600/50');
      });

      document.getElementById(`tab-${tabId}`).classList.remove('hidden');
      const activeBtn = document.getElementById(`btn-${tabId}`);
      activeBtn.classList.add('active-tab', 'bg-indigo-600', 'text-white');
      closeCommentBox();

      if (tabId === 'dashboard') {
        handleDashboardYearChange(defaultAppYear);
      }
    }

    function selectDashboardYear(year) {
      handleDashboardYearChange(year);
      renderCalendar(year);
      switchTab('calendar');
    }

    function initDashboard() {
      const grid = document.getElementById('years-grid');
      grid.innerHTML = '';
      
      for (let y = 2026; y <= 2085; y++) {
        const item = document.createElement('div');
        const isSelectedYear = state.dashSelectedYear !== 'ALL' && parseInt(state.dashSelectedYear) === y;
        const isCurrentRealYear = y === defaultAppYear;

        item.className = `text-center py-1 px-1 rounded text-[10px] font-bold border cursor-pointer transition ${
          isSelectedYear
            ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm' 
            : (isCurrentRealYear ? 'bg-amber-100 text-amber-900 border-amber-400 font-extrabold hover:bg-amber-200' : 'bg-white text-slate-600 border-slate-200 hover:bg-indigo-50 hover:text-indigo-600')
        }`;
        
        item.innerText = y;
        if (isCurrentRealYear) {
          item.title = "Current Active Year";
        }

        item.onclick = () => selectDashboardYear(y);
        grid.appendChild(item);
      }

      updateDashboardCards();
    }

    function updateDashboardCards() {
      const selectedFilter = state.dashSelectedYear;
      const label = document.getElementById('dash-filter-label');

      let filteredBookings = [];

      if (selectedFilter === 'ALL' || !selectedFilter) {
        filteredBookings = state.bookings.filter(b => !b.inactive);
        if (label) label.innerText = "Consolidated Summary (All Years)";
      } else {
        const targetYear = parseInt(selectedFilter);
        filteredBookings = state.bookings.filter(b => {
          if (b.inactive || !b.checkIn) return false;
          const yr = new Date(b.checkIn).getFullYear();
          return yr === targetYear;
        });

        if (label) {
          label.innerText = targetYear === defaultAppYear 
            ? `Year ${targetYear} (Current Year)` 
            : `Year ${targetYear}`;
        }
      }

      const totalBookings = filteredBookings.length;
      const totalAmt = filteredBookings.reduce((sum, b) => sum + (b.totalAmount || 0), 0);
      const totalAdv = filteredBookings.reduce((sum, b) => sum + (b.advanced || 0), 0);
      const totalDue = filteredBookings.reduce((sum, b) => sum + (b.totalDue || 0), 0);

      document.getElementById('dash-total-bookings').innerText = totalBookings;
      document.getElementById('dash-total-amount').innerText = `₹${totalAmt.toLocaleString('en-IN')}`;
      document.getElementById('dash-advanced').innerText = `₹${totalAdv.toLocaleString('en-IN')}`;
      document.getElementById('dash-due').innerText = `₹${totalDue.toLocaleString('en-IN')}`;
    }

    // PRINT INVOICE WITH DYNAMIC READ-ONLY MODE LOGIC
    function printInvoice(bookingId) {
      const bIndex = state.bookings.findIndex(item => item.id === bookingId);
      if (bIndex === -1) return;

      const b = state.bookings[bIndex];
      const now = new Date().getTime();
      const checkOutTime = new Date(b.checkOut).getTime();

      const isClosed = now > checkOutTime;
      const isInactive = b.inactive;
      const isMasterValid = isRoomInMaster(b.roomNo);

      const readOnlyNotice = document.getElementById('inv-readonly-notice');
      const invPrintBtn = document.getElementById('inv-print-btn');
      const invBadge = document.getElementById('inv-badge');
      const invIdContainer = document.getElementById('inv-id-container');

      // Check if entry is Inactive or Master Removed -> Disable print button, keep Read-Only modal view
      if (isInactive || !isMasterValid) {
        readOnlyNotice.classList.remove('hidden');
        if (invPrintBtn) invPrintBtn.classList.add('hidden'); // No Print Option
      } else if (isClosed) {
        // Closed Booking -> Print option available, but system read-only mode banner displayed
        readOnlyNotice.classList.remove('hidden');
        if (invPrintBtn) {
          invPrintBtn.classList.remove('hidden');
          invPrintBtn.disabled = false;
          invPrintBtn.className = "px-4 py-1 bg-indigo-600 text-white rounded font-semibold shadow flex items-center gap-1 transition hover:bg-indigo-700 cursor-pointer";
          invPrintBtn.innerHTML = `<i class="fa-solid fa-print"></i> Print Invoice`;
        }
      } else {
        readOnlyNotice.classList.add('hidden');
        if (invPrintBtn) invPrintBtn.classList.remove('hidden');
      }

      const hasDue = b.totalDue > 0;
      const today = formatDate(new Date());

      if (hasDue && !isClosed && !isInactive && isMasterValid) {
        invBadge.classList.add('hidden');
        invIdContainer.classList.add('hidden');
        if (invPrintBtn) {
          invPrintBtn.disabled = true;
          invPrintBtn.className = "px-4 py-1 bg-slate-300 text-slate-500 rounded font-semibold cursor-not-allowed flex items-center gap-1 border border-slate-400";
          invPrintBtn.innerHTML = `<i class="fa-solid fa-lock"></i> Print Disabled (Clear Due)`;
        }
      } else if (!isClosed && !isInactive && isMasterValid) {
        invBadge.classList.remove('hidden');
        invIdContainer.classList.remove('hidden');
        document.getElementById('inv-id').innerText = b.invoiceNo;
        if (invPrintBtn) {
          invPrintBtn.disabled = false;
          invPrintBtn.className = "px-4 py-1 bg-indigo-600 text-white rounded font-semibold shadow flex items-center gap-1 transition hover:bg-indigo-700 cursor-pointer";
          invPrintBtn.innerHTML = `<i class="fa-solid fa-print"></i> Print Invoice`;
        }
      }

      document.getElementById('inv-booking-id').innerText = b.bookingCode;
      document.getElementById('inv-date').innerText = today;

      const fullLocation = [b.address, b.city, b.state, b.country, b.zipCode].filter(Boolean).map(formatTitleCase).join(', ');
      document.getElementById('inv-guest-name').innerText = formatTitleCase(b.name) || 'N/A';
      document.getElementById('inv-guest-address').innerText = `Address: ${fullLocation || 'N/A'}`;
      document.getElementById('inv-guest-contact').innerText = `Contact: ${b.contactNo || 'N/A'}`;
      document.getElementById('inv-guest-id').innerText = `ID No: ${b.idNo || 'N/A'}`;

      document.getElementById('inv-room').innerText = `Room No: ${b.roomNo} (${b.capacity || 1} Person)`;
      document.getElementById('inv-checkin').innerText = `Check-in: ${formatDateTime(b.checkIn)}`;
      document.getElementById('inv-checkout').innerText = `Check-out: ${formatDateTime(b.checkOut)}`;

      const tbody = document.getElementById('inv-items-tbody');
      tbody.innerHTML = '';

      const roomTotal = (b.noOfDays || 0) * (b.perDayPrice || 0) * (b.capacity || 1);
      const roomTr = document.createElement('tr');
      roomTr.innerHTML = `
        <td class="p-2 font-semibold text-slate-800">Room ${b.roomNo} Accommodation (${b.capacity || 1} Persons)</td>
        <td class="p-2 text-center">${b.noOfDays} Days</td>
        <td class="p-2 text-right">₹${(b.perDayPrice || 0).toLocaleString('en-IN')}</td>
        <td class="p-2 text-right font-semibold text-slate-800">₹${roomTotal.toLocaleString('en-IN')}</td>
      `;
      tbody.appendChild(roomTr);

      if (b.foodOrders && b.foodOrders.length > 0) {
        b.foodOrders.forEach(fo => {
          if (fo.foodCharge > 0) {
            const foodTr = document.createElement('tr');
            const foodDateTimeFmt = fo.foodDateTime ? ` (${formatDateTime(fo.foodDateTime)})` : '';
            foodTr.innerHTML = `
              <td class="p-2 font-semibold text-slate-800">Extra Food <span class="text-[9px] text-slate-500 font-normal block">${fo.foodDesc || 'Food Item'}${foodDateTimeFmt}</span></td>
              <td class="p-2 text-center">${fo.plates || 1} Plates</td>
              <td class="p-2 text-right">₹${(fo.itemPrice || 0).toLocaleString('en-IN')}</td>
              <td class="p-2 text-right font-semibold text-slate-800">₹${(fo.foodCharge || 0).toLocaleString('en-IN')}</td>
            `;
            tbody.appendChild(foodTr);
          }
        });
      }

      const initialAdv = b.initialAdv !== undefined ? b.initialAdv : b.advanced;
      const clearDueAmt = b.clearedDue || 0;

      document.getElementById('inv-sum-total').innerText = `₹${(b.totalAmount || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-advance').innerText = `₹${initialAdv.toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-due').innerText = `₹${(b.totalDue || 0).toLocaleString('en-IN')}`;

      // SHOW CLEAR DUE ROW ONLY IF FILLED
      const clearDueRow = document.getElementById('inv-clear-due-row');
      if (clearDueAmt > 0) {
        document.getElementById('inv-sum-clear-due').innerText = `₹${clearDueAmt.toLocaleString('en-IN')}`;
        clearDueRow.classList.remove('hidden');
      } else {
        clearDueRow.classList.add('hidden');
      }

      document.getElementById('invoice-modal').classList.remove('hidden');
    }

    function closeInvoiceModal() {
      document.getElementById('invoice-modal').classList.add('hidden');
    }

    function addFoodOrderItem(desc = '', plates = 1, itemPrice = 0, charge = 0, dateStr = '', timeStr = '') {
      const container = document.getElementById('food-orders-container');
      const itemRow = document.createElement('div');
      itemRow.className = "food-order-row grid grid-cols-1 sm:grid-cols-12 gap-1.5 items-end bg-white p-2 rounded border border-amber-200/80 shadow-xs";
      
      itemRow.innerHTML = `
        <div class="sm:col-span-3">
          <label class="block font-semibold text-slate-600 mb-0.5">Item Name</label>
          <input type="text" value="${desc}" placeholder="e.g. Thali / Tea" class="cust-food-desc w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500">
        </div>

        <div class="sm:col-span-4">
          <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-regular fa-clock text-amber-600 mr-1"></i> Date & Time</label>
          <div class="flex gap-1">
            <input type="date" value="${dateStr}" class="cust-food-date w-3/5 bg-white border border-slate-300 rounded px-1 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500 font-medium text-[10px]">
            <input type="time" value="${timeStr}" class="cust-food-time w-2/5 bg-white border border-slate-300 rounded px-1 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500 font-medium text-[10px]">
          </div>
        </div>

        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Price/Plate (₹)</label>
          <input type="number" value="${itemPrice}" min="0" oninput="calculateFoodRowTotal(this)" class="cust-food-price w-full bg-white border border-slate-300 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500">
        </div>
        
        <div class="sm:col-span-1">
          <label class="block font-semibold text-slate-600 mb-0.5">Plates</label>
          <input type="number" value="${plates}" min="1" oninput="calculateFoodRowTotal(this)" class="cust-food-plates w-full bg-white border border-slate-300 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500 font-bold">
        </div>

        <div class="sm:col-span-1">
          <label class="block font-semibold text-slate-600 mb-0.5">Total (₹)</label>
          <input type="number" value="${charge}" readonly class="cust-food-charge w-full bg-slate-100 font-bold text-amber-700 border border-slate-300 rounded px-1.5 py-1 cursor-not-allowed">
        </div>

        <div class="sm:col-span-1 flex justify-end">
          <button type="button" onclick="removeFoodOrderItem(this)" class="btn-remove-food-item text-rose-500 hover:text-rose-700 p-1.5" title="Remove Order">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </div>
      `;

      container.appendChild(itemRow);
      calculateModalBilling();
    }

    function calculateFoodRowTotal(inputElem) {
      const row = inputElem.closest('.food-order-row');
      if (!row) return;

      const price = parseFloat(row.querySelector('.cust-food-price').value) || 0;
      const plates = parseInt(row.querySelector('.cust-food-plates').value) || 0;
      const totalCharge = price * plates;

      row.querySelector('.cust-food-charge').value = totalCharge;
      calculateModalBilling();
    }

    function removeFoodOrderItem(btn) {
      const row = btn.closest('.food-order-row');
      if (row) {
        row.remove();
        calculateModalBilling();
      }
    }

    /* POINT 1: OPEN MODAL VERIFICATION WITH 3 DAYS LIMIT */
    function openBookingModal(bookingId = null) {
      if (bookingId) {
        const b = state.bookings.find(item => item.id === bookingId);
        
        if (b) {
          const now = new Date().getTime();
          const checkOutTime = new Date(b.checkOut).getTime();
          const threeDaysMs = 3 * 24 * 60 * 60 * 1000;

          if (b.inactive) {
            alert("This booking details are inactive and cannot be edited.");
            return;
          }

          // Booking is editable till 3 days from check out date, after that non-editable
          if (now > (checkOutTime + threeDaysMs)) {
            alert("This booking closed more than 3 days ago. It is non-editable and can only be viewed or printed.");
            return;
          }

          if (!isRoomInMaster(b.roomNo)) {
            alert("This booking details were deleted from Master Data and cannot be opened or edited.");
            return;
          }
        }
      }

      setMinBookingDates();
      const form = document.getElementById('booking-form');
      form.reset();
      removeAttachedIdProof();
      
      const clearBillInput = document.getElementById('cust-clear-bill');
      if (clearBillInput) clearBillInput.value = 0;

      document.getElementById('food-orders-container').innerHTML = '';

      populateAgentDropdown();

      if (bookingId) {
        const b = state.bookings.find(item => item.id === bookingId);
        if (b) {
          document.getElementById('modal-title').innerText = 'Edit Booking Details';
          document.getElementById('modal-booking-id').value = b.id;
          document.getElementById('cust-name').value = formatTitleCase(b.name);
          document.getElementById('cust-address').value = formatTitleCase(b.address || '');
          document.getElementById('cust-city').value = formatTitleCase(b.city || '');
          document.getElementById('cust-state').value = formatTitleCase(b.state || '');
          document.getElementById('cust-country').value = formatTitleCase(b.country || '');
          document.getElementById('cust-zip').value = b.zipCode || '';
          document.getElementById('cust-id').value = b.idNo || '';
          document.getElementById('cust-contact').value = b.contactNo || '';

          if (b.idProofBase64) {
            document.getElementById('cust-id-file-base64').value = b.idProofBase64;
            document.getElementById('cust-id-file-name').value = b.idProofFileName || 'Attached_ID_Proof.pdf';
            document.getElementById('cust-id-file-status').innerHTML = `<span class="text-emerald-600 font-semibold"><i class="fa-solid fa-circle-check"></i> Attached: ${b.idProofFileName || 'Attached_ID_Proof.pdf'}</span>`;
            document.getElementById('cust-id-file-remove').classList.remove('hidden');
          }
          
          populateRoomDropdown(b.roomNo);
          populateAgentDropdown(b.agentInfo);
          document.getElementById('cust-capacity').value = b.capacity || 1;

          if (b.checkIn) {
            const [inDate, inTime] = b.checkIn.split('T');
            document.getElementById('cust-checkin-date').value = inDate || '';
            document.getElementById('cust-checkin-time').value = inTime || '12:00';
          }
          if (b.checkOut) {
            const [outDate, outTime] = b.checkOut.split('T');
            document.getElementById('cust-checkout-date').value = outDate || '';
            document.getElementById('cust-checkout-time').value = outTime || '11:00';
          }

          if (b.foodOrders && b.foodOrders.length > 0) {
            b.foodOrders.forEach(fo => {
              let fDate = '', fTime = '';
              if (fo.foodDateTime) {
                const parts = fo.foodDateTime.split('T');
                fDate = parts[0] || '';
                fTime = parts[1] || '';
              }
              addFoodOrderItem(fo.foodDesc || '', fo.plates || 1, fo.itemPrice || 0, fo.foodCharge || 0, fDate, fTime);
            });
          }

          document.getElementById('cust-price').value = b.perDayPrice;
          
          const advanceElem = document.getElementById('cust-advance');
          const baseAdv = b.initialAdv !== undefined ? b.initialAdv : b.advanced;
          advanceElem.value = baseAdv;
          advanceElem.setAttribute('data-initial-adv', baseAdv);

          if (b.clearedDue) {
            document.getElementById('cust-clear-bill').value = b.clearedDue;
          }

          calculateModalBilling();
        }
      } else {
        document.getElementById('modal-title').innerText = 'Add New Booking';
        document.getElementById('modal-booking-id').value = '';
        
        populateRoomDropdown(state.roomsCapacity.length > 0 ? state.roomsCapacity[0].roomNo : 1);

        document.getElementById('cust-checkin-time').value = "12:00";
        document.getElementById('cust-checkout-time').value = "11:00";
        document.getElementById('cust-price').value = 1200;
        
        const advanceElem = document.getElementById('cust-advance');
        advanceElem.value = 0;
        advanceElem.setAttribute('data-initial-adv', 0);

        calculateModalBilling();
      }

      document.getElementById('booking-modal').classList.remove('hidden');
    }

    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }

    function handleStayDatesChange() {
      const today = new Date().toISOString().split('T')[0];
      const inDateInput = document.getElementById('cust-checkin-date');
      const outDateInput = document.getElementById('cust-checkout-date');

      if (inDateInput.value && inDateInput.value < today) {
        alert("⚠️ Check-In date cannot be prior to the current date!");
        inDateInput.value = today;
      }
      if (outDateInput.value && outDateInput.value < today) {
        alert("⚠️ Check-Out date cannot be prior to the current date!");
        outDateInput.value = today;
      }

      calculateModalBilling();
    }

    function handleClearBillPayment(clearAmountVal) {
      const clearVal = parseFloat(clearAmountVal) || 0;
      const total = parseFloat(document.getElementById('cust-total').value) || 0;
      
      const advanceElem = document.getElementById('cust-advance');
      let initialAdvance = parseFloat(advanceElem.getAttribute('data-initial-adv'));
      if (isNaN(initialAdvance)) {
        initialAdvance = parseFloat(advanceElem.value) || 0;
        advanceElem.setAttribute('data-initial-adv', initialAdvance);
      }

      if (clearVal + initialAdvance > total) {
        alert("⚠️ Payment amount exceeds the remaining balance due!");
        document.getElementById('cust-clear-bill').value = 0;
        document.getElementById('cust-due').value = Math.max(0, total - initialAdvance);
        return;
      }

      const due = Math.max(0, total - initialAdvance - clearVal);
      document.getElementById('cust-due').value = due;
    }

    function calculateModalBilling() {
      const inDate = document.getElementById('cust-checkin-date').value;
      const inTime = document.getElementById('cust-checkin-time').value || '00:00';
      const outDate = document.getElementById('cust-checkout-date').value;
      const outTime = document.getElementById('cust-checkout-time').value || '00:00';
      
      let days = 0;
      if (inDate && outDate) {
        const d1 = new Date(`${inDate}T${inTime}`);
        const d2 = new Date(`${outDate}T${outTime}`);
        const diff = d2 - d1;
        days = Math.max(0, Math.ceil(diff / (1000 * 60 * 60 * 24)));
      }

      const price = parseFloat(document.getElementById('cust-price').value) || 0;
      const capacity = parseFloat(document.getElementById('cust-capacity').value) || 1;
      
      let foodTotalCharge = 0;
      document.querySelectorAll('.cust-food-charge').forEach(input => {
        foodTotalCharge += parseFloat(input.value) || 0;
      });

      const roomTotal = days * price * capacity;
      const total = roomTotal + foodTotalCharge;

      const advanceInput = document.getElementById('cust-advance');
      let initialAdv = parseFloat(advanceInput.getAttribute('data-initial-adv'));
      if (isNaN(initialAdv)) {
        initialAdv = parseFloat(advanceInput.value) || 0;
        advanceInput.setAttribute('data-initial-adv', initialAdv);
      }

      const clearBillVal = parseFloat(document.getElementById('cust-clear-bill')?.value) || 0;
      const due = Math.max(0, total - initialAdv - clearBillVal);

      document.getElementById('cust-days').value = days;
      document.getElementById('cust-total').value = total;
      document.getElementById('cust-due').value = due;
    }

    function handleSaveBooking(e) {
      e.preventDefault();

      const guestName = formatTitleCase(document.getElementById('cust-name').value.trim());

      if (!guestName) {
        alert("⚠️ Guest Name is a mandatory field!");
        return;
      }

      const today = new Date().toISOString().split('T')[0];
      const inDate = document.getElementById('cust-checkin-date').value;
      const outDate = document.getElementById('cust-checkout-date').value;

      const bookingModalId = document.getElementById('modal-booking-id').value;

      if (!bookingModalId && inDate < today) {
        alert("⚠️ Cannot book new entry before today’s date!");
        return;
      }

      const id = bookingModalId;
      const roomNo = parseInt(document.getElementById('cust-room').value);

      const inTime = document.getElementById('cust-checkin-time').value || '00:00';
      const outTime = document.getElementById('cust-checkout-time').value || '00:00';

      const checkIn = `${inDate}T${inTime}`;
      const checkOut = `${outDate}T${outTime}`;

      const foodOrdersList = [];
      document.querySelectorAll('.food-order-row').forEach(row => {
        const desc = row.querySelector('.cust-food-desc').value || '';
        const itemPrice = parseFloat(row.querySelector('.cust-food-price').value) || 0;
        const plates = parseInt(row.querySelector('.cust-food-plates').value) || 0;
        const charge = parseFloat(row.querySelector('.cust-food-charge').value) || 0;
        const fDate = row.querySelector('.cust-food-date').value || '';
        const fTime = row.querySelector('.cust-food-time').value || '';

        const foodDateTime = (fDate && fTime) ? `${fDate}T${fTime}` : (fDate ? `${fDate}T00:00` : '');

        if (desc || charge > 0) {
          foodOrdersList.push({
            foodDesc: desc,
            itemPrice: itemPrice,
            plates: plates,
            foodCharge: charge,
            foodDateTime: foodDateTime
          });
        }
      });

      const newIn = new Date(checkIn).getTime();
      const newOut = new Date(checkOut).getTime();

      if (newIn >= newOut) {
        alert("Check-Out date & time must be strictly after Check-In date & time.");
        return;
      }

      const conflict = state.bookings.find(b => {
        if (b.inactive) return false;
        if (id && b.id === id) return false;
        if (parseInt(b.roomNo) !== roomNo) return false;

        const existingIn = new Date(b.checkIn).getTime();
        const existingOut = new Date(b.checkOut).getTime();

        return (newIn < existingOut && newOut > existingIn);
      });

      if (conflict) {
        const confOutFormatted = formatDateTime(conflict.checkOut);
        alert(`❌ Booking Conflict Alert!\n\nRoom ${roomNo} is already occupied by ${conflict.name} until ${confOutFormatted}.\n\nPlease select a check-in time after ${confOutFormatted} or assign a different room.`);
        return;
      }

      let existingCode = null;
      let existingInv = null;

      if (id) {
        const existing = state.bookings.find(b => b.id === id);
        if (existing) {
          existingCode = existing.bookingCode;
          existingInv = existing.invoiceNo;
        }
      }

      if (!existingCode) {
        const generated = generateIDsForYear(checkIn);
        existingCode = generated.bookingCode;
        existingInv = generated.invoiceNo;
      }

      const totalAmt = parseFloat(document.getElementById('cust-total').value) || 0;
      const initialAdvAmt = parseFloat(document.getElementById('cust-advance').getAttribute('data-initial-adv')) || parseFloat(document.getElementById('cust-advance').value) || 0;
      const clearedDueAmt = parseFloat(document.getElementById('cust-clear-bill').value) || 0;

      const totalPaid = initialAdvAmt + clearedDueAmt;
      
      const newBooking = {
        id: id || `bk_${Date.now()}`,
        bookingCode: existingCode,
        invoiceNo: existingInv,
        name: guestName,
        address: formatTitleCase(document.getElementById('cust-address').value.trim()),
        city: formatTitleCase(document.getElementById('cust-city').value.trim()),
        state: formatTitleCase(document.getElementById('cust-state').value.trim()),
        country: formatTitleCase(document.getElementById('cust-country').value.trim()),
        zipCode: document.getElementById('cust-zip').value.trim(),
        idNo: document.getElementById('cust-id').value.trim(),
        contactNo: document.getElementById('cust-contact').value.trim(),
        idProofBase64: document.getElementById('cust-id-file-base64').value,
        idProofFileName: document.getElementById('cust-id-file-name').value,
        roomNo: roomNo,
        agentInfo: document.getElementById('cust-agent').value,
        capacity: parseInt(document.getElementById('cust-capacity').value) || 1,
        checkIn: checkIn,
        checkOut: checkOut,
        noOfDays: parseInt(document.getElementById('cust-days').value) || 0,
        perDayPrice: parseFloat(document.getElementById('cust-price').value) || 0,
        foodOrders: foodOrdersList,
        totalAmount: totalAmt,
        initialAdv: initialAdvAmt,
        clearedDue: clearedDueAmt,
        advanced: totalPaid,
        totalDue: Math.max(0, totalAmt - totalPaid),
        inactive: false
      };

      if (id) {
        const idx = state.bookings.findIndex(b => b.id === id);
        if (idx !== -1) state.bookings[idx] = newBooking;
      } else {
        state.bookings.push(newBooking);
      }

      closeBookingModal();
      searchMasterBookingById();
      renderBookingsTable();
      updateDashboardCards();
      renderCalendar(defaultAppYear);
      checkUpcomingCheckoutsWithDue();
      saveChanges(false, false);
    }

    function deleteBooking(id) {
      openMasterDeleteModal('booking', id);
    }

    /* POINT 1: RENDER BOOKING TABLE ACTION CONTROLS EDITABILITY BY 3 DAYS POST CHECKOUT */
    function renderBookingsTable(dateFilter = "") {
      const tbody = document.getElementById('bookings-tbody');
      tbody.innerHTML = '';

      let listToRender = [...state.bookings];

      if (dateFilter) {
        listToRender = listToRender.filter(b => {
          if (!b.checkIn || !b.checkOut) return false;
          const bIn = b.checkIn.split('T')[0];
          const bOut = b.checkOut.split('T')[0];
          return (dateFilter >= bIn && dateFilter <= bOut);
        });
      }

      const now = new Date().getTime();
      const threeDaysMs = 3 * 24 * 60 * 60 * 1000; // 3 Days Window Limit

      const getStatusPriority = (b) => {
        const isMasterValid = isRoomInMaster(b.roomNo);
        if (!isMasterValid || b.inactive) {
          return 4; 
        }

        const cIn = new Date(b.checkIn).getTime();
        const cOut = new Date(b.checkOut).getTime();

        if (now >= cIn && now <= cOut) {
          return 1; 
        } else if (now < cIn) {
          return 2; 
        } else {
          return 3; 
        }
      };

      listToRender.sort((a, b) => {
        const priorityA = getStatusPriority(a);
        const priorityB = getStatusPriority(b);

        if (priorityA !== priorityB) {
          return priorityA - priorityB;
        }

        return (b.bookingCode || '').localeCompare(a.bookingCode || '');
      });

      if (listToRender.length === 0) {
        tbody.innerHTML = `<tr><td colspan="13" class="text-center py-6 text-slate-400">No bookings found for the selected date.</td></tr>`;
        return;
      }

      listToRender.forEach((b) => {
        const isMasterValid = isRoomInMaster(b.roomNo);
        const checkInFmt = formatDateTime(b.checkIn);
        const checkOutFmt = formatDateTime(b.checkOut);

        const checkInTime = new Date(b.checkIn).getTime();
        const checkOutTime = new Date(b.checkOut).getTime();

        const isClosed = now > checkOutTime;
        const isExpiredOver3Days = isClosed && (now > (checkOutTime + threeDaysMs));
        const isInactive = b.inactive;

        let statusBgClass = "hover:bg-slate-50";
        let statusDotHtml = "";

        if (!isMasterValid) {
          statusBgClass = "bg-rose-200/90 hover:bg-rose-300/80 text-rose-900";
        } else if (isInactive) {
          statusBgClass = "bg-slate-100 hover:bg-slate-200 text-slate-500 opacity-75";
        } else if (now < checkInTime) {
          statusDotHtml = `<span class="w-2.5 h-2.5 bg-blue-500 rounded-full inline-block flex-shrink-0" title="Upcoming Booking"></span>`;
        } else if (now >= checkInTime && now <= checkOutTime) {
          statusDotHtml = `
            <span class="relative flex h-2.5 w-2.5 flex-shrink-0" title="Live Booking">
              <span class="animate-ping absolute inline-flex h-full w-full rounded-full bg-amber-400 opacity-75"></span>
              <span class="relative inline-flex rounded-full h-2.5 w-2.5 bg-amber-500"></span>
            </span>
          `;
        } else if (isClosed) {
          statusDotHtml = `<span class="w-2.5 h-2.5 bg-emerald-500 rounded-full inline-block flex-shrink-0" title="Closed Booking"></span>`;
        }

        let foodSummaryHtml = '';
        if (b.foodOrders && b.foodOrders.length > 0) {
          const totalFoodCharge = b.foodOrders.reduce((acc, fo) => acc + (fo.foodCharge || 0), 0);
          if (totalFoodCharge > 0) {
            foodSummaryHtml = `<div class="text-[9px] ${!isMasterValid ? 'text-rose-950 font-bold' : 'text-amber-800 font-semibold'}"><i class="fa-solid fa-utensils text-[8px] mr-0.5"></i>Food (${b.foodOrders.length}): +₹${totalFoodCharge}</div>`;
          }
        }

        const printOnClick = `printInvoice('${b.id}')`;
        const printBtnClass = "bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded-md text-[11px] font-bold transition shadow-sm border border-indigo-700";

        let actionButtonsHtml = "";

        if (isExpiredOver3Days || isInactive || !isMasterValid) {
          // Closed > 3 Days OR Inactive / Removed: Non-editable, Print View only
          actionButtonsHtml = `
            <div class="flex items-center justify-center space-x-1">
              <button onclick="${printOnClick}" class="bg-slate-700 hover:bg-slate-800 text-white px-2.5 py-1 rounded-md text-[11px] font-bold transition shadow-sm border border-slate-900" title="View Entry in Read-Only / Print View Mode">
                <i class="fa-solid fa-eye text-[10px] mr-0.5"></i> Print View
              </button>
            </div>
          `;
        } else {
          // Editable Window (Live, Upcoming, or Closed within 3 Days)
          actionButtonsHtml = `
            <div class="flex items-center justify-center space-x-1">
              <button onclick="openBookingModal('${b.id}')" class="text-indigo-600 hover:text-indigo-800 p-1 text-sm" title="Edit Booking Details">
                <i class="fa-solid fa-pen-to-square"></i>
              </button>              
              <button onclick="${printOnClick}" class="${printBtnClass}">Print</button>
            </div>
          `;
        }

        let idProofCellHtml = `<span class="text-slate-400 italic text-[10px]">None</span>`;
        if (b.idProofBase64) {
          idProofCellHtml = `
            <button onclick="openPdfAttachment('${b.idProofBase64}')" class="bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 px-2 py-0.5 rounded text-[10px] font-bold flex items-center gap-1 transition">
              <i class="fa-solid fa-file-pdf text-rose-600"></i> View PDF
            </button>
          `;
        }

        const tr = document.createElement('tr');
        tr.className = `${statusBgClass} transition border-b border-slate-200/60`;
        tr.innerHTML = `
          <td class="py-2 px-2">
            <div class="flex items-center gap-1.5">
              ${statusDotHtml}
              ${b.inactive ? '<span class="w-2.5 h-2.5 bg-rose-600 rounded-full inline-block flex-shrink-0" title="Inactive Booking"></span>' : ''}
              <span class="bg-indigo-50 border border-indigo-200 text-indigo-700 font-mono font-bold px-1.5 py-0.2 rounded text-[9px] block w-max">${b.bookingCode}</span>
            </div>
            ${b.inactive ? '<span class="bg-slate-600 text-white font-bold px-1 py-0.1 rounded text-[8px] uppercase block mt-0.5 w-max">Inactive</span>' : (!isMasterValid ? '<span class="bg-rose-700 text-white font-bold px-1 py-0.1 rounded text-[8px] uppercase block mt-0.5 w-max">Master Removed</span>' : '')}
          </td>
          <td class="py-2 px-2 font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">${b.name}</td>
          <td class="py-2 px-2 font-medium">${b.contactNo || '-'}</td>
          <td class="py-2 px-2 font-mono text-[10px]">${b.idNo || '-'}</td>
          <td class="py-2 px-2">${idProofCellHtml}</td>
          <td class="py-2 px-2"><span class="bg-indigo-50 text-indigo-700 font-bold px-1.5 py-0.2 rounded text-[10px]">Room ${b.roomNo}</span></td>
          <td class="py-2 px-2 font-bold text-slate-700">${b.capacity || 1} Person</td>
          <td class="py-2 px-2 ${!isMasterValid ? 'text-rose-900' : 'text-slate-600'} text-[10px]">${b.agentInfo || '-'}</td>
          <td class="py-2 px-2 text-[10px]">
            <div class="font-semibold ${!isMasterValid ? 'text-rose-950' : 'text-slate-700'}">${checkInFmt}</div>
            <div class="${!isMasterValid ? 'text-rose-900' : 'text-slate-500'} text-[9px]">to ${checkOutFmt}</div>
          </td>
          <td class="py-2 px-2">
            <div class="font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">₹${b.perDayPrice}/day (${b.noOfDays}d)</div>
            ${foodSummaryHtml}
          </td>
          <td class="py-2 px-2 font-bold ${!isMasterValid ? 'text-rose-950' : 'text-slate-800'}">
            ₹${b.totalAmount}
            <span class="block text-[9px] text-emerald-600 font-medium">Adv: ₹${b.advanced}</span>
          </td>
          <td class="py-2 px-2 font-bold ${b.totalDue > 0 ? 'text-rose-600' : 'text-emerald-600'}">₹${b.totalDue}</td>
          <td class="py-2 px-2 text-center">
            ${actionButtonsHtml}
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    // Render Room Capacity Table in Master Tab
    function renderRoomCapacityTable() {
      const tbody = document.getElementById('room-capacity-tbody');
      tbody.innerHTML = '';

      if (!state.roomsCapacity || state.roomsCapacity.length === 0) {
        tbody.innerHTML = `<tr><td colspan="3" class="text-center py-4 text-slate-400">No room capacity data available.</td></tr>`;
        return;
      }

      state.roomsCapacity.forEach((r, idx) => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2 px-2">
            <input type="number" value="${r.roomNo}" min="1" oninput="state.roomsCapacity[${idx}].roomNo = parseInt(this.value) || 1" onchange="populateRoomDropdown(); renderBookingsTable(); saveChanges(false, true)" class="w-24 bg-transparent font-bold text-indigo-700 focus:bg-white focus:border focus:border-indigo-300 rounded px-1 py-0.5">
          </td>
          <td class="py-2 px-2">
            <input type="number" value="${r.capacity}" min="1" oninput="state.roomsCapacity[${idx}].capacity = parseInt(this.value) || 1" onchange="saveChanges(false, true)" class="w-24 bg-transparent font-semibold text-slate-800 focus:bg-white focus:border focus:border-indigo-300 rounded px-1 py-0.5">
          </td>
          <td class="py-2 px-2 text-center">
            <button type="button" onclick="removeRoomCapacityRow(${idx})" class="text-rose-500 hover:text-rose-700 p-1 text-xs" title="Delete Room Entry">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderMasterAgentTable() {
      const tbody = document.getElementById('agent-tbody');
      tbody.innerHTML = '';

      if (!state.masterAgents || state.masterAgents.length === 0) {
        tbody.innerHTML = `<tr><td colspan="4" class="text-center py-4 text-slate-400">No agent data available.</td></tr>`;
        return;
      }

      state.masterAgents.forEach((a, idx) => {
        const tr = document.createElement('tr');
        tr.className = "bg-white hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2 px-2">
            <input type="text" value="${a.agentName}" oninput="state.masterAgents[${idx}].agentName = formatTitleCase(this.value)" onchange="populateAgentDropdown(); saveChanges(false, true)" class="w-full bg-transparent font-semibold text-slate-800 focus:bg-white focus:border focus:border-indigo-300 rounded px-1 py-0.5">
          </td>
          <td class="py-2 px-2">
            <input type="text" value="${a.phone}" oninput="state.masterAgents[${idx}].phone = this.value" onchange="populateAgentDropdown(); saveChanges(false, true)" class="w-full bg-transparent text-slate-600 focus:bg-white focus:border focus:border-indigo-300 rounded px-1 py-0.5">
          </td>
          <td class="py-2 px-2 font-bold text-indigo-600">
            ${a.roomNo || 'All Rooms'}
          </td>
          <td class="py-2 px-2 text-center">
            <button type="button" onclick="removeAgentRow(${idx})" class="text-rose-500 hover:text-rose-700 p-1 text-xs" title="Delete Agent Entry">
              <i class="fa-solid fa-trash-can"></i>
            </button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function addRoomCapacityRow() {
      if (!state.roomsCapacity) state.roomsCapacity = [];
      const existingRoomNos = state.roomsCapacity.map(r => parseInt(r.roomNo) || 0);
      let nextRoom = 1;
      if (existingRoomNos.length > 0) {
        nextRoom = Math.max(...existingRoomNos) + 1;
      }

      state.roomsCapacity.push({
        roomNo: nextRoom,
        capacity: 4
      });

      renderRoomCapacityTable();
      populateRoomDropdown();
      saveChanges(false, false);
    }

    function removeRoomCapacityRow(index) {
      openMasterDeleteModal('room', index);
    }

    function addAgentRow() {
      if (!state.masterAgents) state.masterAgents = [];
      const nextNum = state.masterAgents.length;
      state.masterAgents.push({
        agentName: `Agent ${nextNum + 1}`,
        phone: "1234567890",
        roomNo: "All Rooms"
      });

      renderMasterAgentTable();
      populateAgentDropdown();
      saveChanges(false, false);
    }

    function removeAgentRow(index) {
      openMasterDeleteModal('agent', index);
    }

    function renderCalendar(year) {
      state.selectedYear = year;
      const calSelect = document.getElementById('cal-year-select');
      if (calSelect) calSelect.value = year;

      const container = document.getElementById('calendar-container');
      container.innerHTML = '';

      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

      months.forEach((monthName, monthIndex) => {
        const monthBox = document.createElement('div');
        monthBox.className = "bg-slate-50 rounded-lg p-2.5 border border-slate-200 shadow-xs flex flex-col justify-between";

        const title = document.createElement('h4');
        title.className = "font-bold text-slate-700 text-[11px] mb-1.5 pb-1 border-b border-slate-200 flex justify-between items-center";
        title.innerHTML = `<span>${monthName}</span> <span class="text-[9px] text-indigo-600 font-mono font-normal">${year}</span>`;
        monthBox.appendChild(title);

        const grid = document.createElement('div');
        grid.className = "grid grid-cols-7 gap-1 text-center text-[9px] font-medium text-slate-500 mb-1";
        
        ['S', 'M', 'T', 'W', 'T', 'F', 'S'].forEach(d => {
          const dh = document.createElement('div');
          dh.innerText = d;
          dh.className = "font-bold text-slate-400";
          grid.appendChild(dh);
        });

        const firstDay = new Date(year, monthIndex, 1).getDay();
        const daysInMonth = new Date(year, monthIndex + 1, 0).getDate();

        for (let i = 0; i < firstDay; i++) {
          const empty = document.createElement('div');
          grid.appendChild(empty);
        }

        const todayObj = new Date();
        const isCurrentYearAndMonth = todayObj.getFullYear() === year && todayObj.getMonth() === monthIndex;

        for (let day = 1; day <= daysInMonth; day++) {
          const cell = document.createElement('div');
          const dateStr = `${year}-${String(monthIndex + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
          
          const matchingBookings = state.bookings.filter(b => {
            if (b.inactive || !b.checkIn || !b.checkOut) return false;
            if (!isRoomInMaster(b.roomNo)) return false;

            const bIn = b.checkIn.split('T')[0];
            const bOut = b.checkOut.split('T')[0];

            return (dateStr >= bIn && dateStr <= bOut);
          });

          const isBooked = matchingBookings.length > 0;
          const isToday = isCurrentYearAndMonth && todayObj.getDate() === day;

          cell.className = `py-1 rounded text-[10px] font-bold cursor-pointer transition relative flex items-center justify-center ${
            isToday 
              ? 'ring-2 ring-indigo-600 ring-offset-1 z-10' 
              : ''
          } ${
            isBooked 
              ? 'bg-amber-400 text-slate-900 hover:bg-amber-500 shadow-xs' 
              : 'bg-white text-slate-700 hover:bg-indigo-50 hover:text-indigo-600 border border-slate-100'
          }`;

          cell.innerText = day;

          if (isBooked) {
            cell.onclick = (e) => {
              e.stopPropagation();
              showExcelCommentBox(e, dateStr, matchingBookings);
            };
          }

          grid.appendChild(cell);
        }

        monthBox.appendChild(grid);
        container.appendChild(monthBox);
      });
    }

    function showExcelCommentBox(e, dateStr, bookings) {
      const box = document.getElementById('excel-comment-box');
      const dateHeader = document.getElementById('comm-date-header');
      const listContainer = document.getElementById('comm-booking-list');

      dateHeader.innerText = formatDate(dateStr);
      listContainer.innerHTML = '';

      const now = new Date().getTime();

      bookings.forEach(b => {
        const item = document.createElement('div');
        item.className = "bg-slate-800 p-2 rounded border border-slate-700 space-y-1 hover:border-amber-400 transition cursor-pointer";
        item.onclick = () => {
          closeCommentBox();
          openBookingModal(b.id);
        };

        const cInMs = new Date(b.checkIn).getTime();
        const cOutMs = new Date(b.checkOut).getTime();

        let statusText = "Upcoming";
        let statusColorClass = "text-blue-400 font-bold";

        if (now >= cInMs && now <= cOutMs) {
          statusText = "Live";
          statusColorClass = "text-amber-400 font-bold";
        } else if (now > cOutMs) {
          statusText = "Closed";
          statusColorClass = "text-emerald-400 font-bold";
        }

        item.innerHTML = `
          <div class="flex justify-between items-center">
            <span class="font-bold text-amber-300 text-[11px]">${b.name}</span>
            <span class="bg-indigo-900 text-indigo-200 px-1.5 py-0.2 rounded text-[9px] font-mono">Room ${b.roomNo}</span>
          </div>
          <div class="text-[9px] text-slate-300">
            Check-In: ${formatDateTime(b.checkIn)}<br>
            Check-Out: ${formatDateTime(b.checkOut)}<br>
            Status: <span class="${statusColorClass}">${statusText}</span>
          </div>
          <div class="flex justify-between items-center text-[9px] pt-1 border-t border-slate-700/60">
            <span class="text-emerald-400 font-semibold">Total: ₹${b.totalAmount}</span>
            <span class="${b.totalDue > 0 ? 'text-rose-400 font-bold' : 'text-emerald-400 font-bold'}">
              ${b.totalDue > 0 ? `Due: ₹${b.totalDue}` : 'Paid'}
            </span>
          </div>
        `;
        listContainer.appendChild(item);
      });

      const rect = e.target.getBoundingClientRect();
      const scrollY = window.scrollY || window.pageYOffset;
      const scrollX = window.scrollX || window.pageXOffset;

      box.style.top = `${rect.bottom + scrollY + 5}px`;
      
      let leftPos = rect.left + scrollX - 20;
      if (leftPos + 260 > window.innerWidth) {
        leftPos = window.innerWidth - 270;
      }
      box.style.left = `${Math.max(10, leftPos)}px`;

      box.classList.remove('hidden');
    }

    function closeCommentBox() {
      const box = document.getElementById('excel-comment-box');
      if (box) box.classList.add('hidden');
    }
  </script>
</body>
</html>
