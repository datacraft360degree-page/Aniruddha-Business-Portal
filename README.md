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
    /* Custom Scrollbar for Clean UI */
    ::-webkit-scrollbar {
      width: 6px;
      height: 6px;
    }
    ::-webkit-scrollbar-track {
      background: #f1f5f9;
    }
    ::-webkit-scrollbar-thumb {
      background: #cbd5e1;
      border-radius: 4px;
    }
    ::-webkit-scrollbar-thumb:hover {
      background: #94a3b8;
    }

    /* Print-specific Styles to hide UI elements during printing */
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
        padding: 20px;
      }
      .no-print {
        display: none !important;
      }
    }

    /* Excel-style Comment Box Arrow Pin */
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
<body class="bg-slate-100 text-slate-800 font-sans min-h-screen flex flex-col relative antialiased" onclick="closeCommentBox()">

  <!-- Excel-Style Comment Box Popout for Calendar Dates -->
  <div id="excel-comment-box" onclick="event.stopPropagation()" class="excel-comment-box hidden absolute z-50 bg-slate-900 text-white text-xs rounded-xl p-3 shadow-2xl border-2 border-amber-400 space-y-2.5 w-72 transition-all duration-150">
    <!-- Comment Box Header -->
    <div class="font-bold text-amber-300 border-b border-slate-700 pb-1.5 flex justify-between items-center text-[11px]">
      <span class="flex items-center gap-1.5">
        <i class="fa-solid fa-comment-dots text-amber-400"></i>
        <span id="comm-date-header">Date Overview</span>
      </span>
      <button onclick="closeCommentBox()" class="text-slate-400 hover:text-white px-1.5 py-0.5 rounded text-xs transition">
        <i class="fa-solid fa-xmark"></i>
      </button>
    </div>
    <!-- Customer Details List Container -->
    <div id="comm-booking-list" class="space-y-2 max-h-64 overflow-y-auto pr-1">
      <!-- Populated dynamically via JavaScript -->
    </div>
  </div>

  <!-- Header & Navigation -->
  <header class="bg-indigo-700 text-white shadow-md sticky top-0 z-40 no-print">
    <div class="max-w-7xl mx-auto px-4 py-3 flex flex-col md:flex-row justify-between items-center gap-3">
      <div class="flex items-center space-x-3">
        <div class="bg-indigo-600 p-2 rounded-lg border border-indigo-500 shadow-inner">
          <i class="fa-solid fa-hotel text-xl text-indigo-100"></i>
        </div>
        <div>
          <h1 class="text-lg font-bold tracking-wide leading-tight">Business Portal</h1>
          <p class="text-[11px] text-indigo-200">Management & Booking Control System</p>
        </div>
      </div>
      
      <!-- Tab Navigation -->
      <nav class="flex space-x-1 bg-indigo-800/80 p-1 rounded-lg text-xs font-medium border border-indigo-600/50">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-3.5 py-1.5 rounded-md transition-all duration-200 active-tab">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-3.5 py-1.5 rounded-md transition-all duration-200 text-indigo-100 hover:bg-indigo-600/50">Booking Details</button>
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-3.5 py-1.5 rounded-md transition-all duration-200 text-indigo-100 hover:bg-indigo-600/50">Master Data</button>
        <button onclick="switchTab('calendar')" id="btn-calendar" class="tab-btn px-3.5 py-1.5 rounded-md transition-all duration-200 text-indigo-100 hover:bg-indigo-600/50">Calendar</button>
      </nav>

      <!-- Action Buttons -->
      <div class="flex items-center space-x-2">
        <button onclick="saveChanges()" class="bg-emerald-500 hover:bg-emerald-600 text-white px-3 py-1.5 rounded-lg text-xs font-semibold shadow-sm flex items-center gap-1.5 transition">
          <i class="fa-solid fa-floppy-disk"></i> Save Changes
        </button>
        <button onclick="exportToExcel()" class="bg-amber-500 hover:bg-amber-600 text-white px-3 py-1.5 rounded-lg text-xs font-semibold shadow-sm flex items-center gap-1.5 transition">
          <i class="fa-solid fa-file-excel"></i> Export
        </button>
      </div>
    </div>
  </header>

  <!-- Notification Toast -->
  <div id="toast" class="hidden fixed bottom-5 right-5 bg-slate-900 text-white px-4 py-2.5 rounded-lg shadow-xl z-50 flex items-center gap-2.5 transition-all duration-300 no-print border border-slate-700 text-xs">
    <i class="fa-solid fa-circle-check text-emerald-400 text-base"></i>
    <span id="toast-message" class="font-medium">Changes saved successfully!</span>
  </div>

  <!-- Main Content Area -->
  <main class="max-w-7xl mx-auto px-4 py-5 flex-1 w-full no-print space-y-5">

    <!-- DASHBOARD TAB -->
    <section id="tab-dashboard" class="tab-content space-y-5">
      <!-- Welcome Banner -->
      <div class="bg-gradient-to-r from-indigo-700 via-indigo-600 to-blue-600 rounded-xl p-5 text-white shadow-md flex justify-between items-center">
        <div>
          <h2 class="text-xl font-bold tracking-tight">Hi Aniruddha, Welcome to your Business Portal 🤝</h2>
          <p class="text-indigo-100 text-xs mt-1">View, schedule, and organize room allocations and customer records.</p>
        </div>
      </div>

      <!-- Quick Summary Cards -->
      <div class="grid grid-cols-2 lg:grid-cols-4 gap-4">
        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Total Bookings</p>
            <p id="dash-total-bookings" class="text-2xl font-black text-slate-800 mt-0.5">0</p>
          </div>
          <div class="p-3 bg-blue-50 text-blue-600 rounded-xl"><i class="fa-solid fa-bookmark text-lg"></i></div>
        </div>
        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Booking Amount</p>
            <p id="dash-total-amount" class="text-2xl font-black text-slate-800 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-indigo-50 text-indigo-600 rounded-xl"><i class="fa-solid fa-receipt text-lg"></i></div>
        </div>
        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Advance Received</p>
            <p id="dash-advanced" class="text-2xl font-black text-emerald-600 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-emerald-50 text-emerald-600 rounded-xl"><i class="fa-solid fa-wallet text-lg"></i></div>
        </div>
        <div class="bg-white p-4 rounded-xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-[10px] uppercase font-bold text-slate-400 tracking-wider">Total Due Amount</p>
            <p id="dash-due" class="text-2xl font-black text-rose-600 mt-0.5">₹0</p>
          </div>
          <div class="p-3 bg-rose-50 text-rose-600 rounded-xl"><i class="fa-solid fa-hand-holding-dollar text-lg"></i></div>
        </div>
      </div>

      <!-- Years Grid -->
      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
        <div class="mb-4">
          <h3 class="text-sm font-bold text-slate-800 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-days text-indigo-600"></i> Active Years Directory (2026 – 2085)
          </h3>
          <p class="text-[11px] text-slate-400 mt-0.5">Select a year to jump directly into its full calendar overview.</p>
        </div>
        <div id="years-grid" class="grid grid-cols-6 sm:grid-cols-10 md:grid-cols-12 gap-2">
          <!-- Populated by JS -->
        </div>
      </div>
    </section>

    <!-- BOOKING DETAILS TAB -->
    <section id="tab-booking" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
        <div class="flex justify-between items-center mb-4 pb-3 border-b border-slate-100">
          <div>
            <h2 class="text-base font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-solid fa-address-card text-indigo-600"></i> Customer & Reservation Directory
            </h2>
            <p class="text-[11px] text-slate-400 mt-0.5">Review check-ins, tariffs, and due balances.</p>
          </div>
          <button onclick="openBookingModal()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-3.5 py-1.5 rounded-lg text-xs font-semibold shadow-sm flex items-center gap-1.5 transition">
            <i class="fa-solid fa-plus"></i> Add Booking
          </button>
        </div>

        <!-- Bookings Table View -->
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                <th class="py-2.5 px-3">Guest Name</th>
                <th class="py-2.5 px-3">Contact / ID</th>
                <th class="py-2.5 px-3">Room</th>
                <th class="py-2.5 px-3">Agent</th>
                <th class="py-2.5 px-3 min-w-[170px]">Stay Window</th>
                <th class="py-2.5 px-3">Tariff</th>
                <th class="py-2.5 px-3">Payment</th>
                <th class="py-2.5 px-3">Balance</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="bookings-tbody" class="divide-y divide-slate-100 text-xs">
              <!-- Populated by JS -->
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- MASTER DATA TAB -->
    <section id="tab-master" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
        <div class="flex justify-between items-center mb-4 pb-3 border-b border-slate-100">
          <div>
            <h2 class="text-base font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-solid fa-users-gear text-indigo-600"></i> Master Agent & Room Directory
            </h2>
            <p class="text-[11px] text-slate-400 mt-0.5">Configure room capacities and default agent details.</p>
          </div>
          <button onclick="addMasterRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition shadow-sm">
            <i class="fa-solid fa-plus"></i> Add Entry
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-[11px] font-bold text-slate-500 uppercase tracking-wider">
                <th class="py-2.5 px-3">Agent Name</th>
                <th class="py-2.5 px-3">Agent Contact</th>
                <th class="py-2.5 px-3">Room No</th>
                <th class="py-2.5 px-3">Capacity</th>
                <th class="py-2.5 px-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="master-tbody" class="divide-y divide-slate-100 text-xs">
              <!-- Populated by JS -->
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- CALENDAR TAB -->
    <section id="tab-calendar" class="tab-content hidden space-y-4">
      <div class="bg-white rounded-xl shadow-sm border border-slate-200 p-5">
        <div class="flex justify-between items-center mb-5">
          <div>
            <h2 class="text-base font-bold text-slate-800 flex items-center gap-1.5">
              <i class="fa-regular fa-calendar-check text-indigo-600"></i> Year Overview Calendar
            </h2>
            <p class="text-[11px] text-slate-400 mt-0.5">Click any highlighted date to open an Excel-style comment popout with booking details.</p>
          </div>
          
          <!-- Calendar Year Dropdown -->
          <div class="flex items-center space-x-2 bg-slate-50 p-1.5 rounded-lg border border-slate-200">
            <label for="cal-year-select" class="text-[10px] font-bold text-slate-500 uppercase tracking-wider pl-1">Year:</label>
            <select id="cal-year-select" onchange="renderCalendar(parseInt(this.value))" class="bg-white border border-slate-300 text-indigo-700 font-bold px-2.5 py-1 rounded-md focus:outline-none cursor-pointer text-xs">
              <!-- Populated by JS -->
            </select>
          </div>
        </div>

        <div id="calendar-container" class="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 xl:grid-cols-4 gap-4">
          <!-- Calendar months rendered dynamically -->
        </div>
      </div>
    </section>

  </main>

  <!-- ADD / EDIT BOOKING MODAL -->
  <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto no-print">
    <div class="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-2xl w-full p-6 space-y-4 my-6">
      <div class="flex justify-between items-center pb-3 border-b border-slate-100">
        <div>
          <h3 id="modal-title" class="text-base font-bold text-slate-800 flex items-center gap-1.5">
            <i class="fa-solid fa-calendar-plus text-indigo-600"></i> Add New Booking
          </h3>
          <p class="text-[11px] text-slate-400">Fill out guest, room allocation, and stay details.</p>
        </div>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 p-1 text-lg"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="booking-form" onsubmit="handleSaveBooking(event)" class="space-y-4 text-xs">
        <input type="hidden" id="modal-booking-id">

        <!-- Guest Details Box -->
        <div class="bg-slate-50 p-3.5 rounded-xl border border-slate-200 space-y-3">
          <h4 class="text-[10px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1">
            <i class="fa-solid fa-user-tag text-indigo-500"></i> Guest Information
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Customer Name</label>
              <input type="text" id="cust-name" required class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Address</label>
              <input type="text" id="cust-address" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">ID Number</label>
              <input type="text" id="cust-id" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Contact No</label>
              <input type="text" id="cust-contact" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
          </div>
        </div>

        <!-- Room & Stay Schedule Box -->
        <div class="bg-slate-50 p-3.5 rounded-xl border border-slate-200 space-y-3">
          <h4 class="text-[10px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1">
            <i class="fa-solid fa-bed text-indigo-500"></i> Room Selection & Stay Dates
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-3 gap-3">
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Room No</label>
              <select id="cust-room" onchange="autoCaptureRoomDetails()" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-bold text-indigo-700">
                <!-- Populated from Master -->
              </select>
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Agent Info <span class="text-[10px] text-indigo-500 font-normal">(Auto)</span></label>
              <input type="text" id="cust-agent" readonly class="w-full bg-slate-200/70 border border-slate-300 text-slate-600 font-medium rounded-lg px-2.5 py-1.5 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Capacity <span class="text-[10px] text-indigo-500 font-normal">(Auto)</span></label>
              <input type="text" id="cust-capacity" readonly class="w-full bg-slate-200/70 border border-slate-300 text-slate-600 font-medium rounded-lg px-2.5 py-1.5 cursor-not-allowed">
            </div>

            <div class="sm:col-span-[1.5]">
              <label class="block font-semibold text-slate-600 mb-1">Check In Date & Time</label>
              <input type="datetime-local" id="cust-checkin" onchange="calculateModalBilling()" required class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div class="sm:col-span-[1.5]">
              <label class="block font-semibold text-slate-600 mb-1">Check Out Date & Time</label>
              <input type="datetime-local" id="cust-checkout" onchange="calculateModalBilling()" required class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
          </div>
        </div>

        <!-- Billing Calculation Box -->
        <div class="bg-indigo-50/50 p-3.5 rounded-xl border border-indigo-100 space-y-3">
          <h4 class="text-[10px] font-bold uppercase tracking-wider text-indigo-700 flex items-center gap-1">
            <i class="fa-solid fa-calculator text-indigo-600"></i> Billing Summary
          </h4>
          <div class="grid grid-cols-5 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Days</label>
              <input type="number" id="cust-days" readonly class="w-full bg-slate-200/80 font-bold text-slate-700 border border-slate-300 rounded-lg px-2 py-1.5 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Price/Day (₹)</label>
              <input type="number" id="cust-price" value="1200" oninput="calculateModalBilling()" class="w-full bg-white font-bold text-slate-700 border border-slate-300 rounded-lg px-2 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Total (₹)</label>
              <input type="number" id="cust-total" readonly class="w-full bg-slate-200/80 text-indigo-700 font-bold border border-slate-300 rounded-lg px-2 py-1.5 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Advance (₹)</label>
              <input type="number" id="cust-advance" value="0" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-300 rounded-lg px-2 py-1.5 focus:outline-none focus:ring-1 focus:ring-indigo-500 font-semibold text-emerald-600">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-1">Due (₹)</label>
              <input type="number" id="cust-due" readonly class="w-full bg-slate-200/80 text-rose-700 font-bold border border-slate-300 rounded-lg px-2 py-1.5 cursor-not-allowed">
            </div>
          </div>
        </div>

        <div class="flex justify-end space-x-2 pt-2">
          <button type="button" onclick="closeBookingModal()" class="px-4 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg font-semibold transition">Cancel</button>
          <button type="submit" class="px-5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg font-semibold shadow transition">Save Booking</button>
        </div>
      </form>
    </div>
  </div>

  <!-- PRINTABLE INVOICE MODAL -->
  <div id="invoice-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
    <div class="bg-white rounded-xl shadow-2xl border border-slate-200 max-w-2xl w-full p-8 space-y-5 relative" id="printable-invoice">
      <!-- Invoice Header -->
      <div class="flex justify-between items-start border-b border-slate-200 pb-4">
        <div>
          <h2 class="text-xl font-black text-indigo-700 uppercase tracking-wide">Business Portal Hotel</h2>
          <p class="text-[11px] text-slate-500 mt-0.5">123 Business Street, Suite 400</p>
          <p class="text-[11px] text-slate-500">Phone: +91 98765 43210 | Email: info@businessportal.com</p>
        </div>
        <div class="text-right">
          <span class="inline-block bg-indigo-100 text-indigo-800 text-[10px] font-bold px-2.5 py-0.5 rounded-full uppercase mb-1">Booking Receipt</span>
          <p class="text-[11px] text-slate-500">Invoice ID: <strong id="inv-id" class="text-slate-800 font-mono">INV-A-0000001</strong></p>
          <p class="text-[11px] text-slate-500">Issued On: <strong id="inv-date" class="text-slate-800"></strong></p>
        </div>
      </div>

      <!-- Guest & Stay Details -->
      <div class="grid grid-cols-2 gap-6 bg-slate-50 p-4 rounded-xl border border-slate-100 text-xs">
        <div>
          <h4 class="font-bold text-slate-400 uppercase text-[10px] tracking-wider mb-1">Guest Information</h4>
          <p class="text-slate-800 font-semibold" id="inv-guest-name">-</p>
          <p class="text-slate-600 mt-0.5" id="inv-guest-address">Address: -</p>
          <p class="text-slate-600" id="inv-guest-contact">Contact: -</p>
          <p class="text-slate-600" id="inv-guest-id">ID No: -</p>
        </div>
        <div>
          <h4 class="font-bold text-slate-400 uppercase text-[10px] tracking-wider mb-1">Reservation Info</h4>
          <p class="text-slate-800 font-semibold" id="inv-room">Room No: -</p>
          <p class="text-slate-600 mt-0.5" id="inv-agent">Agent: -</p>
          <p class="text-slate-600" id="inv-checkin">Check-in: -</p>
          <p class="text-slate-600" id="inv-checkout">Check-out: -</p>
        </div>
      </div>

      <!-- Invoice Items Table -->
      <div class="overflow-x-auto">
        <table class="w-full text-left text-xs">
          <thead>
            <tr class="bg-indigo-50 text-indigo-900 border-b border-indigo-100">
              <th class="p-2.5">Description</th>
              <th class="p-2.5 text-center">Duration</th>
              <th class="p-2.5 text-right">Rate / Day</th>
              <th class="p-2.5 text-right">Total Amount</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100">
            <tr>
              <td class="p-2.5 font-semibold text-slate-800" id="inv-item-desc">Room Accommodation</td>
              <td class="p-2.5 text-center" id="inv-item-days">0 Days</td>
              <td class="p-2.5 text-right" id="inv-item-rate">₹0</td>
              <td class="p-2.5 text-right font-semibold text-slate-800" id="inv-item-total">₹0</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Payment Summary -->
      <div class="flex justify-end pt-1 border-t border-slate-200">
        <div class="w-1/2 space-y-1 text-xs">
          <div class="flex justify-between text-slate-600">
            <span>Total Amount:</span>
            <strong id="inv-sum-total" class="text-slate-800">₹0</strong>
          </div>
          <div class="flex justify-between text-emerald-600">
            <span>Advance Payment:</span>
            <strong id="inv-sum-advance">₹0</strong>
          </div>
          <div class="flex justify-between text-rose-600 text-xs font-bold border-t border-slate-200 pt-1">
            <span>Balance Due:</span>
            <span id="inv-sum-due">₹0</span>
          </div>
        </div>
      </div>

      <!-- Footer & Signature -->
      <div class="pt-6 border-t border-slate-200 flex justify-between items-end text-[11px] text-slate-400">
        <div>
          <p class="font-bold text-slate-600">Thank you for stay with us!</p>
          <p>For inquiries, please contact hotel management.</p>
        </div>
        <div class="text-center border-t border-slate-300 pt-1 w-32">
          <p class="font-semibold text-slate-600">Authorized Signature</p>
        </div>
      </div>

      <!-- Modal Actions -->
      <div class="flex justify-end space-x-2 pt-3 no-print border-t border-slate-100">
        <button type="button" onclick="closeInvoiceModal()" class="px-4 py-1.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-lg text-xs font-semibold transition">Close</button>
        <button type="button" onclick="window.print()" class="px-5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg text-xs font-semibold shadow flex items-center gap-1.5 transition">
          <i class="fa-solid fa-print"></i> Print Invoice
        </button>
      </div>
    </div>
  </div>

  <script>
    // App Data State
    let state = {
      bookings: [
        {
          id: "bk_1",
          name: "Kapil",
          address: "Kolkata",
          idNo: "2578 0001 6658",
          contactNo: "1234567890",
          roomNo: 2,
          agentInfo: "A2 1234567890",
          capacity: "2 Person",
          checkIn: "2026-07-22T11:00",
          checkOut: "2026-07-24T11:00",
          noOfDays: 2,
          perDayPrice: 1200,
          totalAmount: 2400,
          advanced: 1000,
          totalDue: 1400
        },
        {
          id: "bk_2",
          name: "Aniruddha",
          address: "Mumbai",
          idNo: "9988 7766 5544",
          contactNo: "9876543210",
          roomNo: 4,
          agentInfo: "A1 1234567890",
          capacity: "4 Person",
          checkIn: "2026-07-22T14:00",
          checkOut: "2026-07-25T10:00",
          noOfDays: 3,
          perDayPrice: 1500,
          totalAmount: 4500,
          advanced: 2000,
          totalDue: 2500
        }
      ],
      master: [
        { agentName: "A1", phone: "1234567890", roomNo: 1, capacity: 4 },
        { agentName: "A2", phone: "1234567890", roomNo: 2, capacity: 2 },
        { agentName: "A3", phone: "1234567890", roomNo: 3, capacity: 4 },
        { agentName: "A1", phone: "1234567890", roomNo: 4, capacity: 4 },
        { agentName: "A3", phone: "1234567890", roomNo: 5, capacity: 4 }
      ],
      selectedYear: 2026
    };

    function generateInvoiceNumber(index) {
      const numStr = String(index + 1).padStart(7, '0');
      return `INV-A-${numStr}`;
    }

    function loadSavedData() {
      const saved = localStorage.getItem('webapp_data');
      if (saved) {
        try { 
          const parsed = JSON.parse(saved);
          if (parsed.bookings) state = parsed;
        } catch(e){}
      }
    }

    document.addEventListener("DOMContentLoaded", () => {
      loadSavedData();
      initDashboard();
      populateRoomDropdown();
      populateCalendarYearDropdown();
      renderBookingsTable();
      renderMasterTable();
      renderCalendar(state.selectedYear);
    });

    function populateCalendarYearDropdown() {
      const yearSelect = document.getElementById('cal-year-select');
      yearSelect.innerHTML = '';
      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = `Year ${y}`;
        if (y === state.selectedYear) {
          opt.selected = true;
        }
        yearSelect.appendChild(opt);
      }
    }

    function populateRoomDropdown() {
      const roomSelect = document.getElementById('cust-room');
      roomSelect.innerHTML = '';
      
      state.master.forEach(m => {
        const opt = document.createElement('option');
        opt.value = m.roomNo;
        opt.text = `Room ${m.roomNo}`;
        roomSelect.appendChild(opt);
      });

      autoCaptureRoomDetails();
    }

    function autoCaptureRoomDetails() {
      const selectedRoom = document.getElementById('cust-room').value;
      const matched = state.master.find(m => parseInt(m.roomNo) === parseInt(selectedRoom));

      if (matched) {
        document.getElementById('cust-agent').value = `${matched.agentName} ${matched.phone}`;
        document.getElementById('cust-capacity').value = `${matched.capacity} Person`;
      }
    }

    function switchTab(tabId) {
      document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
      document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active-tab', 'bg-indigo-600', 'text-white');
        btn.classList.add('text-indigo-100', 'hover:bg-indigo-600/50');
      });

      document.getElementById(`tab-${tabId}`).classList.remove('hidden');
      const activeBtn = document.getElementById(`btn-${tabId}`);
      activeBtn.classList.add('active-tab', 'bg-indigo-600', 'text-white');
      closeCommentBox();
    }

    function selectDashboardYear(year) {
      state.selectedYear = year;
      initDashboard();
      renderCalendar(year);
      switchTab('calendar');
    }

    function initDashboard() {
      const grid = document.getElementById('years-grid');
      grid.innerHTML = '';
      for (let y = 2026; y <= 2085; y++) {
        const item = document.createElement('div');
        item.className = `text-center py-1.5 px-1 rounded-lg text-[11px] font-bold border cursor-pointer transition ${
          y === state.selectedYear 
            ? 'bg-indigo-600 text-white border-indigo-600 shadow-sm' 
            : 'bg-white text-slate-600 border-slate-200 hover:bg-indigo-50 hover:text-indigo-600'
        }`;
        item.innerText = y;
        item.onclick = () => selectDashboardYear(y);
        grid.appendChild(item);
      }
      updateDashboardCards();
    }

    function updateDashboardCards() {
      const totalBookings = state.bookings.length;
      const totalAmt = state.bookings.reduce((sum, b) => sum + (b.totalAmount || 0), 0);
      const totalAdv = state.bookings.reduce((sum, b) => sum + (b.advanced || 0), 0);
      const totalDue = state.bookings.reduce((sum, b) => sum + (b.totalDue || 0), 0);

      document.getElementById('dash-total-bookings').innerText = totalBookings;
      document.getElementById('dash-total-amount').innerText = `₹${totalAmt.toLocaleString('en-IN')}`;
      document.getElementById('dash-advanced').innerText = `₹${totalAdv.toLocaleString('en-IN')}`;
      document.getElementById('dash-due').innerText = `₹${totalDue.toLocaleString('en-IN')}`;
    }

    function printInvoice(bookingId) {
      const bIndex = state.bookings.findIndex(item => item.id === bookingId);
      if (bIndex === -1) return;

      const b = state.bookings[bIndex];
      const invoiceNo = generateInvoiceNumber(bIndex);

      const today = new Date().toLocaleDateString('en-IN', { year: 'numeric', month: 'short', day: 'numeric' });
      document.getElementById('inv-id').innerText = invoiceNo;
      document.getElementById('inv-date').innerText = today;

      document.getElementById('inv-guest-name').innerText = b.name || 'N/A';
      document.getElementById('inv-guest-address').innerText = `Address: ${b.address || 'N/A'}`;
      document.getElementById('inv-guest-contact').innerText = `Contact: ${b.contactNo || 'N/A'}`;
      document.getElementById('inv-guest-id').innerText = `ID No: ${b.idNo || 'N/A'}`;

      document.getElementById('inv-room').innerText = `Room No: ${b.roomNo} (${b.capacity || 'Standard'})`;
      document.getElementById('inv-agent').innerText = `Agent: ${b.agentInfo || 'Direct'}`;
      document.getElementById('inv-checkin').innerText = `Check-in: ${b.checkIn ? b.checkIn.replace('T', ' ') : '-'}`;
      document.getElementById('inv-checkout').innerText = `Check-out: ${b.checkOut ? b.checkOut.replace('T', ' ') : '-'}`;

      document.getElementById('inv-item-desc').innerText = `Room ${b.roomNo} Reservation Charges`;
      document.getElementById('inv-item-days').innerText = `${b.noOfDays} Days`;
      document.getElementById('inv-item-rate').innerText = `₹${(b.perDayPrice || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-item-total').innerText = `₹${(b.totalAmount || 0).toLocaleString('en-IN')}`;

      document.getElementById('inv-sum-total').innerText = `₹${(b.totalAmount || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-advance').innerText = `₹${(b.advanced || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-due').innerText = `₹${(b.totalDue || 0).toLocaleString('en-IN')}`;

      document.getElementById('invoice-modal').classList.remove('hidden');
    }

    function closeInvoiceModal() {
      document.getElementById('invoice-modal').classList.add('hidden');
    }

    function openBookingModal(bookingId = null) {
      populateRoomDropdown();
      const form = document.getElementById('booking-form');
      form.reset();

      if (bookingId) {
        const b = state.bookings.find(item => item.id === bookingId);
        if (b) {
          document.getElementById('modal-title').innerText = 'Edit Booking Details';
          document.getElementById('modal-booking-id').value = b.id;
          document.getElementById('cust-name').value = b.name;
          document.getElementById('cust-address').value = b.address;
          document.getElementById('cust-id').value = b.idNo;
          document.getElementById('cust-contact').value = b.contactNo;
          document.getElementById('cust-room').value = b.roomNo;
          autoCaptureRoomDetails();
          document.getElementById('cust-checkin').value = b.checkIn;
          document.getElementById('cust-checkout').value = b.checkOut;
          document.getElementById('cust-price').value = b.perDayPrice;
          document.getElementById('cust-advance').value = b.advanced;
          calculateModalBilling();
        }
      } else {
        document.getElementById('modal-title').innerText = 'Add New Booking';
        document.getElementById('modal-booking-id').value = '';
        document.getElementById('cust-price').value = 1200;
        document.getElementById('cust-advance').value = 0;
        calculateModalBilling();
      }

      document.getElementById('booking-modal').classList.remove('hidden');
    }

    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }

    function calculateModalBilling() {
      const checkInVal = document.getElementById('cust-checkin').value;
      const checkOutVal = document.getElementById('cust-checkout').value;
      
      let days = 0;
      if (checkInVal && checkOutVal) {
        const d1 = new Date(checkInVal);
        const d2 = new Date(checkOutVal);
        const diff = d2 - d1;
        days = Math.max(0, Math.ceil(diff / (1000 * 60 * 60 * 24)));
      }

      const price = parseFloat(document.getElementById('cust-price').value) || 0;
      const advance = parseFloat(document.getElementById('cust-advance').value) || 0;
      const total = days * price;
      const due = total - advance;

      document.getElementById('cust-days').value = days;
      document.getElementById('cust-total').value = total;
      document.getElementById('cust-due').value = due;
    }

    function handleSaveBooking(e) {
      e.preventDefault();
      const id = document.getElementById('modal-booking-id').value;
      const roomNo = parseInt(document.getElementById('cust-room').value);
      const checkIn = document.getElementById('cust-checkin').value;
      const checkOut = document.getElementById('cust-checkout').value;

      const newIn = new Date(checkIn).getTime();
      const newOut = new Date(checkOut).getTime();

      if (newIn >= newOut) {
        alert("Check-Out date & time must be strictly after Check-In date & time.");
        return;
      }

      const conflict = state.bookings.find(b => {
        if (id && b.id === id) return false;
        if (parseInt(b.roomNo) !== roomNo) return false;

        const existingIn = new Date(b.checkIn).getTime();
        const existingOut = new Date(b.checkOut).getTime();

        return (newIn < existingOut && newOut > existingIn);
      });

      if (conflict) {
        const confOutFormatted = conflict.checkOut.replace('T', ' ');
        alert(`❌ Booking Conflict Alert!\n\nRoom ${roomNo} is already occupied by ${conflict.name} until ${confOutFormatted}.\n\nPlease select a check-in time after ${confOutFormatted} or assign a different room.`);
        return;
      }
      
      const newBooking = {
        id: id || `bk_${Date.now()}`,
        name: document.getElementById('cust-name').value,
        address: document.getElementById('cust-address').value,
        idNo: document.getElementById('cust-id').value,
        contactNo: document.getElementById('cust-contact').value,
        roomNo: roomNo,
        agentInfo: document.getElementById('cust-agent').value,
        capacity: document.getElementById('cust-capacity').value,
        checkIn: checkIn,
        checkOut: checkOut,
        noOfDays: parseInt(document.getElementById('cust-days').value) || 0,
        perDayPrice: parseFloat(document.getElementById('cust-price').value) || 0,
        totalAmount: parseFloat(document.getElementById('cust-total').value) || 0,
        advanced: parseFloat(document.getElementById('cust-advance').value) || 0,
        totalDue: parseFloat(document.getElementById('cust-due').value) || 0
      };

      if (id) {
        const idx = state.bookings.findIndex(b => b.id === id);
        if (idx !== -1) state.bookings[idx] = newBooking;
      } else {
        state.bookings.push(newBooking);
      }

      closeBookingModal();
      renderBookingsTable();
      updateDashboardCards();
      renderCalendar(state.selectedYear);
      saveChanges(false);
    }

    function deleteBooking(id) {
      if (confirm('Are you sure you want to delete this booking entry?')) {
        state.bookings = state.bookings.filter(b => b.id !== id);
        renderBookingsTable();
        updateDashboardCards();
        renderCalendar(state.selectedYear);
        saveChanges(false);
      }
    }

    function renderBookingsTable() {
      const tbody = document.getElementById('bookings-tbody');
      tbody.innerHTML = '';

      if (state.bookings.length === 0) {
        tbody.innerHTML = `<tr><td colspan="9" class="text-center py-8 text-slate-400">No bookings found. Click "Add Booking" to register guests.</td></tr>`;
        return;
      }

      state.bookings.forEach((b) => {
        const checkInFmt = b.checkIn ? b.checkIn.replace('T', ' ') : '-';
        const checkOutFmt = b.checkOut ? b.checkOut.replace('T', ' ') : '-';
        
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2.5 px-3 font-bold text-slate-800">${b.name}</td>
          <td class="py-2.5 px-3 text-[11px] text-slate-500">
            <div class="font-medium text-slate-700">${b.contactNo}</div>
            <div class="text-slate-400">${b.idNo}</div>
          </td>
          <td class="py-2.5 px-3">
            <span class="bg-indigo-50 text-indigo-700 font-bold px-2 py-0.5 rounded text-[11px] inline-block">Room ${b.roomNo}</span>
            <div class="text-[10px] text-slate-400 mt-0.5">${b.capacity}</div>
          </td>
          <td class="py-2.5 px-3 text-[11px] text-slate-600 font-medium">${b.agentInfo}</td>
          <td class="py-2.5 px-3 text-[11px]">
            <div class="text-emerald-700 font-medium"><i class="fa-solid fa-plane-arrival mr-1"></i> ${checkInFmt}</div>
            <div class="text-rose-600 font-medium mt-0.5"><i class="fa-solid fa-plane-departure mr-1"></i> ${checkOutFmt}</div>
          </td>
          <td class="py-2.5 px-3 text-[11px] font-semibold text-slate-700">${b.noOfDays} d × ₹${b.perDayPrice}</td>
          <td class="py-2.5 px-3 text-[11px]">
            <div class="font-bold text-slate-800">Tot: ₹${b.totalAmount}</div>
            <div class="text-emerald-600 font-medium">Adv: ₹${b.advanced}</div>
          </td>
          <td class="py-2.5 px-3">
            <span class="px-2 py-0.5 rounded-full text-[11px] font-bold inline-block ${b.totalDue > 0 ? 'bg-rose-100 text-rose-700' : 'bg-emerald-100 text-emerald-700'}">
              ₹${b.totalDue} Due
            </span>
          </td>
          <td class="py-2.5 px-3 text-center">
            <div class="flex items-center justify-center space-x-2">
              <button onclick="printInvoice('${b.id}')" title="Print Invoice" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-2 py-1 rounded text-[11px] font-semibold transition flex items-center gap-1 border border-slate-200">
                <i class="fa-solid fa-print text-indigo-600"></i> Inv
              </button>
              <button onclick="openBookingModal('${b.id}')" title="Edit Booking" class="text-indigo-600 hover:text-indigo-800 p-1"><i class="fa-solid fa-pen-to-square"></i></button>
              <button onclick="deleteBooking('${b.id}')" title="Delete Booking" class="text-rose-500 hover:text-rose-700 p-1"><i class="fa-solid fa-trash-can"></i></button>
            </div>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function renderMasterTable() {
      const tbody = document.getElementById('master-tbody');
      tbody.innerHTML = '';
      state.master.forEach((row, index) => {
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 transition";
        tr.innerHTML = `
          <td class="py-2 px-3"><input type="text" value="${row.agentName}" onchange="updateMasterRow(${index}, 'agentName', this.value)" class="w-full bg-white border border-slate-300 rounded-md px-2.5 py-1 focus:border-indigo-500 focus:outline-none text-xs"></td>
          <td class="py-2 px-3"><input type="text" value="${row.phone}" onchange="updateMasterRow(${index}, 'phone', this.value)" class="w-full bg-white border border-slate-300 rounded-md px-2.5 py-1 focus:border-indigo-500 focus:outline-none text-xs"></td>
          <td class="py-2 px-3"><input type="number" value="${row.roomNo}" onchange="updateMasterRow(${index}, 'roomNo', this.value)" class="w-full bg-white border border-slate-300 rounded-md px-2.5 py-1 focus:border-indigo-500 focus:outline-none text-xs font-semibold text-indigo-700"></td>
          <td class="py-2 px-3"><input type="number" value="${row.capacity}" onchange="updateMasterRow(${index}, 'capacity', this.value)" class="w-full bg-white border border-slate-300 rounded-md px-2.5 py-1 focus:border-indigo-500 focus:outline-none text-xs"></td>
          <td class="py-2 px-3 text-center">
            <button onclick="removeMasterRow(${index})" class="text-rose-500 hover:text-rose-700 p-1"><i class="fa-solid fa-trash-can"></i></button>
          </td>
        `;
        tbody.appendChild(tr);
      });
    }

    function updateMasterRow(index, key, value) {
      state.master[index][key] = value;
      populateRoomDropdown();
    }

    function addMasterRow() {
      state.master.push({ agentName: "New Agent", phone: "1234567890", roomNo: state.master.length + 1, capacity: 2 });
      renderMasterTable();
      populateRoomDropdown();
    }

    function removeMasterRow(index) {
      state.master.splice(index, 1);
      renderMasterTable();
      populateRoomDropdown();
    }

    // CALENDAR RENDERER
    function renderCalendar(year) {
      state.selectedYear = year;
      const calSelect = document.getElementById('cal-year-select');
      if (calSelect) {
        calSelect.value = year;
      }

      const container = document.getElementById('calendar-container');
      container.innerHTML = '';

      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

      months.forEach((month, monthIdx) => {
        const monthCard = document.createElement('div');
        monthCard.className = "bg-white border border-slate-200 rounded-xl p-3 shadow-sm";
        
        let header = `<h4 class="font-extrabold text-slate-800 text-center mb-2 text-xs border-b border-slate-100 pb-1.5">${month} ${year}</h4>`;
        let daysHeader = `<div class="grid grid-cols-7 text-center text-[10px] font-bold text-slate-400 mb-1"><div>S</div><div>M</div><div>T</div><div>W</div><div>T</div><div>F</div><div>S</div></div>`;
        
        let firstDay = new Date(year, monthIdx, 1).getDay();
        let totalDays = new Date(year, monthIdx + 1, 0).getDate();
        
        let daysGrid = `<div class="grid grid-cols-7 text-center text-xs gap-1">`;
        for (let i = 0; i < firstDay; i++) {
          daysGrid += `<div></div>`;
        }

        for (let d = 1; d <= totalDays; d++) {
          const currDate = new Date(year, monthIdx, d);
          
          let dayBookings = [];

          for (let b of state.bookings) {
            if (!b.checkIn || !b.checkOut) continue;
            const cIn = new Date(b.checkIn);
            const cOut = new Date(b.checkOut);

            const checkInStart = new Date(cIn.getFullYear(), cIn.getMonth(), cIn.getDate());
            const checkOutEnd = new Date(cOut.getFullYear(), cOut.getMonth(), cOut.getDate());

            if (currDate >= checkInStart && currDate <= checkOutEnd) {
              let statusText = "Reserved Stay";
              if (currDate.getTime() === checkInStart.getTime()) statusText = "Check-in Date";
              else if (currDate.getTime() === checkOutEnd.getTime()) statusText = "Check-out Date";

              dayBookings.push({
                booking: b,
                statusText: statusText
              });
            }
          }

          if (dayBookings.length > 0) {
            const bookingsJson = encodeURIComponent(JSON.stringify(dayBookings));
            const formattedDateStr = `${month} ${d}, ${year}`;
            
            const badgeBg = dayBookings.length > 1 
              ? 'bg-amber-600 text-white font-black hover:bg-amber-700' 
              : 'bg-indigo-600 text-white font-bold hover:bg-indigo-700';

            daysGrid += `<div 
              onclick="toggleCommentBox(event, this, '${formattedDateStr}', '${bookingsJson}')" 
              class="relative py-1 ${badgeBg} rounded cursor-pointer shadow-sm hover:scale-105 transition-transform flex flex-col items-center justify-center">
                <span>${d}</span>
                <!-- Excel Comment Indicator Flag -->
                <span class="absolute top-0 right-0 w-0 h-0 border-t-[5px] border-r-[5px] border-t-amber-300 border-r-transparent rounded-tr-sm"></span>
                ${dayBookings.length > 1 ? `<span class="text-[9px] leading-none bg-amber-900/80 text-amber-200 px-1 py-0.5 rounded-full mt-0.5 font-bold">${dayBookings.length} Guests</span>` : ''}
              </div>`;
          } else {
            daysGrid += `<div class="py-1 hover:bg-slate-100 text-slate-700 cursor-pointer rounded font-medium">${d}</div>`;
          }
        }
        daysGrid += `</div>`;

        monthCard.innerHTML = header + daysHeader + daysGrid;
        container.appendChild(monthCard);
      });
    }

    // EXCEL-STYLE PINNED COMMENT BOX LOGIC
    function toggleCommentBox(e, targetElem, dateStr, encodedBookings) {
      e.stopPropagation();
      
      const commBox = document.getElementById('excel-comment-box');
      const container = document.getElementById('comm-booking-list');
      const dateHeader = document.getElementById('comm-date-header');
      
      const dayBookings = JSON.parse(decodeURIComponent(encodedBookings));

      dateHeader.innerText = `${dateStr} (${dayBookings.length} ${dayBookings.length > 1 ? 'Bookings' : 'Booking'})`;
      container.innerHTML = '';

      dayBookings.forEach((item) => {
        const b = item.booking;
        const card = document.createElement('div');
        card.className = "bg-slate-800/95 border border-slate-700/80 p-2.5 rounded-lg space-y-1 text-xs shadow-inner";
        
        card.innerHTML = `
          <div class="flex justify-between items-center text-amber-300 font-bold border-b border-slate-700/60 pb-1">
            <span class="truncate max-w-[130px]"><i class="fa-solid fa-user text-[10px] mr-1 text-amber-400"></i> ${b.name || 'Guest'}</span>
            <span class="bg-indigo-900/90 text-indigo-200 text-[10px] px-1.5 py-0.5 rounded font-mono border border-indigo-700">Room ${b.roomNo}</span>
          </div>
          <p><span class="text-slate-400">Contact:</span> <strong class="text-slate-200 font-medium">${b.contactNo || 'N/A'}</strong></p>
          <p><span class="text-slate-400">Status:</span> <strong class="text-emerald-400 font-medium">${item.statusText}</strong></p>
          <p><span class="text-slate-400">Check-In:</span> <span class="text-slate-300 font-mono text-[11px]">${b.checkIn ? b.checkIn.replace('T', ' ') : '-'}</span></p>
          <p><span class="text-slate-400">Check-Out:</span> <span class="text-slate-300 font-mono text-[11px]">${b.checkOut ? b.checkOut.replace('T', ' ') : '-'}</span></p>
        `;

        container.appendChild(card);
      });

      // Pin comment box relative to the clicked date element
      const rect = targetElem.getBoundingClientRect();
      const scrollY = window.scrollY || document.documentElement.scrollTop;
      const scrollX = window.scrollX || document.documentElement.scrollLeft;

      commBox.style.top = `${rect.bottom + scrollY + 6}px`;
      
      // Keep box inside screen viewport horizontally
      let leftPos = rect.left + scrollX - 10;
      if (leftPos + 300 > window.innerWidth) {
        leftPos = window.innerWidth - 310;
      }
      commBox.style.left = `${Math.max(10, leftPos)}px`;

      commBox.classList.remove('hidden');
    }

    function closeCommentBox() {
      document.getElementById('excel-comment-box').classList.add('hidden');
    }

    function saveChanges(showToast = true) {
      localStorage.setItem('webapp_data', JSON.stringify(state));
      if (showToast) {
        const toast = document.getElementById('toast');
        toast.classList.remove('hidden');
        setTimeout(() => toast.classList.add('hidden'), 3000);
      }
    }

    function exportToExcel() {
      const wb = XLSX.utils.book_new();

      const dashData = [
        ["Business Portal Operational Summary"],
        ["Total Active Bookings", state.bookings.length],
        [],
        ["Active Years (2026-2085)"],
        [2026, 2027, 2028, 2029, 2030, 2031]
      ];
      const wsDash = XLSX.utils.aoa_to_sheet(dashData);
      XLSX.utils.book_append_sheet(wb, wsDash, "Dashboard");

      const bookingHeaders = [["Invoice No", "Name", "Address", "ID No", "Contact No", "Room No", "Agent Info", "Capacity", "Check In", "Check Out", "Days", "Price/Day", "Total", "Advance", "Due"]];
      const bookingRows = state.bookings.map((b, idx) => [
        generateInvoiceNumber(idx), b.name, b.address, b.idNo, b.contactNo, b.roomNo, b.agentInfo, b.capacity,
        b.checkIn.replace('T', ' '), b.checkOut.replace('T', ' '), b.noOfDays, b.perDayPrice, b.totalAmount, b.advanced, b.totalDue
      ]);
      const wsBooking = XLSX.utils.aoa_to_sheet([...bookingHeaders, ...bookingRows]);
      XLSX.utils.book_append_sheet(wb, wsBooking, "Booking Details");

      const masterHeader = [["Agent Name", "Agent Ph No", "Room No", "Capacity"]];
      const masterRows = state.master.map(m => [m.agentName, m.phone, m.roomNo, m.capacity]);
      const wsMaster = XLSX.utils.aoa_to_sheet([...masterHeader, ...masterRows]);
      XLSX.utils.book_append_sheet(wb, wsMaster, "Master");

      XLSX.writeFile(wb, "Business_Portal_Compact_Export.xlsx");
    }
  </script>
</body>
</html>
