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
  </style>
</head>
<body class="bg-slate-100 text-slate-800 font-sans min-h-screen flex flex-col relative" onclick="hideTooltip()">

  <!-- Floating Tooltip for Calendar Hover/Click -->
  <div id="cal-tooltip" class="hidden absolute z-50 bg-slate-900 text-white text-xs rounded-xl p-3 shadow-2xl border border-slate-700 pointer-events-none transition-opacity duration-150 space-y-1 min-w-[160px]">
    <div class="font-bold text-indigo-300 border-b border-slate-700 pb-1 flex justify-between items-center">
      <span>Room Details</span>
      <i class="fa-solid fa-bed text-indigo-400"></i>
    </div>
    <p><span class="text-slate-400">Room No:</span> <strong id="tt-room" class="text-white"></strong></p>
    <p><span class="text-slate-400">Guest:</span> <strong id="tt-name" class="text-white"></strong></p>
    <p><span class="text-slate-400">Status:</span> <strong id="tt-status" class="text-emerald-400"></strong></p>
  </div>

  <!-- Header & Navigation -->
  <header class="bg-indigo-700 text-white shadow-lg sticky top-0 z-40 no-print">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-4 flex flex-col md:flex-row justify-between items-center gap-4">
      <div class="flex items-center space-x-3">
        <i class="fa-solid fa-hotel text-2xl text-indigo-200"></i>
        <h1 class="text-xl font-bold tracking-wide">Business Portal</h1>
      </div>
      
      <!-- Tab Navigation -->
      <nav class="flex space-x-1 sm:space-x-2 bg-indigo-800/60 p-1 rounded-xl text-sm font-medium">
        <button onclick="switchTab('dashboard')" id="btn-dashboard" class="tab-btn px-3 py-1.5 rounded-lg transition-all duration-200 active-tab">Dashboard</button>
        <button onclick="switchTab('booking')" id="btn-booking" class="tab-btn px-3 py-1.5 rounded-lg transition-all duration-200 text-indigo-100 hover:bg-indigo-600/50">Booking Details</button>
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-3 py-1.5 rounded-lg transition-all duration-200 text-indigo-100 hover:bg-indigo-600/50">Master Data</button>
        <button onclick="switchTab('calendar')" id="btn-calendar" class="tab-btn px-3 py-1.5 rounded-lg transition-all duration-200 text-indigo-100 hover:bg-indigo-600/50">Calendar</button>
      </nav>

      <!-- Action Buttons -->
      <div class="flex items-center space-x-3">
        <button onclick="saveChanges()" class="bg-emerald-500 hover:bg-emerald-600 active:bg-emerald-700 text-white px-4 py-2 rounded-lg text-sm font-semibold shadow flex items-center gap-2 transition">
          <i class="fa-solid fa-floppy-disk"></i> Save Changes
        </button>
        <button onclick="exportToExcel()" class="bg-amber-500 hover:bg-amber-600 active:bg-amber-700 text-white px-4 py-2 rounded-lg text-sm font-semibold shadow flex items-center gap-2 transition">
          <i class="fa-solid fa-file-excel"></i> Export to Excel
        </button>
      </div>
    </div>
  </header>

  <!-- Notification Toast -->
  <div id="toast" class="hidden fixed bottom-5 right-5 bg-slate-900 text-white px-5 py-3 rounded-xl shadow-2xl z-50 flex items-center gap-3 transition-all duration-300 no-print">
    <i class="fa-solid fa-circle-check text-emerald-400 text-lg"></i>
    <span id="toast-message">Changes saved successfully!</span>
  </div>

  <!-- Main Content Area -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 flex-1 w-full no-print">

    <!-- DASHBOARD TAB -->
    <section id="tab-dashboard" class="tab-content space-y-6">
      <div class="bg-gradient-to-r from-indigo-600 to-blue-600 rounded-2xl p-6 text-white shadow-xl">
        <h2 class="text-3xl font-extrabold mb-2">Hi Aniruddha, Welcome to your Business Portal 🤝</h2>
        <p class="text-indigo-100">Manage multiple bookings, agent details, and schedules effortlessly.</p>
      </div>

      <!-- Quick Summary Cards -->
      <div class="grid grid-cols-1 md:grid-cols-4 gap-6">
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-xs uppercase font-semibold text-slate-400">Total Bookings</p>
            <p id="dash-total-bookings" class="text-2xl font-bold text-slate-800 mt-1">1</p>
          </div>
          <div class="p-3 bg-blue-50 text-blue-600 rounded-xl"><i class="fa-solid fa-bookmark text-xl"></i></div>
        </div>
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-xs uppercase font-semibold text-slate-400">Total Booking Amount</p>
            <p id="dash-total-amount" class="text-2xl font-bold text-slate-800 mt-1">₹3,600</p>
          </div>
          <div class="p-3 bg-indigo-50 text-indigo-600 rounded-xl"><i class="fa-solid fa-receipt text-xl"></i></div>
        </div>
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-xs uppercase font-semibold text-slate-400">Advance Received</p>
            <p id="dash-advanced" class="text-2xl font-bold text-emerald-600 mt-1">₹2,000</p>
          </div>
          <div class="p-3 bg-emerald-50 text-emerald-600 rounded-xl"><i class="fa-solid fa-wallet text-xl"></i></div>
        </div>
        <div class="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 flex items-center justify-between">
          <div>
            <p class="text-xs uppercase font-semibold text-slate-400">Total Due</p>
            <p id="dash-due" class="text-2xl font-bold text-rose-600 mt-1">₹1,600</p>
          </div>
          <div class="p-3 bg-rose-50 text-rose-600 rounded-xl"><i class="fa-solid fa-hand-holding-dollar text-xl"></i></div>
        </div>
      </div>

      <!-- Years Grid (2026-2085) -->
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <h3 class="text-lg font-bold text-slate-800 mb-1 flex items-center gap-2">
          <i class="fa-solid fa-calendar-days text-indigo-600"></i> Active Years (2026 - 2085)
        </h3>
        <p class="text-xs text-slate-400 mb-4">Click any year below to switch directly to its calendar view.</p>
        <div id="years-grid" class="grid grid-cols-3 sm:grid-cols-6 md:grid-cols-10 gap-2">
          <!-- Populated by JS -->
        </div>
      </div>
    </section>

    <!-- BOOKING DETAILS TAB -->
    <section id="tab-booking" class="tab-content hidden space-y-6">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xl font-bold text-slate-800 flex items-center gap-2">
              <i class="fa-solid fa-address-card text-indigo-600"></i> Customer & Booking Directory
            </h2>
            <p class="text-xs text-slate-400">Manage multiple room reservations in one place.</p>
          </div>
          <button onclick="openBookingModal()" class="bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white px-4 py-2 rounded-xl text-sm font-semibold shadow flex items-center gap-2 transition">
            <i class="fa-solid fa-plus"></i> Add New Booking
          </button>
        </div>

        <!-- Bookings Table View -->
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-500 uppercase">
                <th class="p-3">Guest Name</th>
                <th class="p-3">Contact / ID</th>
                <th class="p-3">Room & Capacity</th>
                <th class="p-3">Agent Details</th>
                <th class="p-3">Stay Dates</th>
                <th class="p-3">Billing (Days x Rate)</th>
                <th class="p-3">Total / Advance</th>
                <th class="p-3">Due Status</th>
                <th class="p-3 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="bookings-tbody" class="divide-y divide-slate-100 text-sm">
              <!-- Populated by JS -->
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- MASTER DATA TAB -->
    <section id="tab-master" class="tab-content hidden space-y-6">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div class="flex justify-between items-center mb-6 pb-2 border-b border-slate-100">
          <h2 class="text-xl font-bold text-slate-800 flex items-center gap-2">
            <i class="fa-solid fa-users-gear text-indigo-600"></i> Master Agent & Room Directory
          </h2>
          <button onclick="addMasterRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-3 py-1.5 rounded-lg text-sm font-medium flex items-center gap-1 transition">
            <i class="fa-solid fa-plus"></i> Add Entry
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-xs font-semibold text-slate-500 uppercase">
                <th class="p-3">Agent Name</th>
                <th class="p-3">Agent Ph No</th>
                <th class="p-3">Room No</th>
                <th class="p-3">Capacity</th>
                <th class="p-3 text-center">Action</th>
              </tr>
            </thead>
            <tbody id="master-tbody" class="divide-y divide-slate-100 text-sm">
              <!-- Populated by JS -->
            </tbody>
          </table>
        </div>
      </div>
    </section>

    <!-- CALENDAR TAB -->
    <section id="tab-calendar" class="tab-content hidden space-y-6">
      <div class="bg-white rounded-2xl shadow-sm border border-slate-200 p-6">
        <div class="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4 mb-6">
          <h2 class="text-xl font-bold text-slate-800 flex items-center gap-2">
            <i class="fa-regular fa-calendar-check text-indigo-600"></i> Interactive Calendar
          </h2>
          
          <!-- Calendar Year Dropdown (2026-2085) -->
          <div class="flex items-center space-x-3">
            <label for="cal-year-select" class="text-xs font-bold text-slate-500 uppercase">Select Year:</label>
            <select id="cal-year-select" onchange="renderCalendar(parseInt(this.value))" class="bg-white border border-indigo-200 text-indigo-700 font-bold px-4 py-2 rounded-xl focus:outline-none focus:ring-2 focus:ring-indigo-500 cursor-pointer shadow-sm">
              <!-- Populated by JS from 2026 to 2085 -->
            </select>
          </div>
        </div>

        <div id="calendar-container" class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          <!-- Calendar months rendered dynamically -->
        </div>
      </div>
    </section>

  </main>

  <!-- ADD / EDIT BOOKING MODAL -->
  <div id="booking-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto no-print">
    <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-3xl w-full p-6 space-y-6 my-8">
      <div class="flex justify-between items-center pb-3 border-b border-slate-100">
        <h3 id="modal-title" class="text-lg font-bold text-slate-800 flex items-center gap-2">
          <i class="fa-solid fa-calendar-plus text-indigo-600"></i> Add New Booking
        </h3>
        <button onclick="closeBookingModal()" class="text-slate-400 hover:text-slate-600 text-lg"><i class="fa-solid fa-xmark"></i></button>
      </div>

      <form id="booking-form" onsubmit="handleSaveBooking(event)" class="space-y-6">
        <input type="hidden" id="modal-booking-id">

        <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">Customer Name</label>
            <input type="text" id="cust-name" required class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">Address</label>
            <input type="text" id="cust-address" class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">ID No</label>
            <input type="text" id="cust-id" class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">Contact No</label>
            <input type="text" id="cust-contact" class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm">
          </div>

          <!-- Room No Dropdown (Editable Selection) -->
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">Room No</label>
            <select id="cust-room" onchange="autoCaptureRoomDetails()" class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 font-semibold text-indigo-700 text-sm">
              <!-- Populated from Master -->
            </select>
          </div>

          <!-- Auto-Captured Non-Editable Fields (Grey) -->
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">
              Agent Info <span class="text-xs text-indigo-500 font-normal">(Auto)</span>
            </label>
            <input type="text" id="cust-agent" readonly class="w-full bg-slate-200/80 border border-slate-300 text-slate-600 font-medium rounded-xl px-3.5 py-2 cursor-not-allowed text-sm">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">
              Room Capacity <span class="text-xs text-indigo-500 font-normal">(Auto)</span>
            </label>
            <input type="text" id="cust-capacity" readonly class="w-full bg-slate-200/80 border border-slate-300 text-slate-600 font-medium rounded-xl px-3.5 py-2 cursor-not-allowed text-sm">
          </div>

          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">Check In Date & Time</label>
            <input type="datetime-local" id="cust-checkin" onchange="calculateModalBilling()" required class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm">
          </div>
          <div>
            <label class="block text-xs font-semibold text-slate-500 uppercase mb-1">Check Out Date & Time</label>
            <input type="datetime-local" id="cust-checkout" onchange="calculateModalBilling()" required class="w-full bg-white border border-slate-300 rounded-xl px-3.5 py-2 focus:outline-none focus:ring-2 focus:ring-indigo-500 text-sm">
          </div>
        </div>

        <!-- Billing Section -->
        <div class="bg-slate-50 p-4 rounded-xl border border-slate-200 space-y-3">
          <h4 class="text-xs font-bold uppercase tracking-wider text-slate-500">Billing Calculation</h4>
          <div class="grid grid-cols-2 sm:grid-cols-5 gap-3">
            <div>
              <label class="block text-xs font-semibold text-slate-500 mb-1">No. Days</label>
              <!-- Non-editable (Grey) -->
              <input type="number" id="cust-days" readonly class="w-full bg-slate-200/80 font-bold text-slate-700 border border-slate-300 rounded-lg px-3 py-1.5 cursor-not-allowed text-sm">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-500 mb-1">Price/Day (₹)</label>
              <!-- Editable (White) -->
              <input type="number" id="cust-price" value="1200" oninput="calculateModalBilling()" class="w-full bg-white font-bold text-slate-700 border border-slate-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-500 mb-1">Total (₹)</label>
              <!-- Non-editable (Grey) -->
              <input type="number" id="cust-total" readonly class="w-full bg-slate-200/80 text-indigo-700 font-bold border border-slate-300 rounded-lg px-3 py-1.5 cursor-not-allowed text-sm">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-500 mb-1">Advance (₹)</label>
              <!-- Editable (White) -->
              <input type="number" id="cust-advance" value="0" oninput="calculateModalBilling()" class="w-full bg-white border border-slate-300 rounded-lg px-3 py-1.5 text-sm focus:outline-none focus:ring-2 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block text-xs font-semibold text-slate-500 mb-1">Due (₹)</label>
              <!-- Non-editable (Grey) -->
              <input type="number" id="cust-due" readonly class="w-full bg-slate-200/80 text-rose-700 font-bold border border-slate-300 rounded-lg px-3 py-1.5 cursor-not-allowed text-sm">
            </div>
          </div>
        </div>

        <div class="flex justify-end space-x-3 pt-2">
          <button type="button" onclick="closeBookingModal()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-sm font-semibold transition">Cancel</button>
          <button type="submit" class="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 active:bg-indigo-800 text-white rounded-xl text-sm font-semibold shadow transition">Save Booking</button>
        </div>
      </form>
    </div>
  </div>

  <!-- PRINTABLE INVOICE MODAL / VIEW CONTAINER -->
  <div id="invoice-modal" class="hidden fixed inset-0 z-50 bg-slate-900/70 backdrop-blur-sm flex items-center justify-center p-4 overflow-y-auto">
    <div class="bg-white rounded-2xl shadow-2xl border border-slate-200 max-w-2xl w-full p-8 space-y-6 relative" id="printable-invoice">
      <!-- Invoice Header -->
      <div class="flex justify-between items-start border-b border-slate-200 pb-6">
        <div>
          <h2 class="text-2xl font-black text-indigo-700 uppercase tracking-wide">Hotel & Business Portal</h2>
          <p class="text-xs text-slate-500 mt-1">123 Business Street, Suite 400</p>
          <p class="text-xs text-slate-500">Phone: +91 98765 43210 | Email: info@businessportal.com</p>
        </div>
        <div class="text-right">
          <span class="inline-block bg-indigo-100 text-indigo-800 text-xs font-bold px-3 py-1 rounded-full uppercase mb-2">Booking Invoice</span>
          <p class="text-xs text-slate-500">Invoice ID: <strong id="inv-id" class="text-slate-800 font-mono">INV-A-0000001</strong></p>
          <p class="text-xs text-slate-500">Date: <strong id="inv-date" class="text-slate-800"></strong></p>
        </div>
      </div>

      <!-- Guest & Stay Details -->
      <div class="grid grid-cols-2 gap-6 bg-slate-50 p-4 rounded-xl border border-slate-100 text-xs">
        <div>
          <h4 class="font-bold text-slate-400 uppercase tracking-wider mb-2">Guest Information</h4>
          <p class="text-slate-800 font-semibold text-sm" id="inv-guest-name">-</p>
          <p class="text-slate-600" id="inv-guest-address">Address: -</p>
          <p class="text-slate-600" id="inv-guest-contact">Contact: -</p>
          <p class="text-slate-600" id="inv-guest-id">ID No: -</p>
        </div>
        <div>
          <h4 class="font-bold text-slate-400 uppercase tracking-wider mb-2">Reservation Info</h4>
          <p class="text-slate-800 font-semibold" id="inv-room">Room No: -</p>
          <p class="text-slate-600" id="inv-agent">Agent: -</p>
          <p class="text-slate-600" id="inv-checkin">Check-in: -</p>
          <p class="text-slate-600" id="inv-checkout">Check-out: -</p>
        </div>
      </div>

      <!-- Invoice Items Table -->
      <div class="overflow-x-auto">
        <table class="w-full text-left text-xs">
          <thead>
            <tr class="bg-indigo-50 text-indigo-900 border-b border-indigo-100">
              <th class="p-3">Description</th>
              <th class="p-3 text-center">Duration</th>
              <th class="p-3 text-right">Rate / Day</th>
              <th class="p-3 text-right">Amount</th>
            </tr>
          </thead>
          <tbody class="divide-y divide-slate-100">
            <tr>
              <td class="p-3 font-semibold text-slate-800" id="inv-item-desc">Room Accommodation</td>
              <td class="p-3 text-center" id="inv-item-days">0 Days</td>
              <td class="p-3 text-right" id="inv-item-rate">₹0</td>
              <td class="p-3 text-right font-semibold text-slate-800" id="inv-item-total">₹0</td>
            </tr>
          </tbody>
        </table>
      </div>

      <!-- Payment Summary -->
      <div class="flex justify-end pt-2 border-t border-slate-200">
        <div class="w-1/2 space-y-2 text-xs">
          <div class="flex justify-between text-slate-600">
            <span>Total Amount:</span>
            <strong id="inv-sum-total" class="text-slate-800">₹0</strong>
          </div>
          <div class="flex justify-between text-emerald-600">
            <span>Advance Received:</span>
            <strong id="inv-sum-advance">₹0</strong>
          </div>
          <div class="flex justify-between text-rose-600 text-sm font-bold border-t border-slate-200 pt-2">
            <span>Balance Due:</span>
            <span id="inv-sum-due">₹0</span>
          </div>
        </div>
      </div>

      <!-- Footer & Signature -->
      <div class="pt-8 border-t border-slate-200 flex justify-between items-end text-xs text-slate-400">
        <div>
          <p class="font-bold text-slate-600">Thank you for your business!</p>
          <p>For inquiries, contact management.</p>
        </div>
        <div class="text-center border-t border-slate-300 pt-2 w-36">
          <p class="font-semibold text-slate-600">Authorized Signatory</p>
        </div>
      </div>

      <!-- Modal Actions (Hidden when printing) -->
      <div class="flex justify-end space-x-3 pt-4 no-print border-t border-slate-100">
        <button type="button" onclick="closeInvoiceModal()" class="px-4 py-2 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded-xl text-sm font-semibold transition">Close</button>
        <button type="button" onclick="window.print()" class="px-5 py-2 bg-indigo-600 hover:bg-indigo-700 text-white rounded-xl text-sm font-semibold shadow flex items-center gap-2 transition">
          <i class="fa-solid fa-print"></i> Print Now
        </button>
      </div>
    </div>
  </div>

  <script>
    // App Data State with Multiple Bookings
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
          checkIn: "2026-01-22T11:00",
          checkOut: "2026-01-25T11:00",
          noOfDays: 3,
          perDayPrice: 1200,
          totalAmount: 3600,
          advanced: 2000,
          totalDue: 1600
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

    // Helper to generate padded linear invoice number starting from INV-A-0000001
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
        item.className = `text-center p-2 rounded-lg text-sm font-semibold border cursor-pointer transition ${
          y === state.selectedYear 
            ? 'bg-indigo-600 text-white border-indigo-600 shadow-md' 
            : 'bg-white text-slate-600 border-slate-200 hover:bg-indigo-50 hover:text-indigo-600 hover:border-indigo-300'
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

    // Print Invoice Handler with Linear Series Numbering
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

    // Modal & Multiple Bookings Management
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

      const newIn = new Date(checkIn);
      const newOut = new Date(checkOut);

      if (newIn >= newOut) {
        alert("Check-Out date & time must be after Check-In date & time.");
        return;
      }

      // Check for overlapping bookings for the same room
      const conflict = state.bookings.find(b => {
        if (id && b.id === id) return false; // Skip the current booking being edited
        if (parseInt(b.roomNo) !== roomNo) return false; // Different room, no conflict

        const existingIn = new Date(b.checkIn);
        const existingOut = new Date(b.checkOut);

        // Date overlap condition
        return (newIn < existingOut && newOut > existingIn);
      });

      if (conflict) {
        alert(`❌ Booking Conflict Alert!\n\nRoom ${roomNo} is already booked from ${conflict.checkIn.replace('T', ' ')} to ${conflict.checkOut.replace('T', ' ')} by ${conflict.name}.\n\nPlease either:\n1. Change the Room Number, or\n2. Change the Check-In / Check-Out date & time.`);
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
        tbody.innerHTML = `<tr><td colspan="9" class="text-center py-6 text-slate-400">No bookings available. Click "Add New Booking" to create one.</td></tr>`;
        return;
      }

      state.bookings.forEach((b) => {
        const checkInFmt = b.checkIn ? b.checkIn.replace('T', ' ') : '-';
        const checkOutFmt = b.checkOut ? b.checkOut.replace('T', ' ') : '-';
        
        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 transition";
        tr.innerHTML = `
          <td class="p-3 font-semibold text-slate-800">${b.name}</td>
          <td class="p-3 text-xs text-slate-500"><div>${b.contactNo}</div><div class="text-slate-400">${b.idNo}</div></td>
          <td class="p-3"><span class="bg-indigo-50 text-indigo-700 font-bold px-2.5 py-1 rounded-lg text-xs">Room ${b.roomNo}</span> <div class="text-xs text-slate-400 mt-0.5">${b.capacity}</div></td>
          <td class="p-3 text-xs text-slate-600">${b.agentInfo}</td>
          <td class="p-3 text-xs"><div><strong class="text-emerald-600">In:</strong> ${checkInFmt}</div><div><strong class="text-rose-600">Out:</strong> ${checkOutFmt}</div></td>
          <td class="p-3 text-xs font-medium">${b.noOfDays} days × ₹${b.perDayPrice}</td>
          <td class="p-3 text-xs"><div><strong>Total:</strong> ₹${b.totalAmount}</div><div class="text-emerald-600"><strong>Adv:</strong> ₹${b.advanced}</div></td>
          <td class="p-3"><span class="px-2.5 py-1 rounded-full text-xs font-bold ${b.totalDue > 0 ? 'bg-rose-100 text-rose-700' : 'bg-emerald-100 text-emerald-700'}">₹${b.totalDue} Due</span></td>
          <td class="p-3 text-center">
            <div class="flex items-center justify-center space-x-2">
              <button onclick="printInvoice('${b.id}')" title="Print Invoice" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-2.5 py-1 rounded-lg text-xs font-medium transition flex items-center gap-1">
                <i class="fa-solid fa-print text-indigo-600"></i> Print Invoice
              </button>
              <button onclick="openBookingModal('${b.id}')" title="Edit Booking" class="text-indigo-600 hover:text-indigo-800 p-1.5"><i class="fa-solid fa-pen-to-square"></i></button>
              <button onclick="deleteBooking('${b.id}')" title="Delete Booking" class="text-rose-500 hover:text-rose-700 p-1.5"><i class="fa-solid fa-trash-can"></i></button>
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
        tr.innerHTML = `
          <td class="p-3"><input type="text" value="${row.agentName}" onchange="updateMasterRow(${index}, 'agentName', this.value)" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:border-indigo-500 focus:outline-none"></td>
          <td class="p-3"><input type="text" value="${row.phone}" onchange="updateMasterRow(${index}, 'phone', this.value)" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:border-indigo-500 focus:outline-none"></td>
          <td class="p-3"><input type="number" value="${row.roomNo}" onchange="updateMasterRow(${index}, 'roomNo', this.value)" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:border-indigo-500 focus:outline-none"></td>
          <td class="p-3"><input type="number" value="${row.capacity}" onchange="updateMasterRow(${index}, 'capacity', this.value)" class="w-full bg-white border border-slate-300 rounded-lg px-2.5 py-1.5 focus:border-indigo-500 focus:outline-none"></td>
          <td class="p-3 text-center">
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

    // Calendar & Tooltip Logic for Multiple Bookings
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
        monthCard.className = "bg-slate-50 border border-slate-200 rounded-xl p-4";
        
        let header = `<h4 class="font-bold text-slate-700 text-center mb-3">${month} ${year}</h4>`;
        let daysHeader = `<div class="grid grid-cols-7 text-center text-xs font-semibold text-slate-400 mb-2"><div>S</div><div>M</div><div>T</div><div>W</div><div>T</div><div>F</div><div>S</div></div>`;
        
        let firstDay = new Date(year, monthIdx, 1).getDay();
        let totalDays = new Date(year, monthIdx + 1, 0).getDate();
        
        let daysGrid = `<div class="grid grid-cols-7 text-center text-xs gap-1">`;
        for (let i = 0; i < firstDay; i++) {
          daysGrid += `<div></div>`;
        }

        for (let d = 1; d <= totalDays; d++) {
          const currDate = new Date(year, monthIdx, d);
          
          let matchedBooking = null;
          let statusText = "Booked";

          for (let b of state.bookings) {
            if (!b.checkIn || !b.checkOut) continue;
            const cIn = new Date(b.checkIn);
            const cOut = new Date(b.checkOut);

            const checkInStart = new Date(cIn.getFullYear(), cIn.getMonth(), cIn.getDate());
            const checkOutEnd = new Date(cOut.getFullYear(), cOut.getMonth(), cOut.getDate());

            if (currDate >= checkInStart && currDate <= checkOutEnd) {
              matchedBooking = b;
              if (currDate.getTime() === checkInStart.getTime()) statusText = "Check-in Date";
              else if (currDate.getTime() === checkOutEnd.getTime()) statusText = "Check-out Date";
              else statusText = "Reserved Stay";
              break;
            }
          }

          if (matchedBooking) {
            daysGrid += `<div 
              onmousemove="showTooltip(event, '${matchedBooking.roomNo}', '${matchedBooking.name}', '${statusText}')" 
              onclick="showTooltip(event, '${matchedBooking.roomNo}', '${matchedBooking.name}', '${statusText}')" 
              onmouseleave="hideTooltip()"
              class="py-1.5 bg-indigo-600 text-white font-bold rounded-full cursor-pointer shadow hover:scale-105 transition-transform">${d}</div>`;
          } else {
            daysGrid += `<div class="py-1.5 hover:bg-slate-200 text-slate-700 cursor-pointer rounded-md">${d}</div>`;
          }
        }
        daysGrid += `</div>`;

        monthCard.innerHTML = header + daysHeader + daysGrid;
        container.appendChild(monthCard);
      });
    }

    function showTooltip(e, roomNo, guestName, status) {
      e.stopPropagation();
      const tt = document.getElementById('cal-tooltip');
      document.getElementById('tt-room').innerText = `Room ${roomNo}`;
      document.getElementById('tt-name').innerText = guestName || 'Guest';
      document.getElementById('tt-status').innerText = status;

      tt.style.left = `${e.pageX + 10}px`;
      tt.style.top = `${e.pageY + 10}px`;
      tt.classList.remove('hidden');
    }

    function hideTooltip() {
      document.getElementById('cal-tooltip').classList.add('hidden');
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
        ["Hi Aniruddha Welcome to your business Portal"],
        ["Total Active Bookings", state.bookings.length],
        [],
        ["Years(2026-2085)"],
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

      XLSX.writeFile(wb, "Business_Portal_MultipleBookings.xlsx");
    }
  </script>
</body>
</html>
