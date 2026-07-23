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
        <button onclick="switchTab('master')" id="btn-master" class="tab-btn px-2.5 py-1 rounded transition-all text-indigo-100 hover:bg-indigo-600/50">Master Data</button>
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
      <div class="bg-gradient-to-r from-indigo-700 via-indigo-600 to-blue-600 rounded-lg p-3 text-white shadow-sm flex justify-between items-center">
        <div>
          <h2 class="text-base font-bold tracking-tight">Hi Aniruddha, Welcome back! 🏠</h2>
          <p class="text-indigo-100 text-[10px] mt-0.5">Quickly view, schedule, and manage room allocations and orders.</p>
        </div>
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

      <!-- Compact Years Grid -->
      <div class="bg-white rounded-lg shadow-sm border border-slate-200 p-3">
        <div class="mb-2">
          <h3 class="text-xs font-bold text-slate-800 flex items-center gap-1">
            <i class="fa-solid fa-calendar-days text-indigo-600"></i> Active Years Directory (2026 – 2085)
          </h3>
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
          </div>
          
          <div class="flex items-center space-x-2 w-full md:w-auto">
            <div class="flex items-center bg-slate-50 border border-slate-300 rounded p-0.5 space-x-1">
              <select id="booking-search-select" class="bg-transparent text-[11px] px-1.5 py-0.5 focus:outline-none w-36 sm:w-44 text-slate-700 font-semibold cursor-pointer"></select>
              <button onclick="searchBookingById()" class="bg-slate-700 hover:bg-slate-800 text-white px-2 py-0.5 rounded text-[10px] font-semibold flex items-center gap-1 transition">
                <i class="fa-solid fa-magnifying-glass"></i> Search
              </button>
              <button onclick="clearSearchBooking()" class="text-slate-400 hover:text-slate-600 px-1 text-[10px]" title="Clear">
                <i class="fa-solid fa-xmark"></i>
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
                <th class="py-2 px-2">ID Number</th>
                <th class="py-2 px-2">Room</th>
                <th class="py-2 px-2">Agent</th>
                <th class="py-2 px-2 min-w-[150px]">Stay Window</th>
                <th class="py-2 px-2">Tariff & Extras</th>
                <th class="py-2 px-2">Payment</th>
                <th class="py-2 px-2">Balance</th>
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
      <div class="bg-white rounded-lg shadow-sm border border-slate-200 p-3">
        <div class="flex justify-between items-center mb-3 pb-2 border-b border-slate-100">
          <div>
            <h2 class="text-xs font-bold text-slate-800 flex items-center gap-1">
              <i class="fa-solid fa-users-gear text-indigo-600"></i> Master Agent & Room Directory
            </h2>
          </div>
          <button onclick="addMasterRow()" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded text-[11px] font-medium flex items-center gap-1 transition shadow-sm">
            <i class="fa-solid fa-plus text-[10px]"></i> Add Entry
          </button>
        </div>

        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-slate-50 border-b border-slate-200 text-[10px] font-bold text-slate-500 uppercase tracking-wider">
                <th class="py-2 px-2">Agent Name</th>
                <th class="py-2 px-2">Agent Contact</th>
                <th class="py-2 px-2">Room No</th>
                <th class="py-2 px-2">Capacity</th>
                <th class="py-2 px-2 text-center">Actions</th>
              </tr>
            </thead>
            <tbody id="master-tbody" class="divide-y divide-slate-100 text-[11px]"></tbody>
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
          
          <div class="flex items-center space-x-1.5 bg-slate-50 p-1 rounded border border-slate-200">
            <label for="cal-year-select" class="text-[9px] font-bold text-slate-500 uppercase">Year:</label>
            <select id="cal-year-select" onchange="renderCalendar(parseInt(this.value))" class="bg-white border border-slate-300 text-indigo-700 font-bold px-2 py-0.5 rounded focus:outline-none cursor-pointer text-[11px]"></select>
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
          <h3 class="text-xs font-bold">Upcoming Check-out Dues Alert (2 hrs)</h3>
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

        <!-- Guest Details Box -->
        <div class="bg-slate-50 p-2.5 rounded-md border border-slate-200 space-y-2">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-slate-500 flex items-center gap-1">
            <i class="fa-solid fa-user-tag text-indigo-500"></i> Guest Information
          </h4>
          <div class="grid grid-cols-2 sm:grid-cols-4 gap-2">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Guest Name <span class="text-rose-500">*</span></label>
              <!-- Mandatory Guest Name enforcing characters/letters only -->
              <input type="text" id="cust-name" required pattern="[A-Za-z\s]+" oninput="this.value = this.value.replace(/[^A-Za-z\s]/g, '')" title="Please enter Guest Name using characters only (letters and spaces)" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Address <span class="text-rose-500">*</span></label>
              <input type="text" id="cust-address" required title="Address is mandatory" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">ID Number <span class="text-rose-500">*</span></label>
              <!-- Mandatory ID Number allowing both characters and numbers -->
              <input type="text" id="cust-id" required pattern="[A-Za-z0-9\s]+" oninput="this.value = this.value.replace(/[^A-Za-z0-9\s]/g, '')" title="Please enter letters and numbers" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Contact No <span class="text-rose-500">*</span></label>
              <!-- Mandatory Contact No enforcing numbers only -->
              <input type="text" id="cust-contact" required pattern="[0-9]+" oninput="this.value = this.value.replace(/[^0-9]/g, '')" title="Please enter numbers only" class="w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-indigo-500">
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
              <label class="block font-semibold text-slate-600 mb-0.5">Agent Info <span class="text-[9px] text-indigo-500">(Auto)</span></label>
              <input type="text" id="cust-agent" readonly class="w-full bg-slate-200/70 border border-slate-300 text-slate-600 rounded px-2 py-1 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Capacity <span class="text-[9px] text-indigo-500">(Auto)</span></label>
              <input type="text" id="cust-capacity" readonly class="w-full bg-slate-200/70 border border-slate-300 text-slate-600 rounded px-2 py-1 cursor-not-allowed">
            </div>

            <!-- CHECK-IN / CHECK-OUT INPUTS -->
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

        <!-- EXTRA FOOD & BEVERAGES ORDERS BOX (DYNAMIC LIST) -->
        <div class="bg-amber-50/70 p-2.5 rounded-md border border-amber-200 space-y-2">
          <div class="flex justify-between items-center">
            <h4 class="text-[9px] font-bold uppercase tracking-wider text-amber-800 flex items-center gap-1">
              <i class="fa-solid fa-utensils text-amber-600"></i> Extra Food / Drink Orders List
            </h4>
            <button type="button" onclick="addFoodOrderItem()" class="bg-amber-600 hover:bg-amber-700 text-white px-2 py-0.5 rounded text-[10px] font-semibold flex items-center gap-1 transition">
              <i class="fa-solid fa-plus text-[9px]"></i> Add Food Order
            </button>
          </div>
          <p class="text-[9px] text-amber-700 italic font-semibold">(Allowed only from 30 mins after Check-In to 45 mins before Check-Out)</p>
          
          <div id="food-orders-container" class="space-y-2 max-h-40 overflow-y-auto pr-1"></div>
        </div>

        <!-- Billing Calculation Box -->
        <div class="bg-indigo-50/50 p-2.5 rounded-md border border-indigo-100 space-y-2">
          <h4 class="text-[9px] font-bold uppercase tracking-wider text-indigo-700 flex items-center gap-1">
            <i class="fa-solid fa-calculator text-indigo-600"></i> Billing Summary
          </h4>
          <div class="grid grid-cols-5 gap-1.5">
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Days</label>
              <input type="number" id="cust-days" readonly class="w-full bg-slate-200/80 font-bold text-slate-700 border border-slate-300 rounded px-1.5 py-1 cursor-not-allowed">
            </div>
            <div>
              <label class="block font-semibold text-slate-600 mb-0.5">Price/Day</label>
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
          </div>
        </div>

        <div class="flex justify-end space-x-2 pt-1">
          <button type="button" onclick="closeBookingModal()" class="px-3 py-1 bg-slate-100 text-slate-700 rounded font-semibold transition">Cancel</button>
          <button type="submit" class="px-4 py-1 bg-indigo-600 hover:bg-indigo-700 text-white rounded font-semibold shadow transition">Save Booking</button>
        </div>
      </form>
    </div>
  </div>

  <!-- PRINTABLE INVOICE MODAL -->
  <div id="invoice-modal" class="hidden fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-sm flex items-center justify-center p-3 overflow-y-auto">
    <div class="bg-white rounded-lg shadow-xl border border-slate-200 max-w-xl w-full p-6 space-y-4 relative" id="printable-invoice">
      <div class="flex justify-between items-start border-b border-slate-200 pb-3">
        <div>
          <h2 class="text-lg font-black text-indigo-700 uppercase tracking-wide">Aniruddha Homestay</h2>
          <p class="text-[10px] text-slate-500 mt-0.5">Sittong, Village in West Bengal</p>
          <p class="text-[10px] text-slate-500">Phone: +91 9804396541 | Email: info@businessportal.com</p>
        </div>
        <div class="text-right">
          <span class="inline-block bg-indigo-100 text-indigo-800 text-[9px] font-bold px-2 py-0.5 rounded-full uppercase mb-0.5">e-Invoice</span>
          <p class="text-[10px] text-slate-500">Invoice ID: <strong id="inv-id" class="text-slate-800 font-mono">INV-2026-0000001</strong></p>
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
          <p class="text-slate-600" id="inv-agent">Agent: -</p>
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
        <button type="button" onclick="window.print()" class="px-4 py-1 bg-indigo-600 text-white rounded font-semibold shadow flex items-center gap-1 transition">
          <i class="fa-solid fa-print"></i> Print Invoice
        </button>
      </div>
    </div>
  </div>

  <script>
    // Tab Close / Reload Prompt[cite: 1]
    window.addEventListener('beforeunload', function (e) {
      // Prevents tab closure without confirmation[cite: 1]
      e.preventDefault();
      // Required for browser compatibility to display the prompt dialog[cite: 1]
      e.returnValue = 'Please click "Save Changes" button to save the history.'; 
      return e.returnValue;
    });

    // Helper function to format ISO/DateTime strings to dd-mm-yyyy hh:mm (24-hour)
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

    // Helper function to format Date object to dd-mm-yyyy
    function formatDate(d) {
      if (!d) return '-';
      const dateObj = typeof d === 'string' ? new Date(d) : d;
      if (isNaN(dateObj.getTime())) return d;
      const day = String(dateObj.getDate()).padStart(2, '0');
      const month = String(dateObj.getMonth() + 1).padStart(2, '0');
      const year = dateObj.getFullYear();
      return `${day}-${month}-${year}`;
    }

    // State Directory
    let state = {
      yearlyCounters: { "2026": 2 },
      bookings: [
        {
          id: "bk_1",
          bookingCode: "BKG-2026-0000001",
          invoiceNo: "INV-2026-0000001",
          name: "Kapil",
          address: "Kolkata",
          idNo: "ABC25780001",
          contactNo: "1234567890",
          roomNo: 2,
          agentInfo: "A2 1234567890",
          capacity: "2 Person",
          checkIn: "2026-07-22T11:00",
          checkOut: "2026-07-24T11:00",
          noOfDays: 2,
          perDayPrice: 1200,
          foodOrders: [
            { foodDesc: "2x Breakfast, Tea", foodDateTime: "2026-07-22T16:30", foodCharge: 300 }
          ],
          totalAmount: 2700,
          advanced: 1000,
          totalDue: 1700
        },
        {
          id: "bk_2",
          bookingCode: "BKG-2026-0000002",
          invoiceNo: "INV-2026-0000002",
          name: "Aniruddha",
          address: "Mumbai",
          idNo: "XYZ99887766",
          contactNo: "9876543210",
          roomNo: 4,
          agentInfo: "A1 1234567890",
          capacity: "4 Person",
          checkIn: "2026-07-22T14:00",
          checkOut: "2026-07-25T10:00",
          noOfDays: 3,
          perDayPrice: 1500,
          foodOrders: [],
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

    function populateBookingSearchDropdown() {
      const select = document.getElementById('booking-search-select');
      if (!select) return;

      const selectedVal = select.value;
      select.innerHTML = '<option value="">-- All Booking IDs --</option>';
      
      state.bookings.forEach(b => {
        if (b.bookingCode) {
          const opt = document.createElement('option');
          opt.value = b.bookingCode;
          opt.text = b.bookingCode;
          select.appendChild(opt);
        }
      });

      select.value = selectedVal;
    }

    function searchBookingById() {
      const searchTerm = document.getElementById('booking-search-select').value.trim().toLowerCase();
      renderBookingsTable(searchTerm);
    }

    function clearSearchBooking() {
      document.getElementById('booking-search-select').value = '';
      renderBookingsTable();
    }

    function generateIDsForYear(checkInDateStr) {
      let targetYear = 2026;
      if (checkInDateStr) {
        targetYear = new Date(checkInDateStr).getFullYear() || 2026;
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
            parsed.bookings.forEach(b => {
              if (!b.foodOrders) {
                b.foodOrders = [];
                if (b.foodDesc || b.foodCharge) {
                  b.foodOrders.push({
                    foodDesc: b.foodDesc || '',
                    foodDateTime: b.foodDateTime || '',
                    foodCharge: b.foodCharge || 0
                  });
                }
              }
            });
            state = parsed;
            if (!state.yearlyCounters) {
              state.yearlyCounters = { "2026": state.bookings.length || 0 };
            }
          }
        } catch(e){}
      }
    }

    document.addEventListener("DOMContentLoaded", () => {
      loadSavedData();
      initDashboard();
      populateRoomDropdown();
      populateCalendarYearDropdown();
      populateBookingSearchDropdown();
      renderBookingsTable();
      renderMasterTable();
      renderCalendar(state.selectedYear);

      checkUpcomingCheckoutsWithDue();
      setInterval(checkUpcomingCheckoutsWithDue, 60000);

      // Auto-save data every 15 minutes (900,000 milliseconds)
      setInterval(saveChanges, 900000);

      // Notify the user immediately on opening/loading the web page
      showNotificationToast("Changes Auto save successfully!");
    });

    function checkUpcomingCheckoutsWithDue() {
      const now = new Date().getTime();
      const twoHoursMs = 2 * 60 * 60 * 1000;

      const upcomingDues = state.bookings.filter(b => {
        if (!b.checkOut || (b.totalDue || 0) <= 0) return false;
        const checkOutTime = new Date(b.checkOut).getTime();
        const diff = checkOutTime - now;
        return diff > -3600000 && diff <= twoHoursMs;
      });

      const badge = document.getElementById('alert-badge');
      if (upcomingDues.length > 0) {
        badge.innerText = upcomingDues.length;
        badge.classList.remove('hidden');
      } else {
        badge.classList.add('hidden');
      }

      renderAlertModalList(upcomingDues);
    }

    function renderAlertModalList(duesList) {
      const container = document.getElementById('alert-list-container');
      const textCount = document.getElementById('alert-list-count-text');
      container.innerHTML = '';

      textCount.innerText = `${duesList.length} active warnings found`;

      if (duesList.length === 0) {
        container.innerHTML = `
          <div class="text-center py-8 space-y-1">
            <div class="bg-emerald-50 text-emerald-600 w-10 h-10 rounded-full flex items-center justify-center mx-auto text-base">
              <i class="fa-solid fa-circle-check"></i>
            </div>
            <p class="font-bold text-slate-700">No Pending Checkout Dues</p>
            <p class="text-slate-400 text-[10px]">All upcoming check-outs within 2 hours are clear or fully paid.</p>
          </div>
        `;
        return;
      }

      duesList.forEach((b, i) => {
        const timeFormatted = formatDateTime(b.checkOut);

        const card = document.createElement('div');
        card.className = "bg-amber-50/60 border border-amber-200 rounded-md overflow-hidden shadow-sm";
        
        card.innerHTML = `
          <div class="p-2.5 flex justify-between items-center cursor-pointer hover:bg-amber-100/50 transition" onclick="toggleAlertDetails('alert-details-${i}')">
            <div class="flex items-center space-x-2">
              <span class="bg-amber-500 text-white p-1.5 rounded text-[10px] font-bold"><i class="fa-solid fa-clock"></i></span>
              <div>
                <h4 class="font-bold text-slate-800 text-[11px] flex items-center gap-1.5">
                  ${b.name} <span class="bg-indigo-100 text-indigo-800 text-[9px] px-1.5 py-0.2 rounded font-mono">${b.bookingCode || 'N/A'}</span>
                  <span class="bg-slate-100 text-slate-700 text-[9px] px-1.5 py-0.2 rounded font-medium">Room ${b.roomNo}</span>
                </h4>
                <p class="text-[10px] text-slate-500">Checkout: <strong class="text-amber-700">${timeFormatted}</strong></p>
              </div>
            </div>
            <div class="flex items-center space-x-1.5">
              <span class="text-[11px] font-black text-rose-600 bg-rose-50 border border-rose-200 px-2 py-0.5 rounded">
                ₹${b.totalDue.toLocaleString('en-IN')} Due
              </span>
              <i class="fa-solid fa-chevron-down text-slate-400 text-[10px]"></i>
            </div>
          </div>

          <div id="alert-details-${i}" class="hidden bg-white border-t border-amber-200/60 p-2.5 space-y-2 text-[10px]">
            <div class="grid grid-cols-2 gap-1 text-slate-600">
              <div>Contact: <strong>${b.contactNo || 'N/A'}</strong></div>
              <div>ID Number: <strong>${b.idNo || 'N/A'}</strong></div>
              <div>Total Charges: <strong>₹${b.totalAmount}</strong></div>
              <div>Advance Paid: <strong class="text-emerald-600">₹${b.advanced}</strong></div>
            </div>
            <div class="flex justify-end pt-1 border-t border-slate-100">
              <button onclick="closeAlertModal(); openBookingModal('${b.id}')" class="bg-indigo-600 hover:bg-indigo-700 text-white px-2.5 py-1 rounded font-bold text-[10px] flex items-center gap-1 transition">
                <i class="fa-solid fa-wallet"></i> Collect Payment / Edit
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
      yearSelect.innerHTML = '';
      for (let y = 2026; y <= 2085; y++) {
        const opt = document.createElement('option');
        opt.value = y;
        opt.text = `Year ${y}`;
        if (y === state.selectedYear) opt.selected = true;
        yearSelect.appendChild(opt);
      }
    }

    function populateRoomDropdown(selectedRoomNo = 1) {
      const roomSelect = document.getElementById('cust-room');
      roomSelect.innerHTML = '';
      state.master.forEach(m => {
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
        item.className = `text-center py-1 px-1 rounded text-[10px] font-bold border cursor-pointer transition ${
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

      const today = formatDate(new Date());
      document.getElementById('inv-id').innerText = b.invoiceNo;
      document.getElementById('inv-booking-id').innerText = b.bookingCode;
      document.getElementById('inv-date').innerText = today;

      document.getElementById('inv-guest-name').innerText = b.name || 'N/A';
      document.getElementById('inv-guest-address').innerText = `Address: ${b.address || 'N/A'}`;
      document.getElementById('inv-guest-contact').innerText = `Contact: ${b.contactNo || 'N/A'}`;
      document.getElementById('inv-guest-id').innerText = `ID No: ${b.idNo || 'N/A'}`;

      document.getElementById('inv-room').innerText = `Room No: ${b.roomNo} (${b.capacity || 'Standard'})`;
      document.getElementById('inv-agent').innerText = `Agent: ${b.agentInfo || 'Direct'}`;
      document.getElementById('inv-checkin').innerText = `Check-in: ${formatDateTime(b.checkIn)}`;
      document.getElementById('inv-checkout').innerText = `Check-out: ${formatDateTime(b.checkOut)}`;

      const tbody = document.getElementById('inv-items-tbody');
      tbody.innerHTML = '';

      const roomTotal = (b.noOfDays || 0) * (b.perDayPrice || 0);
      const roomTr = document.createElement('tr');
      roomTr.innerHTML = `
        <td class="p-2 font-semibold text-slate-800">Room ${b.roomNo} Accommodation</td>
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
              <td class="p-2 font-semibold text-slate-800">Extra Food & Drinks <span class="text-[9px] text-slate-500 font-normal block">${fo.foodDesc || 'Food Order'}${foodDateTimeFmt}</span></td>
              <td class="p-2 text-center">1 Order</td>
              <td class="p-2 text-right">₹${(fo.foodCharge || 0).toLocaleString('en-IN')}</td>
              <td class="p-2 text-right font-semibold text-slate-800">₹${(fo.foodCharge || 0).toLocaleString('en-IN')}</td>
            `;
            tbody.appendChild(foodTr);
          }
        });
      }

      document.getElementById('inv-sum-total').innerText = `₹${(b.totalAmount || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-advance').innerText = `₹${(b.advanced || 0).toLocaleString('en-IN')}`;
      document.getElementById('inv-sum-due').innerText = `₹${(b.totalDue || 0).toLocaleString('en-IN')}`;

      document.getElementById('invoice-modal').classList.remove('hidden');
    }

    function closeInvoiceModal() {
      document.getElementById('invoice-modal').classList.add('hidden');
    }

    // MULTIPLE FOOD ORDERS LIST DYNAMIC HANDLERS
    function addFoodOrderItem(desc = '', dateStr = '', timeStr = '', charge = 0) {
      const container = document.getElementById('food-orders-container');
      const itemRow = document.createElement('div');
      itemRow.className = "food-order-row grid grid-cols-1 sm:grid-cols-12 gap-1.5 items-end bg-white p-2 rounded border border-amber-200/80 shadow-xs";
      
      itemRow.innerHTML = `
        <div class="sm:col-span-5">
          <label class="block font-semibold text-slate-600 mb-0.5">Item Details</label>
          <input type="text" value="${desc}" placeholder="e.g. 2x Tea, Snacks" class="cust-food-desc w-full bg-white border border-slate-300 rounded px-2 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500">
        </div>
        
        <div class="sm:col-span-4">
          <label class="block font-semibold text-slate-600 mb-0.5"><i class="fa-regular fa-clock text-amber-600 mr-1"></i> Date & Time</label>
          <div class="flex gap-1">
            <input type="date" value="${dateStr}" onchange="validateAndCalculateFoodTimeRow(this)" class="cust-food-date w-3/5 bg-white border border-slate-300 rounded px-1 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500 font-medium text-[10px]">
            <input type="time" value="${timeStr}" onchange="validateAndCalculateFoodTimeRow(this)" class="cust-food-time w-2/5 bg-white border border-slate-300 rounded px-1 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500 font-medium text-[10px]">
          </div>
        </div>

        <div class="sm:col-span-2">
          <label class="block font-semibold text-slate-600 mb-0.5">Charge (₹)</label>
          <input type="number" value="${charge}" min="0" oninput="calculateModalBilling()" class="cust-food-charge w-full bg-white font-bold text-amber-700 border border-slate-300 rounded px-1.5 py-1 focus:outline-none focus:ring-1 focus:ring-amber-500">
        </div>

        <div class="sm:col-span-1 flex justify-end">
          <button type="button" onclick="removeFoodOrderItem(this)" class="text-rose-500 hover:text-rose-700 p-1.5" title="Remove Order">
            <i class="fa-solid fa-trash-can"></i>
          </button>
        </div>
      `;

      container.appendChild(itemRow);
      calculateModalBilling();
    }

    function removeFoodOrderItem(btn) {
      const row = btn.closest('.food-order-row');
      if (row) {
        row.remove();
        calculateModalBilling();
      }
    }

    function openBookingModal(bookingId = null) {
      const form = document.getElementById('booking-form');
      form.reset();
      document.getElementById('food-orders-container').innerHTML = '';

      if (bookingId) {
        const b = state.bookings.find(item => item.id === bookingId);
        if (b) {
          document.getElementById('modal-title').innerText = 'Edit Booking Details';
          document.getElementById('modal-booking-id').value = b.id;
          document.getElementById('cust-name').value = b.name;
          document.getElementById('cust-address').value = b.address;
          document.getElementById('cust-id').value = b.idNo;
          document.getElementById('cust-contact').value = b.contactNo;
          
          populateRoomDropdown(b.roomNo);

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
              addFoodOrderItem(fo.foodDesc || '', fDate, fTime, fo.foodCharge || 0);
            });
          }

          document.getElementById('cust-price').value = b.perDayPrice;
          document.getElementById('cust-advance').value = b.advanced;
          calculateModalBilling();
        }
      } else {
        document.getElementById('modal-title').innerText = 'Add New Booking';
        document.getElementById('modal-booking-id').value = '';
        
        populateRoomDropdown(1);

        document.getElementById('cust-checkin-time').value = "12:00";
        document.getElementById('cust-checkout-time').value = "11:00";
        document.getElementById('cust-price').value = 1200;
        document.getElementById('cust-advance').value = 0;
        calculateModalBilling();
      }

      document.getElementById('booking-modal').classList.remove('hidden');
    }

    function closeBookingModal() {
      document.getElementById('booking-modal').classList.add('hidden');
    }

    function handleStayDatesChange() {
      calculateModalBilling();
      validateAllFoodTimes();
    }

    function validateAndCalculateFoodTimeRow(inputElem) {
      const row = inputElem.closest('.food-order-row');
      if (!row) return;

      const dateElem = row.querySelector('.cust-food-date');
      const timeElem = row.querySelector('.cust-food-time');

      const foodDate = dateElem.value;
      const foodTime = timeElem.value;

      if (!foodDate && !foodTime) return true;

      const inDate = document.getElementById('cust-checkin-date').value;
      const inTime = document.getElementById('cust-checkin-time').value || '00:00';
      const outDate = document.getElementById('cust-checkout-date').value;
      const outTime = document.getElementById('cust-checkout-time').value || '00:00';

      if (!inDate || !outDate || !foodDate) return true;

      const checkInMs = new Date(`${inDate}T${inTime}`).getTime();
      const checkOutMs = new Date(`${outDate}T${outTime}`).getTime();
      
      const thirtyMinsMs = 30 * 60 * 1000;
      const fortyFiveMinsMs = 45 * 60 * 1000;

      const minFoodTimeMs = checkInMs + thirtyMinsMs;
      const maxFoodTimeMs = checkOutMs - fortyFiveMinsMs;

      if (maxFoodTimeMs <= minFoodTimeMs) {
        alert("⚠️ Stay duration is too short to allow extra food/drink orders with the required 30-minute post Check-In and 45-minute pre Check-Out buffers.");
        dateElem.value = '';
        timeElem.value = '';
        return false;
      }

      const foodMs = new Date(`${foodDate}T${foodTime || '00:00'}`).getTime();

      if (foodMs < minFoodTimeMs || foodMs > maxFoodTimeMs) {
        const formattedMin = formatDateTime(minFoodTimeMs);
        const formattedMax = formatDateTime(maxFoodTimeMs);

        alert(`⚠️ Invalid Food/Drink Order Time!\n\nExtra orders are ONLY permitted within:\n• Earliest: ${formattedMin} (30 mins after Check-In)\n• Latest: ${formattedMax} (45 mins before Check-Out)`);
        
        dateElem.value = '';
        timeElem.value = '';
        return false;
      }

      return true;
    }

    function validateAllFoodTimes() {
      const rows = document.querySelectorAll('.food-order-row');
      for (let r of rows) {
        const dateInput = r.querySelector('.cust-food-date');
        if (dateInput && dateInput.value) {
          if (!validateAndCalculateFoodTimeRow(dateInput)) return false;
        }
      }
      return true;
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
      
      let foodTotalCharge = 0;
      document.querySelectorAll('.cust-food-charge').forEach(input => {
        foodTotalCharge += parseFloat(input.value) || 0;
      });

      const advance = parseFloat(document.getElementById('cust-advance').value) || 0;
      
      const roomTotal = days * price;
      const total = roomTotal + foodTotalCharge;
      const due = total - advance;

      document.getElementById('cust-days').value = days;
      document.getElementById('cust-total').value = total;
      document.getElementById('cust-due').value = due;
    }

    function handleSaveBooking(e) {
      e.preventDefault();

      // Additional strict check for mandatory guest information fields
      const guestName = document.getElementById('cust-name').value.trim();
      const guestAddress = document.getElementById('cust-address').value.trim();
      const guestId = document.getElementById('cust-id').value.trim();
      const guestContact = document.getElementById('cust-contact').value.trim();

      if (!guestName || !guestAddress || !guestId || !guestContact) {
        alert("⚠️ All Guest Information fields (Guest Name, Address, ID Number, and Contact No) are mandatory to proceed!");
        return;
      }

      if (!validateAllFoodTimes()) {
        return;
      }

      const id = document.getElementById('modal-booking-id').value;
      const roomNo = parseInt(document.getElementById('cust-room').value);
      
      const inDate = document.getElementById('cust-checkin-date').value;
      const inTime = document.getElementById('cust-checkin-time').value || '00:00';
      const outDate = document.getElementById('cust-checkout-date').value;
      const outTime = document.getElementById('cust-checkout-time').value || '00:00';

      const checkIn = `${inDate}T${inTime}`;
      const checkOut = `${outDate}T${outTime}`;

      const foodOrdersList = [];
      document.querySelectorAll('.food-order-row').forEach(row => {
        const desc = row.querySelector('.cust-food-desc').value || '';
        const fDate = row.querySelector('.cust-food-date').value || '';
        const fTime = row.querySelector('.cust-food-time').value || '';
        const charge = parseFloat(row.querySelector('.cust-food-charge').value) || 0;

        const foodDateTime = (fDate && fTime) ? `${fDate}T${fTime}` : (fDate ? `${fDate}T00:00` : '');

        if (desc || charge > 0 || foodDateTime) {
          foodOrdersList.push({
            foodDesc: desc,
            foodDateTime: foodDateTime,
            foodCharge: charge
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
      
      const newBooking = {
        id: id || `bk_${Date.now()}`,
        bookingCode: existingCode,
        invoiceNo: existingInv,
        name: guestName,
        address: guestAddress,
        idNo: guestId,
        contactNo: guestContact,
        roomNo: roomNo,
        agentInfo: document.getElementById('cust-agent').value,
        capacity: document.getElementById('cust-capacity').value,
        checkIn: checkIn,
        checkOut: checkOut,
        noOfDays: parseInt(document.getElementById('cust-days').value) || 0,
        perDayPrice: parseFloat(document.getElementById('cust-price').value) || 0,
        foodOrders: foodOrdersList,
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
      populateBookingSearchDropdown();
      renderBookingsTable();
      updateDashboardCards();
      renderCalendar(state.selectedYear);
      checkUpcomingCheckoutsWithDue();
      saveChanges(false);
    }

    function deleteBooking(id) {
      if (confirm('Are you sure you want to delete this booking entry?')) {
        state.bookings = state.bookings.filter(b => b.id !== id);
        populateBookingSearchDropdown();
        renderBookingsTable();
        updateDashboardCards();
        renderCalendar(state.selectedYear);
        checkUpcomingCheckoutsWithDue();
        saveChanges(false);
      }
    }

    function renderBookingsTable(filterQuery = "") {
      const tbody = document.getElementById('bookings-tbody');
      tbody.innerHTML = '';

      let listToRender = state.bookings;

      if (filterQuery) {
        listToRender = state.bookings.filter(b => 
          (b.bookingCode && b.bookingCode.toLowerCase().includes(filterQuery))
        );
      }

      if (listToRender.length === 0) {
        tbody.innerHTML = `<tr><td colspan="11" class="text-center py-6 text-slate-400">No bookings found matching your search.</td></tr>`;
        return;
      }

      listToRender.forEach((b) => {
        const checkInFmt = formatDateTime(b.checkIn);
        const checkOutFmt = formatDateTime(b.checkOut);

        let foodSummaryHtml = '';
        if (b.foodOrders && b.foodOrders.length > 0) {
          const totalFoodCharge = b.foodOrders.reduce((acc, fo) => acc + (fo.foodCharge || 0), 0);
          if (totalFoodCharge > 0) {
            foodSummaryHtml = `<div class="text-[9px] text-amber-700 font-semibold"><i class="fa-solid fa-utensils text-[8px] mr-0.5"></i>Food (${b.foodOrders.length}): +₹${totalFoodCharge}</div>`;
          }
        }

        const tr = document.createElement('tr');
        tr.className = "hover:bg-slate-50 transition border-b border-slate-100";
        tr.innerHTML = `
          <td class="py-2 px-2">
            <span class="bg-indigo-50 border border-indigo-200 text-indigo-700 font-mono font-bold px-1.5 py-0.2 rounded text-[9px] block w-max">${b.bookingCode}</span>
          </td>
          <td class="py-2 px-2 font-bold text-slate-800">${b.name}</td>
          <td class="py-2 px-2 text-[10px] font-medium text-slate-700">${b.contactNo || '-'}</td>
          <td class="py-2 px-2 text-[10px] text-slate-500 font-mono">${b.idNo || '-'}</td>
          <td class="py-2 px-2">
            <span class="bg-indigo-50 text-indigo-700 font-bold px-1.5 py-0.2 rounded text-[10px] inline-block">Room ${b.roomNo}</span>
            <div class="text-[9px] text-slate-400">${b.capacity}</div>
          </td>
          <td class="py-2 px-2 text-[10px] text-slate-600 font-medium">${b.agentInfo}</td>
          <td class="py-2 px-2 text-[10px]">
            <div class="text-emerald-700 font-medium"><i class="fa-solid fa-plane-arrival mr-1"></i> ${checkInFmt}</div>
            <div class="text-rose-600 font-medium"><i class="fa-solid fa-plane-departure mr-1"></i> ${checkOutFmt}</div>
          </td>
          <td class="py-2 px-2 text-[10px]">
            <div class="font-semibold text-slate-700">${b.noOfDays} d × ₹${b.perDayPrice}</div>
            ${foodSummaryHtml}
          </td>
          <td class="py-2 px-2 text-[10px]">
            <div class="font-bold text-slate-800">Tot: ₹${b.totalAmount}</div>
            <div class="text-emerald-600 font-medium">Adv: ₹${b.advanced}</div>
          </td>
          <td class="py-2 px-2">
            <span class="px-1.5 py-0.2 rounded text-[10px] font-bold inline-block ${b.totalDue > 0 ? 'bg-rose-100 text-rose-700' : 'bg-emerald-100 text-emerald-700'}">
              ₹${b.totalDue} Due
            </span>
          </td>
          <td class="py-2 px-2 text-center">
            <div class="flex items-center justify-center space-x-1.5">
              <button onclick="printInvoice('${b.id}')" title="Print Invoice" class="bg-slate-100 hover:bg-slate-200 text-slate-700 px-1.5 py-0.5 rounded text-[10px] font-semibold transition border border-slate-200">
                <i class="fa-solid fa-print text-indigo-600"></i> Inv
              </button>
              <button onclick="openBookingModal('${b.id}')" title="Edit Booking" class="text-indigo-600 hover:text-indigo-800 p-0.5"><i class="fa-solid fa-pen-to-square"></i></button>
              <button onclick="deleteBooking('${b.id}')" title="Delete Booking" class="text-rose-500 hover:text-rose-700 p-0.5"><i class="fa-solid fa-trash-can"></i></button>
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
          <td class="py-1.5 px-2"><input type="text" value="${row.agentName}" onchange="updateMasterRow(${index}, 'agentName', this.value)" class="w-full bg-white border border-slate-300 rounded px-2 py-0.5 focus:border-indigo-500 text-[11px]"></td>
          <td class="py-1.5 px-2"><input type="text" pattern="[0-9]+" oninput="this.value = this.value.replace(/[^0-9]/g, '')" value="${row.phone}" onchange="updateMasterRow(${index}, 'phone', this.value)" class="w-full bg-white border border-slate-300 rounded px-2 py-0.5 focus:border-indigo-500 text-[11px]"></td>
          <td class="py-1.5 px-2"><input type="number" value="${row.roomNo}" onchange="updateMasterRow(${index}, 'roomNo', this.value)" class="w-full bg-white border border-slate-300 rounded px-2 py-0.5 focus:border-indigo-500 font-semibold text-indigo-700 text-[11px]"></td>
          <td class="py-1.5 px-2"><input type="number" value="${row.capacity}" onchange="updateMasterRow(${index}, 'capacity', this.value)" class="w-full bg-white border border-slate-300 rounded px-2 py-0.5 focus:border-indigo-500 text-[11px]"></td>
          <td class="py-1.5 px-2 text-center">
            <button onclick="removeMasterRow(${index})" class="text-rose-500 hover:text-rose-700 p-0.5"><i class="fa-solid fa-trash-can"></i></button>
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

    function renderCalendar(year) {
      state.selectedYear = year;
      const calSelect = document.getElementById('cal-year-select');
      if (calSelect) calSelect.value = year;

      const container = document.getElementById('calendar-container');
      container.innerHTML = '';

      const months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];

      months.forEach((month, monthIdx) => {
        const monthCard = document.createElement('div');
        monthCard.className = "bg-white border border-slate-200 rounded-lg p-2 shadow-sm";
        
        let header = `<h4 class="font-bold text-slate-800 text-center mb-1 text-[11px] border-b border-slate-100 pb-1">${month} ${year}</h4>`;
        let daysHeader = `<div class="grid grid-cols-7 text-center text-[9px] font-bold text-slate-400 mb-0.5"><div>S</div><div>M</div><div>T</div><div>W</div><div>T</div><div>F</div><div>S</div></div>`;
        
        let firstDay = new Date(year, monthIdx, 1).getDay();
        let totalDays = new Date(year, monthIdx + 1, 0).getDate();
        
        let daysGrid = `<div class="grid grid-cols-7 text-center text-[10px] gap-0.5">`;
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
            const formattedDateStr = formatDate(currDate);
            
            const badgeBg = dayBookings.length > 1 
              ? 'bg-amber-600 text-white font-black hover:bg-amber-700' 
              : 'bg-indigo-600 text-white font-bold hover:bg-indigo-700';

            daysGrid += `<div 
              onclick="toggleCommentBox(event, this, '${formattedDateStr}', '${bookingsJson}')" 
              class="relative py-0.5 ${badgeBg} rounded cursor-pointer shadow-sm flex flex-col items-center justify-center">
                <span>${d}</span>
                <span class="absolute top-0 right-0 w-0 h-0 border-t-[4px] border-r-[4px] border-t-amber-300 border-r-transparent rounded-tr-sm"></span>
              </div>`;
          } else {
            daysGrid += `<div class="py-0.5 hover:bg-slate-100 text-slate-700 cursor-pointer rounded font-medium">${d}</div>`;
          }
        }
        daysGrid += `</div>`;

        monthCard.innerHTML = header + daysHeader + daysGrid;
        container.appendChild(monthCard);
      });
    }

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
        
        let foodSummaryHtml = '';
        if (b.foodOrders && b.foodOrders.length > 0) {
          const totalCharge = b.foodOrders.reduce((acc, fo) => acc + (fo.foodCharge || 0), 0);
          if (totalCharge > 0) {
            foodSummaryHtml = `<p><span class="text-slate-400">Food Extra:</span> <span class="text-amber-300 font-medium">₹${totalCharge} (${b.foodOrders.length} item/s)</span></p>`;
          }
        }

        const card = document.createElement('div');
        card.className = "bg-slate-800/95 border border-slate-700/80 p-2 rounded space-y-0.5 text-[10px] shadow-inner";
        
        card.innerHTML = `
          <div class="flex justify-between items-center text-amber-300 font-bold border-b border-slate-700/60 pb-0.5">
            <span class="truncate max-w-[100px]"><i class="fa-solid fa-user text-[9px] mr-1 text-amber-400"></i> ${b.name || 'Guest'}</span>
            <span class="bg-indigo-900/90 text-indigo-200 text-[8px] px-1 py-0.2 rounded font-mono border border-indigo-700">${b.bookingCode}</span>
          </div>
          <p><span class="text-slate-400">Contact:</span> <strong class="text-slate-200">${b.contactNo || 'N/A'}</strong></p>
          <p><span class="text-slate-400">Invoice:</span> <span class="text-slate-300 font-mono text-[9px]">${b.invoiceNo}</span></p>
          <p><span class="text-slate-400">Status:</span> <strong class="text-emerald-400">${item.statusText}</strong></p>
          ${foodSummaryHtml}
          <p><span class="text-slate-400">Check-In:</span> <span class="text-slate-300 font-mono text-[9px]">${formatDateTime(b.checkIn)}</span></p>
          <p><span class="text-slate-400">Check-Out:</span> <span class="text-slate-300 font-mono text-[9px]">${formatDateTime(b.checkOut)}</span></p>
        `;

        container.appendChild(card);
      });

      const rect = targetElem.getBoundingClientRect();
      const scrollY = window.scrollY || document.documentElement.scrollTop;
      const scrollX = window.scrollX || document.documentElement.scrollLeft;

      commBox.style.top = `${rect.bottom + scrollY + 4}px`;
      
      let leftPos = rect.left + scrollX - 10;
      if (leftPos + 260 > window.innerWidth) {
        leftPos = window.innerWidth - 270;
      }
      commBox.style.left = `${Math.max(10, leftPos)}px`;

      commBox.classList.remove('hidden');
    }

    function closeCommentBox() {
      document.getElementById('excel-comment-box').classList.add('hidden');
    }

    function showNotificationToast(msg = "Changes Auto save successfully!") {
      const toast = document.getElementById('toast');
      const toastMsg = document.getElementById('toast-message');
      if (toastMsg) toastMsg.innerText = msg;
      
      if (toast) {
        toast.classList.remove('hidden');
        setTimeout(() => toast.classList.add('hidden'), 4000);
      }
    }

    function saveChanges(showToast = true) {
      localStorage.setItem('webapp_data', JSON.stringify(state));
      if (showToast) {
        showNotificationToast("Changes Auto save successfully!");
      }
    }

    function exportToExcel() {
      const wb = XLSX.utils.book_new();

      const dashData = [
        ["Business Portal Summary"],
        ["Total Active Bookings", state.bookings.length],
        [],
        ["Active Years (2026-2085)"],
        [2026, 2027, 2028, 2029, 2030, 2031]
      ];
      const wsDash = XLSX.utils.aoa_to_sheet(dashData);
      XLSX.utils.book_append_sheet(wb, wsDash, "Dashboard");

      const bookingHeaders = [["Booking ID", "Invoice ID", "Guest Name", "Address", "ID No", "Contact No", "Room No", "Agent Info", "Capacity", "Check In", "Check Out", "Days", "Price/Day", "Total Food Charges", "Total", "Advance", "Due"]];
      const bookingRows = state.bookings.map((b) => {
        const totalFoodCharge = (b.foodOrders || []).reduce((acc, fo) => acc + (fo.foodCharge || 0), 0);
        return [
          b.bookingCode,
          b.invoiceNo,
          b.name, b.address, b.idNo, b.contactNo, b.roomNo, b.agentInfo, b.capacity,
          formatDateTime(b.checkIn), formatDateTime(b.checkOut), b.noOfDays, b.perDayPrice, totalFoodCharge, b.totalAmount, b.advanced, b.totalDue
        ];
      });
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
