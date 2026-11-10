/* Luck Swan Resort sales system user manual — English (US) content (translated from content.th.js, the master version)
 * To edit or add a chapter: add an object to chapters (or to quick for the short guides by department)
 * in content.th.js, then add the same id here and in content.my.js — ids must match in all 3 languages
 * Images: img/<image name>-<th|en|my>.webp (screenshots of the real app in all 3 languages) */
window.MANUAL = window.MANUAL || {};
window.MANUAL.en = {
  "ui": {
    "lang": "Language",
    "full": "Full manual",
    "quick": "Quick guides by department",
    "contents": "Contents",
    "search": "Search the manual…",
    "who": "Used by",
    "steps": "Steps",
    "tips": "Good to know",
    "print": "Print / Save PDF",
    "top": "Back to top",
    "version": "Version",
    "updated": "Last updated",
    "changelog": "Manual change history",
    "menu": "Menus",
    "can": "What you can do",
    "daily": "Daily work",
    "open": "Open the full chapter",
    "noresult": "No matching topics found",
    "zoom": "Tap an image to enlarge"
  },
  "meta": {
    "title": "Sales System User Manual",
    "subtitle": "Luck Swan Resort — quotations, meeting rooms, guest rooms, catering, function sheets, staff, work roster, leave and assignments",
    "url": "https://luckswan-sales.it-luckswan.workers.dev",
    "version": "1.5",
    "updated": "Nov 10, 2026",
    "intro": "The sales system is a web app that works on computers, iPads and phones. Everyone sees the same data in real time. The menus you see depend on the role and permissions the administrator gives you. This manual covers every chapter, from signing in to each department’s work, with real screenshots for every step."
  },
  "groups": [
    {
      "id": "start",
      "title": "Getting started",
      "color": "#A8262D"
    },
    {
      "id": "sell",
      "title": "Sales",
      "color": "#0E9F8E"
    },
    {
      "id": "event",
      "title": "Catering and function sheets",
      "color": "#D4A21A"
    },
    {
      "id": "people",
      "title": "Staff and internal work",
      "color": "#2B3F86"
    },
    {
      "id": "admin",
      "title": "Administrator",
      "color": "#17211F"
    }
  ],
  "chapters": [
    {
      "id": "login",
      "group": "start",
      "icon": "key",
      "title": "Signing in and installing on your device",
      "who": "Everyone",
      "intro": "Open luckswan-sales.it-luckswan.workers.dev in a browser (Safari, Chrome or Edge) and sign in with the account the administrator created for you.",
      "steps": [
        "Type your email or employee ID in the first field.",
        "Type your password (at least 6 characters). Tap “Show” to see what you typed.",
        "Turn on “Remember me on this device” if it is your own device (leave it off on shared devices).",
        "Tap “Sign in”. The first time, you may be asked to set a new password.",
        "Choose a property: after signing in, property cards appear — Luckswan Resort · Luck Swan Boutique Hotel · Manee Inn Hotel. You only see the properties the administrator allows. Tap a card to start.",
        "To change property later, tap the property name ▾ next to the logo in the top bar.",
        "Forgot your password? Tap “Forgot password?” or call the administrator at 053-750-855.",
        "Install as an app on a phone/iPad: tap your profile picture at the top right › “How to install on phone / tablet” and follow the steps (Safari: Share button › Add to Home Screen)."
      ],
      "shots": [
        {
          "img": "login",
          "cap": "Sign-in page — the language and color theme buttons are at the top right"
        },
        {
          "img": "site",
          "cap": "Choose a property after signing in"
        }
      ],
      "tips": [
        "Data is separated by property: quotations, calendar, function sheets, staff shifts, Dashboard and notifications show only the selected property · new quotations get the current property automatically.",
        "The language and color theme buttons work right from the sign-in page.",
        "The app updates itself automatically. If an update was just released, close the app and open it again."
      ]
    },
    {
      "id": "home",
      "group": "start",
      "icon": "home",
      "title": "Home page and menu bar",
      "who": "Everyone",
      "intro": "After you sign in, you see the Home page. The menu bar is at the top, and shortcuts and a work summary are in the middle.",
      "steps": [
        "Top menu bar: Home · Rooms & Venues ▾ (Meeting Rooms, Guest Rooms) · Catering · Quotations · Function Sheets · Calendar · Staff · Users · Leave Requests · Assignment & Approve · Dashboard (you only see the menus you have permission for).",
        "On a phone, the main menus are in the bottom bar. Tap ☰ at the top left to see all menus.",
        "Colored cards at the top: shortcuts to meeting rooms and events · group guest rooms · catering and food · seminar packages with accommodation.",
        "The “+ New quotation” button at the top right starts a quotation right away.",
        "In the “Latest quotations” table, tap a row to open that quotation.",
        "The “… target” box shows closed sales against the monthly target, and the “Assigned tasks” box shows your latest tasks.",
        "The search box at the top right finds customers or quotation numbers from any page."
      ],
      "shots": [
        {
          "img": "home",
          "cap": "Home page: shortcuts, latest quotations, sales target and assigned tasks"
        },
        {
          "img": "usermenu",
          "cap": "Profile menu: change password, how to install, back up data, log out"
        }
      ],
      "tips": [
        "The profile menu has: Change password · How to install on phone / tablet · Back up data (download) · Import backup · Log out."
      ]
    },
    {
      "id": "langtheme",
      "group": "start",
      "icon": "globe",
      "title": "Changing the language and color theme",
      "who": "Everyone",
      "intro": "Make the screen easy for you to read. Settings are saved on the device you use and do not affect anyone else.",
      "steps": [
        "Change the language: tap the letter button (labeled TH/EN/MY) next to the bell, then choose ไทย (default) · English (US) · မြန်မာ.",
        "Change the color theme: tap the seven-color circle button next to the language button.",
        "Pick one of the 8 “Ready-made themes”, such as Standard, Ocean, Emerald or Royal purple — or tap a color circle under “Any color”.",
        "Mix your own: tap the Main color / Background circle to open the color picker, or type a color code such as #1E88E5.",
        "Eyedropper: tap the eyedropper icon, then tap the color you want on the screen.",
        "Tap “Back to the standard colors” to return to the original colors."
      ],
      "shots": [
        {
          "img": "lang",
          "cap": "Language menu with 3 languages"
        },
        {
          "img": "theme",
          "cap": "Color theme menu: ready-made themes, any color, mix your own, color code and eyedropper"
        }
      ],
      "tips": [
        "Text people type themselves (names, customer names, reasons) is shown as typed and is not translated.",
        "Documents printed or sent to customers are always in Thai with the standard colors."
      ]
    },
    {
      "id": "notifs",
      "group": "start",
      "icon": "bell",
      "title": "Notifications",
      "who": "Everyone",
      "intro": "The bell at the top right shows things that need your attention. The red number is how many you have not read yet.",
      "steps": [
        "Tap the bell to see your notifications, such as function sheets awaiting approval, tasks assigned to you, or leave requests waiting to be signed.",
        "Tap a notification to go straight to the related page.",
        "Tap “View all tasks” to open the Assignment & Approve page."
      ],
      "shots": [
        {
          "img": "notifs",
          "cap": "Notifications window"
        }
      ],
      "tips": []
    },
    {
      "id": "meet",
      "group": "sell",
      "icon": "meet",
      "title": "Meeting Rooms",
      "who": "Sales · Sales Manager · Administrator",
      "intro": "Rooms & Venues › Meeting Rooms shows free rooms, capacity and prices, and lets you add a room to a quotation.",
      "steps": [
        "Set the options at the top: event date · number of attendees · Full day / Half day.",
        "Choose a room layout: Theatre · Classroom · U-shape · Banquet · Cocktail — the system works out the capacity for that layout.",
        "Check the room cards: size, capacity, full-day price and status, such as “Free all 2 days”, or a red bar showing dates that are already booked.",
        "Tap a room card to see its details on the right: room photos, seating plan, room conditions and equipment.",
        "Tap “Choose this room” to add it to the quotation you are working on (the quotation number “For QT-…” shows under the heading; tap “Change” to switch quotations).",
        "Room photos: swipe left/right or tap the ‹ › arrows to see the next photo, both on the room card and in the panel on the right (a counter such as 2/3 shows which photo you are viewing).",
        "Full-screen view: tap the ⤢ icon at the bottom right of a photo — move with ‹ ›, a swipe or the arrow keys; close with ✕, a tap on the dark area or Esc.",
        "The pencil button ✏️ at the bottom right of a room card (people who can edit prices): edit the sample photos (upload several, delete, reorder — the first photo is the cover), the full-day/half-day prices, the capacity for each layout (min–max, such as 150–300 guests), the area, the banquet minimum guarantee, the description and what is in the room, then tap “Save” — there is also a “Reset to default” button."
      ],
      "shots": [
        {
          "img": "meet",
          "cap": "Meeting room list with capacity and availability"
        },
        {
          "img": "meet-sel",
          "cap": "Room selected: seating plan, conditions and equipment"
        },
        {
          "img": "meet-edit",
          "cap": "Editing a meeting room: photos, prices and capacity"
        },
        {
          "img": "meet-zoom",
          "cap": "Room photo in full screen"
        }
      ],
      "tips": [
        "Roles without “Create and send quotations” (such as junior sales) can see rooms, photos, capacity and prices, but not the quotation bar, provisional booking or the add / update in quotation buttons.",
        "If a booking overlaps another event, the system warns you and an administrator must approve it first.",
        "Rooms that are over capacity for the chosen layout show a warning.",
        "A new room price applies to items added to quotations from now on; existing quotations do not change."
      ]
    },
    {
      "id": "rooms",
      "group": "sell",
      "icon": "bed",
      "title": "Guest Rooms and room rates",
      "who": "Sales · Administrator (rates)",
      "intro": "Rooms & Venues › Guest Rooms is where you set up a group’s room block. You can choose between 2 places: Resort (prices follow the season automatically) and Boutique.",
      "steps": [
        "Choose the “Place” at the top: Resort or Boutique.",
        "Resort: choose the check-in date and number of nights at the top, then choose With ABF (breakfast included) or Room only.",
        "On each room type card, tap − / + to set the number of rooms and choose the bed type (Twin / Double / Triple for 3 guests).",
        "Check the room block summary on the right: number of rooms, guests and free rooms for the organiser (1 free for every 20 rooms).",
        "Turn Late check-out 14:00 on or off as agreed with the customer.",
        "Check the release date (cut-off) in the yellow box. Rooms without a guest list by then go back to the system automatically.",
        "Tap “Update quote” to save the guest rooms to the quotation.",
        "Boutique: Standard 400 · Superior 500 · Deluxe 600 THB per room per night (including congee, coffee and crackers for 2). Tap − / + to set the number of rooms.",
        "Boutique: extra meals per person — Breakfast (extra) 50 / 100 / 150 THB and Dinner party 350 / 399 / 450 THB. Choose a price, then enter the number of guests (the first time you tap +, the system fills in the quotation’s number of attendees). It is multiplied by the number of nights automatically.",
        "Boutique: tap “Update quote” — the rooms on the quotation become boutique rooms (names start with “Boutique ·”) and the meals go into the Food & beverage section · When you open Guest Rooms again, it comes back to Boutique with the same quantities.",
        "Users allowed to edit rates: tap “Room rates” to edit the rate periods (seasons), number of rooms and the With ABF / Room only price for each type — under “Boutique”, edit the 3 room prices and the 3 meal price levels, then save."
      ],
      "shots": [
        {
          "img": "rooms",
          "cap": "Resort room block: choose room types, quantities and bed types"
        },
        {
          "img": "rooms-bq",
          "cap": "Boutique guest rooms and extra meals"
        },
        {
          "img": "roomrates",
          "cap": "Setting room rates by season"
        },
        {
          "img": "roomrates-bq",
          "cap": "Boutique rates: guest rooms and meals"
        }
      ],
      "tips": [
        "Roles without “Create and send quotations” (such as junior sales) can see rooms, photos, capacity and prices, but not the quotation bar, provisional booking or the add / update in quotation buttons.",
        "Room rates are per room per night for 2 guests (Triple: 3 guests), net prices with service charge and VAT included (worked out the same way as the quotation).",
        "A quotation holds guest rooms from one place at a time — tapping Update quote from Boutique replaces the resort rooms already on the quotation (and vice versa).",
        "Boutique does not use the resort’s 1 free room for every 20 rooms or Late check-out."
      ]
    },
    {
      "id": "food",
      "group": "sell",
      "icon": "food",
      "title": "Catering and menus",
      "who": "Sales · Administrator (menu editing)",
      "intro": "The Catering menu covers seminar packages, coffee breaks, lunch, gala dinners, cocktails and weddings.",
      "steps": [
        "Choose a category on the left, or filter by price range and dietary suitability (Halal · Vegetarian · Gluten-free).",
        "Check the menu cards: photo, items in the package, price per person and minimum.",
        "Tap “+ Add to quote” to add it to the quotation you are working on — the bottom bar shows the total and a “Back to quote” button.",
        "If you can edit prices: tap “Edit” on a card to change the name, price, minimum, description, suitability labels, main photo and package items.",
        "Tap “+ Add menu / package” to add a new menu.",
        "Manage catering categories (people who may edit prices): the ＋ and ✏️ buttons next to “Category” — ＋ adds a category · ✏️ renames or deletes 🗑 (each row shows how many menus it has), then tap “Save”."
      ],
      "shots": [
        {
          "img": "food",
          "cap": "Catering menu list and filters"
        },
        {
          "img": "food-cats",
          "cap": "Manage catering categories: add, rename, delete"
        },
        {
          "img": "food-edit",
          "cap": "Editing a package: price, minimum, photo and food items"
        }
      ],
      "tips": [
        "Roles without “Create and send quotations” (such as junior sales) can see menus and prices, but not the quotation bar, the “+ Add to quote” buttons or the quotation bar at the bottom.",
        "New prices apply to quotations made from now on. Quotations already sent do not change.",
        "A category that still has menus cannot be deleted (move its menus first) · “All” cannot be deleted · “Seminar packages” can be renamed but not deleted · names must be unique."
      ]
    },
    {
      "id": "quotes",
      "group": "sell",
      "icon": "doc",
      "title": "Quotations: list and editing",
      "who": "Sales · Sales Manager · Accounting (view)",
      "intro": "The Quotations menu lists every quotation. Search and filter by status: Draft · Awaiting approval · Tentative · Confirmed · Cancelled — the steps: Draft › Awaiting approval › (an administrator approves) Tentative › (the customer confirms) Confirmed › make the function sheet.",
      "steps": [
        "Search by customer name or quotation number, then tap a row to open it — the “Owner” column after the number shows the user who made the quotation (on phones it shows under the event name).",
        "On the quotation page, items are grouped into Meeting Rooms · Food & beverage · Guest Rooms. Edit quantities or prices right in the fields, and tap “+ Add item…” to add more.",
        "Switch between the “Table / Timeline” views at the top right to see the daily schedule.",
        "Summary box on the right: prices on the quotation are Net prices (service charge and VAT included) — subtotal before discount · discount · price before service charge and VAT · 10% service charge · total before VAT · 7% VAT · net total, plus the average per person.",
        "Discounts over 10% need the sales manager’s approval.",
        "At the bottom, set Price valid until · Deposit (for example 30% within 7 days of confirmation) · tentative room hold until a date.",
        "Tap “Edit customer details and terms” to change the customer name, contact person, phone and email.",
        "When the draft is ready: in the “Status” box tap “Request approval” — the status becomes “Awaiting approval” › tap “Print quotation” and print it for an administrator to sign in the “Approved by” space · to keep editing, tap “Take back to edit” · the change history is at the end of the box.",
        "Administrators: quotations awaiting approval show on the Home page (Follow up today) › open one and tap “Approve quotation” — it becomes “Tentative” automatically (rooms held 10 days if no date was set). The approval box turns yellow, “Waiting for the customer to confirm”, and the owner is notified.",
        "Not approved: the administrator taps “Send back” and can add a reason — the quotation goes back to Draft, and the salesperson sees the reason in the status box and on the Home page; fix it, then tap “Request approval” again.",
        "After approval the quotation can go to the customer (email · LINE · link). When the customer confirms: sales staff/the sales manager tap “Customer confirmed booking” — the status becomes “Confirmed”, the approval box turns green, and the “Function sheet” button appears.",
        "If the total changes after approval, the box says “Total changed after approval”; administrators tap “Approve again” or “Cancel approval”."
      ],
      "shots": [
        {
          "img": "quotes",
          "cap": "Quotation list with status filters (provisional = yellow)"
        },
        {
          "img": "quote-edit",
          "cap": "Quotation edit page and summary box"
        },
        {
          "img": "quote-ask",
          "cap": "Status box: Awaiting approval — the administrator approves the quotation or sends it back"
        },
        {
          "img": "quote-approve",
          "cap": "Status box: quotation approved, waiting for the customer to confirm"
        }
      ],
      "tips": [
        "The system saves automatically every time you make a change.",
        "The “Customer confirmed booking” and “Tentative” buttons appear only after an administrator has approved · a function sheet can be made once the customer has confirmed the booking.",
        "Older quotations that were “Sent” now show as “Awaiting approval” (or “Tentative” if an administrator had already approved them).",
        "If a meeting room overlaps another event and the overlap is not yet approved, the quotation can be approved but its status does not change to provisional.",
        "How the total is broken down: price before SC/VAT = net total ÷ 1.177 → service charge is 10% of that price → total before VAT → 7% VAT. For example, 1,200 = 1,019.54 + 101.95 + 78.51."
      ]
    },
    {
      "id": "newquote",
      "group": "sell",
      "icon": "plus",
      "title": "Creating a new quotation (5 steps)",
      "who": "Sales · Sales Manager",
      "intro": "Tap “+ New quotation” and the system walks you through it step by step. The net total (including service charge + VAT) is always shown on the right.",
      "steps": [
        "Step 1 Customer: fill in the event name* and customer/company*, Tax ID (13 digits, optional), contact person, phone, email, start–end dates and number of attendees, then tap “Next”.",
        "Step 2 Meeting rooms: choose a room layout and tick the rooms (the system shows capacity and price). You can choose half day/full day and an AV package.",
        "Step 3 Food: turn each meal on or off (morning coffee break · lunch · afternoon coffee break · gala dinner · cocktail), then choose menus. You can add dietary requirements.",
        "Step 4 Guest rooms: set the number of nights and the number of rooms of each type.",
        "Step 5 Summary: check the items and total, add a discount, then tap “Create quotation”.",
        "Use the back button (it shows the previous step’s name) to change an earlier step at any time, or “Cancel” to stop."
      ],
      "shots": [
        {
          "img": "wiz1",
          "cap": "Step 1 Customer and event details"
        },
        {
          "img": "wiz2",
          "cap": "Step 2 Choosing meeting rooms"
        },
        {
          "img": "wiz3",
          "cap": "Step 3 Food & beverage"
        },
        {
          "img": "wiz4",
          "cap": "Step 4 Guest rooms"
        },
        {
          "img": "wiz5",
          "cap": "Step 5 Summary before creating the quotation"
        }
      ],
      "tips": [
        "Fields marked * are required. You cannot go to the next step until they are filled in.",
        "The colored bar on the total at the right shows the split between meeting rooms / food / guest rooms."
      ]
    },
    {
      "id": "sendquote",
      "group": "sell",
      "icon": "send",
      "title": "Printing and sending a quotation",
      "who": "Sales · Sales Manager",
      "intro": "On the quotation page, tap “Send / Print”. You see a preview of the real document on the left and the sending panel on the right.",
      "steps": [
        "Check the document on the left (company header, customer, items, prices, terms).",
        "Choose the quotation format in the panel on the right: “Customer copy” (hides the separate service charge and VAT lines and shows the net total) or “Full quotation” (shows every line) — it is printed and sent in the format you choose.",
        "Tap “Edit details” (sales): the text on the quotation turns into dashed boxes you can type in, such as the customer name, address, Tax ID, item names, details, quantities, terms and the signatory’s name — the ↺ button = Use the system value. Then tap “Save changes” (it applies to that quotation only).",
        "Meeting room section on the quotation: shows the room photo and price. In edit mode you can type your own room details (tap “+ Insert the room’s details from the system” to bring in the existing text), and tap the photo to upload a new one, ✕ to remove the photo, or ↺ to use the room’s photo.",
        "Room photos on the document: each room type chosen (e.g. Superior) has its photos under the room details — in edit mode tap the 📷 “Add room photo” box (several at once), ✕ to remove, ↺ to use the system photos (this quotation only).",
        "Tap “Print / PDF” to print or save as PDF.",
        "Before an administrator approves: email, LINE and link cannot be sent (a red box says so) — Print / PDF still works, so an administrator can sign in the “Approved by” space at the bottom of the quotation.",
        "Send by email: enter the recipient’s email, edit the message, and choose “Let the customer confirm and sign online” and “Show unit prices”.",
        "Set “Price valid until”, then tap “Send quotation”.",
        "The sending history is at the bottom of the panel."
      ],
      "shots": [
        {
          "img": "quote-doc",
          "cap": "Customer copy and the email panel"
        },
        {
          "img": "quote-doc-full",
          "cap": "Full quotation (shows separate service charge and VAT lines)"
        },
        {
          "img": "quote-doc-edit",
          "cap": "Editing details on the quotation"
        },
        {
          "img": "quote-doc-rooms",
          "cap": "Room photos under the chosen rooms"
        }
      ],
      "tips": [
        "Documents sent to customers are always in Thai, even if you use the app in another language.",
        "The logo and “Branch” on the document are those of the quotation’s property (Resort / Boutique / Manee Inn).",
        "Boutique rooms have no photos in the system yet; add them yourself.",
        "Amounts and totals cannot be edited on the quotation — change the items on the quotation page.",
        "The terms at the end include the line “Credit card payments carry an extra 3% fee on the quotation amount” · If you edited the terms field on a quotation yourself, tap ↺ to get the latest text."
      ]
    },
    {
      "id": "beo",
      "group": "event",
      "icon": "list",
      "title": "Function Sheets (BEO)",
      "who": "Sales · Managers · Heads of Department · Accounting (view)",
      "intro": "A function sheet is the work order for every department. It is created from a confirmed quotation, and each department acknowledges it in the system.",
      "steps": [
        "The list on the left is split into Upcoming · All · Draft. Labels show the status, such as “All departments” or “2 to acknowledge”.",
        "Tap “+ Create from quote” and choose a quotation; rooms, food and guests come in automatically — only quotations an administrator has approved and the customer has confirmed are listed (or tap “Function sheet” in the quotation’s status box).",
        "The header shows the event date, confirmed guests (+5% spare), rooms, contact person and salesperson.",
        "The “Attendees and type of event” box: tap a branch heading to open it, then tap to choose as many as you need, e.g. Secondary school students · Teachers / lecturers / Team building · Training class. The branch heading shows what you chose · ✏️ Edit / add / delete choices inside the branch (Administrator, Sales Manager, Sales Staff); other roles can only view.",
        "Daily schedule: time, details, place and responsible department (switch between “Day 1 / Day 2”). You can import it from an Excel file.",
        "Each department card (kitchen, banquet, AV, front office, housekeeping, accounting) shows its tasks and “Acknowledged / Waiting” — tap “Edit” to add the department’s tasks. The card and that department’s box in “Function details” are one list: edit either and both change.",
        "Every change made after sending is counted as “Revision …” and goes to an administrator for approval before it is sent to departments.",
        "Buttons at the top: Function details · Sales summary · Floor plan (administrators, sales, F&B and banquet) · Room setup (administrators) · Print / PDF."
      ],
      "shots": [
        {
          "img": "beo",
          "cap": "Function sheet: daily schedule and department cards"
        },
        {
          "img": "beo-tags",
          "cap": "Attendees and type of event"
        },
        {
          "img": "beo-list",
          "cap": "Function sheet list with acknowledgement status"
        }
      ],
      "tips": [
        "A quotation that is not approved, or not yet confirmed by the customer, cannot become a function sheet; the system says “The admin must approve the quotation first” or “A function sheet can be issued once the customer confirms the booking” · existing function sheets can still be edited.",
        "Heads of department should tap “Acknowledge” every time there is a new sheet or a change.",
        "Choosing attendees and type of event does not need a new edit approval · the choices are shared by all sheets · what you pick shows in the “Type of event” box of Function details."
      ]
    },
    {
      "id": "plan",
      "group": "event",
      "icon": "grid",
      "title": "Floor plans and room setup",
      "who": "Administrator · Sales Manager · Sales Staff · F&B department · Banquet department (heads and staff)",
      "intro": "Draw the room layout from a function sheet and print it for the setup team. Works on computers, iPads and phones.",
      "steps": [
        "On the function sheet, tap “Floor plan” and choose the day (Day 1 / Day 2) — if the event uses more than one meeting room, choose the room first; each room has its own plan.",
        "The “Tables / rows” panel on the left takes the number of guests from the quotation and works out the rows and tables per row (Classroom · Theatre = chairs per row · Banquet), with seats per table and a centre aisle.",
        "Change “Rows”, “Tables per row” or “Guests” yourself; the total updates at once, e.g. “3 rows × 6 tables = 18 tables · 180 seats”, with a red warning when seats are short or the rows are more than the room takes — tap “Lay out like this” to lay all tables out again, or “Recalculate” to work it out again.",
        "Or choose a template on the left (From the quotation · Theatre · Meeting · Round tables, etc.); the system places the tables and chairs automatically.",
        "Tap “+ Round table / + Long table / + Chair row / + Stage / + Text box …” to add items, then drag them into place — a text box: add it, then type in the “Text” box on the left straight away, e.g. “VIP entrance”.",
        "Tap an item to select, move, copy or delete it. Tap “Undo” if you make a mistake.",
        "Saving: while you arrange the plan nothing is sent to the server (others still see the saved plan), and the top shows “Unsaved changes” — when you are done tap “Save floor plan” (or Ctrl+S) · “Discard changes” goes back to the saved plan.",
        "Rotate: drag the red circle above the selected item’s frame (5° steps), or tap ⟳ 15° · 45° · 90° · 180° in the “Selected” panel (clockwise; the current angle is shown).",
        "A floor-plan picture (.jpg .png): tap + next to “Templates” › upload the picture — it shows as the background at full strength; fade it to place tables over it.",
        "Use the picture as the plan: in the “Background picture” panel tap “Use this picture as the plan” — the function sheet and print use it instead; type the number of seats yourself · “Back to the drawn plan” brings all the tables back · “Remove picture” takes it away.",
        "“Download PNG” saves the plan as a picture.",
        "“Room setup” (administrators): set the room size (meters) and fixed items that cannot be moved, such as the stage, doors, pillars and text boxes — this applies to every event in that room."
      ],
      "shots": [
        {
          "img": "plan",
          "cap": "Floor plan: the tables / rows panel, templates and the room plan"
        },
        {
          "img": "plan-tools",
          "cap": "Selected item: rotate handle, 15° · 45° · 90° · 180° buttons and a text box"
        },
        {
          "img": "plan-image",
          "cap": "An uploaded floor-plan picture used as the plan, with the seats typed in"
        },
        {
          "img": "plansetup",
          "cap": "Room setup: room size and fixed items"
        }
      ],
      "tips": [
        "“Lay out like this” lays out all the tables again (fixed items stay); if the room cannot take them all, it says how many did not fit · Undo works.",
        "The plan remembers the rows / tables it was last laid out with.",
        "A text box takes no table space when plans are laid out automatically.",
        "Plans can be printed as PDF and attached to the function sheet.",
        "Sales are allowed by their role; F&B and banquet by the “Department” field in the user account: Kitchen / F&B / Food & beverage / Banquet — other people can see the plan thumbnail on the function sheet but cannot open the designer.",
        "“Room setup” is for administrators only, and also needs “Save floor plan” (the “Standard item sizes” table saves straight away).",
        "Leaving the page, or changing day or room, with unsaved changes asks first · if the app is closed or refreshed by mistake, the unsaved plan comes back when you open the same page, “(restored from last time)” — it is kept on that device only.",
        "If someone else saves the same plan while you are working on it, you are asked first: “save yours over it?”"
      ]
    },
    {
      "id": "fnforms",
      "group": "event",
      "icon": "print",
      "title": "Function details and sales summary",
      "who": "Administrator · Sales Manager · Sales Staff (edit) · Accounting · Heads of Department (view)",
      "intro": "Printable forms that match the hotel’s original paper forms. Data comes from the quotation and function sheet.",
      "steps": [
        "Tap “Function details” — switch between “Print preview / Edit details” and enter the Function No. (numbered automatically, or type your own) — the Edit details tab, the “＋ Add details” button and numbering are only for administrators, sales managers and sales staff.",
        "The “Accounting / Finance” section (total, cost details, billing address) is only visible to administrators, sales managers, sales staff and accounting — other roles see only the Front Office section, at full width.",
        "“For sales / For operations” buttons (administrators, sales managers, sales staff, accounting): For sales = the full sheet with the finance part · For operations = the copy the operations team sees, without the finance part — printing and saving use the copy shown.",
        "The “Room layout”, kitchen, housekeeping, AV and front office boxes come from the department cards on the function sheet — edit them in the Edit details tab and the card changes; edit the card and the sheet changes · the “Type of event” box comes from the attendees and type of event chosen.",
        "The “Front office” section: each guest-room line of the quotation gets its own block, as on the old paper form — Check-in · Check-out (n nights) · Room (type) · Number of rooms, with no lines inside a block and a line under each room · every box can be edited, e.g. “1 room (Khun Thip + Phueng)”.",
        "Tap “Sales summary” — a summary of function revenue, food, snacks, guest rooms and tax invoice details.",
        "Check it, then tap “Print / PDF”.",
        "Tap “Save” (after Print / PDF) and choose PNG picture (sharp) · JPG picture (small, for LINE) · PDF (A4) — the file goes to your device, named after the sheet and the copy, e.g. function-FS-2610-003-operation.pdf."
      ],
      "shots": [
        {
          "img": "fn",
          "cap": "Function details: For sales / For operations and Save buttons"
        },
        {
          "img": "fn-front",
          "cap": "Front office: one block per guest-room line"
        },
        {
          "img": "fn-save",
          "cap": "Save as a PNG · JPG · PDF file"
        },
        {
          "img": "ss",
          "cap": "Function summary for sales"
        }
      ],
      "tips": [
        "Print forms are always in Thai.",
        "Sheets whose room boxes were typed over before this update still show one combined block as typed — tap ↺ in the room boxes and save to use the block layout.",
        "With unsaved edits, tap “Save changes” before printing or saving a file.",
        "On iPad / iPhone, after Save choose “Save to Files” or share it."
      ]
    },
    {
      "id": "cal",
      "group": "event",
      "icon": "cal",
      "title": "Booking calendar",
      "who": "Sales · Managers · Heads of Department · Accounting",
      "intro": "The Calendar menu shows all events by date, color-coded by status.",
      "steps": [
        "Tap ‹ › to change the month, or “Today” to return to the current month.",
        "Switch between the Month · Week · Pipeline views.",
        "Filters on the left: status (Draft · Provisional · Confirmed) with counts and totals, and room/area — colours: draft grey · provisional yellow · confirmed green.",
        "Tap an event name in the calendar to open its function sheet (or the quotation if there is no function sheet yet)."
      ],
      "shots": [
        {
          "img": "cal-month",
          "cap": "Monthly calendar"
        },
        {
          "img": "cal-week",
          "cap": "Weekly view"
        },
        {
          "img": "cal-pipe",
          "cap": "Pipeline view"
        }
      ],
      "tips": []
    },
    {
      "id": "shifts",
      "group": "people",
      "icon": "clock",
      "title": "Shifts and shift extensions (QT)",
      "who": "Heads of Department · Managers · Administrator",
      "intro": "Staff › Shifts is where you put people on shifts for each event. The system works out how many staff are needed from the number of guests.",
      "steps": [
        "Choose the event at the top (tap the quotation name ▾) and choose the day.",
        "If there is no plan yet, tap “Create plan from the quotation” — shifts are made only for departments with “Apply rules” on in Head-count rules, with people per shift by the guest count (a summary shows first, e.g. Kitchen/F&B 1 per shift) — “Apply rules to this plan” works out an existing plan again.",
        "Departments without rules: add shifts yourself with “+ Add shift”.",
        "Check the summary boxes: Staff needed · Assigned · Short-staffed · Extend shift.",
        "Department timelines: tap a shift (Morning shift/Afternoon shift) to set the number needed and the times, tick the staff, then tap “Save”.",
        "Extend shift (QT): choose +1/+2/+3 hrs or a custom time, enter a reason, choose the staff, then send it for approval.",
        "People with approval rights tap Approve / Send back on shift extension requests."
      ],
      "shots": [
        {
          "img": "staff-shifts",
          "cap": "Shifts: summary, timeline and shift extensions"
        },
        {
          "img": "staff-assign",
          "cap": "Assigning staff to a shift"
        }
      ],
      "tips": [
        "Hours worked this week are shown next to each name (e.g. 40 / 48 hrs) to help avoid overtime."
      ]
    },
    {
      "id": "emps",
      "group": "people",
      "icon": "users",
      "title": "All staff, locations, departments and positions",
      "who": "People who can plan staff shifts (add/edit) · everyone who sees the menu (view)",
      "intro": "Staff › All staff holds the staff list for every location: Resort · Boutique · มณีอินทร์ · ช.ดรักส์ · ช.เฮ้าส์ · ภูมณีพันธุ์.",
      "steps": [
        "Choose a location from the red pin bar (or “All locations”), then filter by department or search by name, position or phone number.",
        "Tap “+ Add staff”, fill in the full name, choose the location, department and position, enter the phone number and hours worked, then tap “Add staff”.",
        "Tap ✏️ to edit or 🗑 to delete (the system also removes the person from their shifts).",
        "The + button next to the Place / Department / Position fields: add a new one, rename it (all staff using it are updated) or delete it (only when no staff use it).",
        "Import from Excel/CSV: tap “Import staff list” › download the template › choose the file › check the table › choose “Add to the current list” or “Replace all at the locations in the file” › tap Import.",
        "In the import file, the location column is optional (left blank = the location currently selected), and departments not yet in the system are added as new departments."
      ],
      "shots": [
        {
          "img": "staff-emps",
          "cap": "All staff by location and department"
        },
        {
          "img": "staff-plus",
          "cap": "The + button for managing locations, departments and positions"
        },
        {
          "img": "staff-import",
          "cap": "Importing a staff list from an Excel / CSV file"
        }
      ],
      "tips": [
        "The 6 main departments can be renamed but not deleted · Resort cannot be deleted.",
        "Adding, renaming or deleting locations, departments and positions is saved as soon as you tap."
      ]
    },
    {
      "id": "rules",
      "group": "people",
      "icon": "sliders",
      "title": "Head-count rules",
      "who": "People who can plan staff shifts",
      "intro": "Staff › Head-count rules sets how many staff are needed per shift for a given number of guests.",
      "steps": [
        "Turn on “Apply rules” for the departments you want (off = no shifts for that department when a plan is made; add them yourself later).",
        "Set guest ranges, for example 1–50 guests = 1 staff, 51–100 guests = 2 staff.",
        "Tap “+ Add guest range” or 🗑 to delete a range, then save.",
        "The rules are used with “Create plan from the quotation” or “Apply rules to this plan”."
      ],
      "shots": [
        {
          "img": "staff-rules",
          "cap": "Head-count rules per shift"
        }
      ],
      "tips": []
    },
    {
      "id": "roster",
      "group": "people",
      "icon": "cal",
      "title": "Work roster (monthly shift table)",
      "who": "Heads of Department (own department) · HR · Administrators (edit) · people who see the Staff menu (view)",
      "intro": "Staff › Work roster replaces the paper shift table. Each department has its own table in the same form; each period runs from the 26th to the 25th, for the selected property.",
      "steps": [
        "Choose the department in the bar at the top (✓ = already has this period’s table) and tap ‹ › to change period — heads of department open on their own department.",
        "Tap “Create this period’s table” — it uses the names and leave balances from the last period, or the department’s staff list.",
        "Fill in shifts: pick a code in the bar (ช6 · ช7 · ช8 · บ2 · ค4 · DAY OFF · EXTRA · CDO · WOP · VC · clear) and tap several cells in a row, or tap one cell to choose — CDO asks for the day taken off instead (dd/mm).",
        "Top rows: Room Forecast OCC% (filled in by HR) and Function Forecast/pax (the guests of the quotations booked that day, added up by the system).",
        "Worked out at the end of each row: working days · DAY OFF (with CDO) · EXTRA · total · leave left · leave used — HR taps “leave left” to enter the opening balance; the next period carries it over.",
        "Rows: “+ Add row” (stand-ins, extra shifts) · “+ Add staff not yet in the table” · tap a name to rename it, add a note under it or delete the row · tap the code cell to enter the staff code.",
        "Sign in order: Head (that department’s head or an administrator) → HR → Manager (administrator), using the “Sign (…)” button in the signature boxes.",
        "Tap “Print / PDF” — one A4 landscape page in colour, with the legend and signature boxes.",
        "Administrators: the “Shift codes” button adds or edits codes and shift times."
      ],
      "shots": [
        {
          "img": "roster",
          "cap": "Work roster: department tabs, shift-code bar and the table"
        }
      ],
      "tips": [
        "Editing the table after someone has signed clears the signatures; sign again in order.",
        "Heads edit and sign only their own department’s table · HR means accounts with “See leave details (HR)”.",
        "Combined tables made before the split by department: tap “Create this period’s table” in each department and the system copies that department’s names and shifts."
      ]
    },
    {
      "id": "tasks",
      "group": "people",
      "icon": "check",
      "title": "Assignment & Approve (assigned tasks and approvals)",
      "who": "Everyone (receive tasks) · Administrators and roles given permission (assign tasks) · Administrator (Approve · ⚙ permissions)",
      "intro": "Assign tasks to users or whole departments, track their status, and approve edited function sheets.",
      "steps": [
        "Assignment tab: see all tasks and filter by Pending / Confirmed — urgent tasks have an “Urgent” label and overdue tasks are red.",
        "Tap “+ Assign task”, enter the task name, details, assignee (a user or a department), due date and urgency. You can link it to a quotation.",
        "Assigning to a department: tap a location button (All locations · Resort · Boutique · มณีอินทร์ · ช.ดรักส์ · ช.เฮ้าส์ · ภูมณีพันธุ์) — the number on the button is how many staff that department has at that location — then tick the staff (none ticked = the whole department). The location is shown in the task list.",
        "The ⚙ button next to the heading (administrator): Assignment permissions, 2 ways — “Org chart” (a role higher up can assign work to every role and department below it; change “Reports to” to move one; − + Fit 100% buttons to zoom; and ✏️ rename, ＋ add a role/department, 🗑 delete on each box) or “Selection form” (choose, for each role, which roles and departments it can assign work to) — tap Save and the chosen way is used.",
        "Roles given permission see the “+ Assign task” button (only the people and departments they are allowed to pick) and a “Tasks I assigned” list, where they can edit/delete only their own tasks, and they are notified when the assignee confirms.",
        "The assignee taps the task on the Home page or in the bell to see the details, then taps “Confirm” with a note, such as “All set up”.",
        "Approve tab (administrator): function sheets awaiting approval. Tap “View” to compare the changes, then “Approve” or send it back.",
        "Approved sheets can be sent to departments and move to the “Approved” section."
      ],
      "shots": [
        {
          "img": "tasks",
          "cap": "Assigned task list and the ⚙ permissions button"
        },
        {
          "img": "task-new",
          "cap": "Assigning a new task"
        },
        {
          "img": "task-site",
          "cap": "Assigning to a department: choosing the location and staff"
        },
        {
          "img": "task-perms-org",
          "cap": "Assignment permissions: Org chart"
        },
        {
          "img": "task-perms-form",
          "cap": "Assignment permissions: Selection form"
        },
        {
          "img": "task-view",
          "cap": "Details of a received task"
        },
        {
          "img": "tasks-approve",
          "cap": "Approve: function sheets awaiting approval"
        }
      ],
      "tips": [
        "Tasks assigned to a department can be seen and confirmed by everyone in that department.",
        "Until permissions are set, only administrators can assign tasks · Administrators can always assign work to anyone.",
        "In the chart, you can only delete roles or departments you added yourself that have nobody in them · Changes to the chart take effect when you tap Save."
      ]
    },
    {
      "id": "leave",
      "group": "people",
      "icon": "leave",
      "title": "Leave Requests (QT · Change day off · Missed fingerprint scan)",
      "who": "Everyone (submit) · Administrator (sign/approve) · Human Resources (details)",
      "intro": "The Leave Requests menu replaces paper form HRSD08.1. There are 3 types: QT request · Change day off · Missed fingerprint scan.",
      "steps": [
        "Tap “+ New” on the card for the type you need.",
        "Fill in your full name, nickname, position, department and branch (filled in from your account; you can edit them).",
        "QT request: choose Work QT or Take QT (use QT), the start–end time and date. The system works out the hours for you.",
        "Change day off: choose Change day off or Change working time, from date… to date….",
        "Missed fingerprint scan: choose the reason, the in/out time, the date and a witness.",
        "Enter a reason, then tap “Submit”.",
        "Track the status in the My requests · Awaiting approval · All tabs. Tap a request to see the signing steps (Head of Department · General Manager · Deputy General Manager · Human Resources).",
        "While nobody has signed a request yet, its owner can edit ✏️ or delete 🗑 it · Tap “Print form” to print the form.",
        "Human Resources/administrators: the “Details” tab lets you search, filter by month, type, status and department, see a summary per person (Work QT/Take QT/QT balance) and export to Excel."
      ],
      "shots": [
        {
          "img": "leave",
          "cap": "Leave Requests page and My requests list"
        },
        {
          "img": "leave-qt",
          "cap": "QT request form"
        },
        {
          "img": "leave-change",
          "cap": "Change day off form"
        },
        {
          "img": "leave-scan",
          "cap": "Missed fingerprint scan form"
        },
        {
          "img": "leave-view",
          "cap": "Request details and signing steps"
        },
        {
          "img": "leave-info",
          "cap": "Details for Human Resources"
        }
      ],
      "tips": [
        "QT must be at least 2 hours · QT may not be worked before the start of your shift.",
        "Only administrators can sign and approve leave requests."
      ]
    },
    {
      "id": "users",
      "group": "admin",
      "icon": "shield",
      "title": "Users, roles and permissions",
      "who": "Administrator · Sales Manager (with the Manage users permission)",
      "intro": "The Users menu is where you create accounts, set roles and turn individual permissions on or off.",
      "steps": [
        "Search or filter by role: Sales Staff · Sales Manager · Head of Department · Accounting · Administrator.",
        "Tap “+ Add user”, enter the name, email (used to sign in), phone, role and department, then save. Tell the user their starting password and ask them to change it the first time they sign in.",
        "Tap a user’s name to open the panel on the right and turn each permission on or off, such as Create and send quotations · Discounts up to 10% · Approve discounts over 10% · View function sheets and the booking calendar · Manage customer data · Plan staff shifts · Manage users · Edit prices and packages · View leave request details (HR).",
        "Tap ✏️ to edit details, change the status (Active / Disabled), set up two-step verification or reset the password.",
        "“Properties allowed” (administrators): in the right-hand panel choose Luck Swan Resort · Luck Swan Boutique · Manee Inn (at least one), then tap “Save” — administrators can always open every property.",
        "Roles: tap “+ Add role”, enter a name, choose the visible menus and permissions, then save — saving a role updates the permissions of everyone in that role (you can also add/rename/delete roles from the org chart in Assignment ⚙)."
      ],
      "shots": [
        {
          "img": "users",
          "cap": "User list and permissions panel"
        },
        {
          "img": "user-edit",
          "cap": "Editing user details"
        },
        {
          "img": "role",
          "cap": "Individual permissions"
        }
      ],
      "tips": [
        "A browse-only role, such as junior sales: untick “Create and send quotations” on the role — the meeting room, room and catering pages are then for viewing only, with no buttons that add to a quotation.",
        "Accounts without set properties can open every property · users cannot see or change data of properties they are not allowed.",
        "When you change someone’s role, their permissions change to match the new role (you can fine-tune them afterwards).",
        "Quotations already sent do not change when permissions change."
      ]
    },
    {
      "id": "dash",
      "group": "admin",
      "icon": "chart",
      "title": "Dashboard sales reports",
      "who": "Administrator · Management",
      "intro": "The Dashboard menu summarizes revenue by event date, by year and by month.",
      "steps": [
        "Choose the year (‹ year ›) and a month, or “Whole year”.",
        "Summary boxes: confirmed revenue for the period · compared with last year · awaiting confirmation (Awaiting approval + Tentative) · % closed.",
        "Monthly revenue chart (Confirmed / awaiting confirmation): point at or tap a dot to see the figures, or tap “View as table”.",
        "You can export the data as a CSV file."
      ],
      "shots": [
        {
          "img": "dash",
          "cap": "Dashboard of yearly and monthly revenue"
        }
      ],
      "tips": []
    },
    {
      "id": "backup",
      "group": "admin",
      "icon": "save",
      "title": "Backups and system updates",
      "who": "Administrator",
      "intro": "All data is stored in a Cloudflare D1 database and syncs to every device automatically.",
      "steps": [
        "Back up: tap your profile picture › “Back up data (download)” and keep the file somewhere safe.",
        "Restore: use “Import backup” and choose the file you downloaded (only when really needed).",
        "Back up everything (database, photos, user accounts and code): GitHub › Actions › “Backup data” › Run workflow, then download the .7z file under Artifacts (locked with the BACKUP_PASSWORD set in GitHub) — how to restore is in RESTORE.md inside it.",
        "If no administrator can sign in: GitHub › Actions › “Reset admin account” (put the e-mail and a new password in the secrets ADMIN_RESET_EMAIL / ADMIN_RESET_PASSWORD first), then change the password again in the app.",
        "System updates: when new features are released, the app on every device updates itself within a few minutes."
      ],
      "shots": [],
      "tips": [
        "Never send passwords or API tokens by chat or email.",
        "If everyone cannot sign in at once and see “This user was not found”, the free daily database quota on Cloudflare may be used up — wait until 07:00 (Thailand) and refresh; do not save work meanwhile."
      ]
    }
  ],
  "quick": [
    {
      "id": "q-sales",
      "title": "Sales (sales staff)",
      "color": "#0E9F8E",
      "menus": "Home · Rooms & Venues · Catering · Quotations · Function Sheets · Calendar · Leave Requests · Assignment",
      "can": [
        "Create and send quotations",
        "Discounts up to 10%",
        "View function sheets and the booking calendar",
        "Manage customer data"
      ],
      "steps": [
        "Tap “+ New quotation” › complete the 5 steps: Customer (you can enter the Tax ID) › Meeting rooms › Food › Guest rooms › Summary.",
        "Open the quotation › adjust items, discount and terms › tap Send / Print › choose Customer copy / Full quotation › you can edit details on the quotation › tap “Request approval” › print it for an administrator to sign.",
        "Wait for an administrator to “Approve quotation” (status becomes Tentative) › email / LINE it to the customer › the customer confirms › tap “Customer confirmed booking” › Function sheets › “+ Create from quote” › choose attendees and type of event.",
        "Fill in the daily schedule and each department’s tasks › follow up on acknowledgements.",
        "Check the calendar before taking a new event, to see free rooms · design the floor plan from the function sheet (the tables / rows panel works it out from the guests).",
        "Function details › “For operations” › “Save” as JPG/PDF and send it to the operations team on LINE."
      ],
      "shots": [
        "wiz1",
        "quote-edit",
        "beo"
      ],
      "links": [
        "newquote",
        "quotes",
        "sendquote",
        "beo",
        "cal"
      ]
    },
    {
      "id": "q-manager",
      "title": "Sales Manager",
      "color": "#A8262D",
      "menus": "All Sales menus + Staff · Users",
      "can": [
        "Everything Sales can do",
        "Approve discounts over 10%",
        "Plan staff shifts / approve shift extensions",
        "Manage users"
      ],
      "steps": [
        "Review quotations with discounts over 10% and approve them.",
        "Check the Home page: monthly target and latest quotations.",
        "Assign work to the team: Assignment › “+ Assign task” (once an administrator gives permission in ⚙).",
        "Approve shift extension (QT) requests on the Shifts page.",
        "Add/edit users in the sales team and set their permissions."
      ],
      "shots": [
        "quote-edit",
        "tasks",
        "users"
      ],
      "links": [
        "quotes",
        "tasks",
        "shifts",
        "users"
      ]
    },
    {
      "id": "q-ops",
      "title": "Operations heads of department (Kitchen/F&B · Banquet · Engineering/AV · Front Office · Housekeeping)",
      "color": "#D4A21A",
      "menus": "Home · Function Sheets · Calendar · Staff · Leave Requests · Assignment",
      "can": [
        "View function sheets and the booking calendar",
        "Plan staff shifts / approve shift extensions"
      ],
      "steps": [
        "Check the bell every morning: new or edited function sheets › open the sheet › tap “Acknowledge” on your department’s card.",
        "Check the daily schedule, attendees, type of event and floor plan to prepare items and staff (F&B and banquet can design the plan: set rows/tables, rotate, add text boxes, or use a floor-plan picture as the plan).",
        "Edit the tasks in your department’s card — your department’s box in Function details changes at once.",
        "Staff › Shifts › choose the event › assign staff to shifts › request shift extensions when needed.",
        "Keep your department’s staff list up to date (add/edit/import).",
        "Staff › Work roster › your department › fill in shifts for the period (26–25) › sign the “Head” box.",
        "Tasks assigned to your department › open them › tap “Confirm” when everything is ready."
      ],
      "shots": [
        "beo",
        "staff-shifts",
        "roster"
      ],
      "links": [
        "beo",
        "shifts",
        "roster",
        "emps",
        "tasks"
      ]
    },
    {
      "id": "q-acc",
      "title": "Accounting",
      "color": "#2B3F86",
      "menus": "Home · Quotations (view) · Function Sheets · Calendar · Leave Requests",
      "can": [
        "View function sheets and the booking calendar",
        "Manage customer data"
      ],
      "steps": [
        "Review confirmed quotations: net total, deposit terms and customer details.",
        "Function Sheets › “Sales summary”: check revenue, food, guest rooms and tax invoice details › Print / PDF.",
        "Function details › “Accounting / Finance” section: see the total and cost details.",
        "Acknowledge the “Accounting” card on the function sheet.",
        "Calendar › Pipeline view to see amounts awaiting confirmation."
      ],
      "shots": [
        "ss",
        "quotes",
        "cal-pipe"
      ],
      "links": [
        "quotes",
        "fnforms",
        "cal"
      ]
    },
    {
      "id": "q-hr",
      "title": "Human Resources (HR)",
      "color": "#7B4FA0",
      "menus": "Leave Requests (Details) · Staff",
      "can": [
        "View leave request details (HR)",
        "View the staff list"
      ],
      "steps": [
        "Leave Requests › “Details” tab.",
        "Filter by month, type, status or department, or search by name/nickname/number.",
        "See the summary per person: Work QT · Take QT · QT balance · Change day off · Missed fingerprint scan · Awaiting approval.",
        "Tap “Export to Excel” for payroll/reports.",
        "Check each location’s staff list in the Staff menu.",
        "Staff › Work roster: enter OCC% and opening leave balances, edit any department › sign the “HR” box."
      ],
      "shots": [
        "leave-info",
        "roster",
        "staff-emps"
      ],
      "links": [
        "leave",
        "roster",
        "emps"
      ]
    },
    {
      "id": "q-admin",
      "title": "Administrator",
      "color": "#17211F",
      "menus": "All menus",
      "can": [
        "All permissions",
        "Approve quotations",
        "Sign and approve leave requests",
        "Approve edited function sheets and overlapping room bookings",
        "Edit guest room rates and menus",
        "Manage users, roles and permissions"
      ],
      "steps": [
        "Home › quotations “Awaiting approval” › check the copy the salesperson printed for you to sign › tap “Approve quotation” (it becomes Tentative automatically) or “Send back” with a reason.",
        "Assignment & Approve › Approve tab: approve edited function sheets · ⚙ Assignment permissions (org chart / form).",
        "Leave Requests › Awaiting approval tab › open the request › Sign / Reject.",
        "Users: create accounts, set roles, permissions and allowed properties, turn users on/off.",
        "Work roster: sign the “Manager” box · set the shift codes.",
        "Set guest room rates, edit meeting rooms (✏️ photos, prices, capacity) and update catering menus when prices change.",
        "Dashboard: follow sales · back up now and then (GitHub › Actions › Backup data).",
        "Junior sales role: Users › Roles › untick “Create and send quotations” so they can only look at rooms and menus."
      ],
      "shots": [
        "quote-approve",
        "tasks-approve",
        "leave-view",
        "dash"
      ],
      "links": [
        "quotes",
        "tasks",
        "leave",
        "users",
        "roster",
        "meet",
        "rooms",
        "food",
        "dash",
        "backup"
      ]
    },
    {
      "id": "q-all",
      "title": "All staff",
      "color": "#6B716E",
      "menus": "Home · Leave Requests · Assigned tasks",
      "can": [
        "Submit a QT request · Change day off · Missed fingerprint scan",
        "Receive and confirm assigned tasks",
        "Change your own language and color theme"
      ],
      "steps": [
        "Sign in › change the language with the TH/EN/MY button.",
        "Leave Requests › choose the type › “+ New” › fill in the form › “Submit”.",
        "Check the status in the “My requests” tab. You can edit or delete a request until someone signs it.",
        "Assigned tasks: tap one on the Home page or in the bell › read it › tap “Confirm” when done."
      ],
      "shots": [
        "leave-qt",
        "task-view",
        "lang"
      ],
      "links": [
        "login",
        "langtheme",
        "leave",
        "tasks"
      ]
    }
  ],
  "changelog": [
    {
      "v": "1.5",
      "date": "Nov 10, 2026",
      "text": "Quotations: “Awaiting approval” replaces “Sent” — request approval › an administrator approves or sends it back › sending to the customer and confirming the booking only after approval › the function sheet after the customer confirms · Floor plan: saved when you tap “Save floor plan” (Ctrl+S), changes can be discarded, a prompt before leaving, and unsaved plans are restored · Function details: the front office section has one block per guest-room line."
    },
    {
      "v": "1.4",
      "date": "Nov 5, 2026",
      "text": "Floor plans: tables / rows panel worked out from the guests and adjustable · room picker when an event uses several rooms · text box · rotate handle and 15°/45°/90°/180° buttons · a floor-plan picture as background or as the plan · Function details: For sales / For operations buttons · Save as PNG / JPG / PDF · department cards and the sheet as one list · type of event from the tags · Quotations: Owner column · Staff shifts: plans only for departments with rules · browse-only roles (junior sales) · backup and admin recovery through GitHub"
    },
    {
      "v": "1.3",
      "date": "Oct 29, 2026",
      "text": "Choose a property after signing in (Resort / Boutique / Manee Inn), data separated by property, and allowed properties per person · Quotations: “Approve quotation” (administrators) makes it provisional automatically, yellow/green box, “Customer confirmed the booking” button, provisional in yellow, and room photos on the document · function sheets only after the quotation is approved · Floor plans: open to sales · Catering: add/rename/delete categories · Meeting rooms: full-screen photos · Staff: monthly work roster by department, signed Head → HR → Manager"
    },
    {
      "v": "1.2",
      "date": "Oct 26, 2026",
      "text": "Guest rooms: choose the place, Resort / Boutique (Standard · Superior · Deluxe rooms with extra breakfast and dinner party), and Boutique rates · Room rate text states net prices · Function sheets: the Attendees and type of event box (choose by branch, edit/add/delete choices)."
    },
    {
      "v": "1.1",
      "date": "Oct 22, 2026",
      "text": "Meeting rooms: ✏️ button to edit photos, prices and capacity, and swipe through sample photos · Quotations: Net prices with the service charge/VAT worked back out, Tax ID field, Customer copy / Full quotation, editing details on the quotation, changing the meeting room photo, and the 3% credit card term · Function sheets: permissions for editing function details, the Accounting/Finance section and floor plans (F&B · Banquet) · Assignment: choose a location when assigning to a department, and ⚙ Assignment permissions (an org chart with zoom and rename/add/delete, or a form)."
    },
    {
      "v": "1.0",
      "date": "Oct 21, 2026",
      "text": "First edition, covering every chapter: quotations, meeting rooms, guest rooms, catering, function sheets, floor plans, calendar, staff (locations, departments, positions, import), shifts, assigned tasks, leave requests, users, Dashboard, 3 languages and color themes."
    }
  ]
};
