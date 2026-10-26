/* Luck Swan Resort sales system user manual — English (US) content (translated from content.th.js, the master version)
 * To edit or add a chapter: add an object to chapters (or to quick for the short guides by department)
 * in content.th.js, then add the same id here and in content.my.js — ids must match in all 3 languages
 * Images: img/<image name>-<th|en|my>.webp (screenshots of the real app in all 3 languages) */
window.MANUAL = window.MANUAL || {};
window.MANUAL.en = {
  ui: {
    lang: 'Language', full: 'Full manual', quick: 'Quick guides by department', contents: 'Contents', search: 'Search the manual…',
    who: 'Used by', steps: 'Steps', tips: 'Good to know', print: 'Print / Save PDF', top: 'Back to top',
    version: 'Version', updated: 'Last updated', changelog: 'Manual change history', menu: 'Menus', can: 'What you can do',
    daily: 'Daily work', open: 'Open the full chapter', noresult: 'No matching topics found', zoom: 'Tap an image to enlarge'
  },
  meta: {
    title: 'Sales System User Manual',
    subtitle: 'Luck Swan Resort — quotations, meeting rooms, guest rooms, catering, function sheets, staff, leave requests and assigned tasks',
    url: 'https://luckswan-sales.it-luckswan.workers.dev',
    version: '1.1', updated: 'Oct 22, 2026',
    intro: 'The sales system is a web app that works on computers, iPads and phones. Everyone sees the same data in real time. The menus you see depend on the role and permissions the administrator gives you. This manual covers every chapter, from signing in to each department’s work, with real screenshots for every step.'
  },
  groups: [
    { id: 'start', title: 'Getting started', color: '#A8262D' },
    { id: 'sell', title: 'Sales', color: '#0E9F8E' },
    { id: 'event', title: 'Catering and function sheets', color: '#D4A21A' },
    { id: 'people', title: 'Staff and internal work', color: '#2B3F86' },
    { id: 'admin', title: 'Administrator', color: '#17211F' }
  ],
  chapters: [
    {
      id: 'login', group: 'start', icon: 'key', title: 'Signing in and installing on your device', who: 'Everyone',
      intro: 'Open luckswan-sales.it-luckswan.workers.dev in a browser (Safari, Chrome or Edge) and sign in with the account the administrator created for you.',
      steps: [
        'Type your email or employee ID in the first field.',
        'Type your password (at least 6 characters). Tap “Show” to see what you typed.',
        'Turn on “Remember me on this device” if it is your own device (leave it off on shared devices).',
        'Tap “Sign in”. The first time, you may be asked to set a new password.',
        'Forgot your password? Tap “Forgot password?” or call the administrator at 053-750-855.',
        'Install as an app on a phone/iPad: tap your profile picture at the top right › “How to install on phone / tablet” and follow the steps (Safari: Share button › Add to Home Screen).'
      ],
      shots: [{ img: 'login', cap: 'Sign-in page — the language and color theme buttons are at the top right' }],
      tips: ['The language and color theme buttons work right from the sign-in page.', 'The app updates itself automatically. If an update was just released, close the app and open it again.']
    },
    {
      id: 'home', group: 'start', icon: 'home', title: 'Home page and menu bar', who: 'Everyone',
      intro: 'After you sign in, you see the Home page. The menu bar is at the top, and shortcuts and a work summary are in the middle.',
      steps: [
        'Top menu bar: Home · Rooms & Venues ▾ (Meeting Rooms, Guest Rooms) · Catering · Quotations · Function Sheets · Calendar · Staff · Users · Leave Requests · Assignment & Approve · Dashboard (you only see the menus you have permission for).',
        'On a phone, the main menus are in the bottom bar. Tap ☰ at the top left to see all menus.',
        'Colored cards at the top: shortcuts to meeting rooms and events · group guest rooms · catering and food · seminar packages with accommodation.',
        'The “+ New quotation” button at the top right starts a quotation right away.',
        'In the “Latest quotations” table, tap a row to open that quotation.',
        'The “… target” box shows closed sales against the monthly target, and the “Assigned tasks” box shows your latest tasks.',
        'The search box at the top right finds customers or quotation numbers from any page.'
      ],
      shots: [{ img: 'home', cap: 'Home page: shortcuts, latest quotations, sales target and assigned tasks' }, { img: 'usermenu', cap: 'Profile menu: change password, how to install, back up data, log out' }],
      tips: ['The profile menu has: Change password · How to install on phone / tablet · Back up data (download) · Import backup · Log out.']
    },
    {
      id: 'langtheme', group: 'start', icon: 'globe', title: 'Changing the language and color theme', who: 'Everyone',
      intro: 'Make the screen easy for you to read. Settings are saved on the device you use and do not affect anyone else.',
      steps: [
        'Change the language: tap the letter button (labeled TH/EN/MY) next to the bell, then choose ไทย (default) · English (US) · မြန်မာ.',
        'Change the color theme: tap the seven-color circle button next to the language button.',
        'Pick one of the 8 “Ready-made themes”, such as Standard, Ocean, Emerald or Royal purple — or tap a color circle under “Any color”.',
        'Mix your own: tap the Main color / Background circle to open the color picker, or type a color code such as #1E88E5.',
        'Eyedropper: tap the eyedropper icon, then tap the color you want on the screen.',
        'Tap “Back to the standard colors” to return to the original colors.'
      ],
      shots: [{ img: 'lang', cap: 'Language menu with 3 languages' }, { img: 'theme', cap: 'Color theme menu: ready-made themes, any color, mix your own, color code and eyedropper' }],
      tips: ['Text people type themselves (names, customer names, reasons) is shown as typed and is not translated.', 'Documents printed or sent to customers are always in Thai with the standard colors.']
    },
    {
      id: 'notifs', group: 'start', icon: 'bell', title: 'Notifications', who: 'Everyone',
      intro: 'The bell at the top right shows things that need your attention. The red number is how many you have not read yet.',
      steps: [
        'Tap the bell to see your notifications, such as function sheets awaiting approval, tasks assigned to you, or leave requests waiting to be signed.',
        'Tap a notification to go straight to the related page.',
        'Tap “View all tasks” to open the Assignment & Approve page.'
      ],
      shots: [{ img: 'notifs', cap: 'Notifications window' }],
      tips: []
    },
    {
      id: 'meet', group: 'sell', icon: 'meet', title: 'Meeting Rooms', who: 'Sales · Sales Manager · Administrator',
      intro: 'Rooms & Venues › Meeting Rooms shows free rooms, capacity and prices, and lets you add a room to a quotation.',
      steps: [
        'Set the options at the top: event date · number of attendees · Full day / Half day.',
        'Choose a room layout: Theatre · Classroom · U-shape · Banquet · Cocktail — the system works out the capacity for that layout.',
        'Check the room cards: size, capacity, full-day price and status, such as “Free all 2 days”, or a red bar showing dates that are already booked.',
        'Tap a room card to see its details on the right: room photos, seating plan, room conditions and equipment.',
        'Tap “Choose this room” to add it to the quotation you are working on (the quotation number “For QT-…” shows under the heading; tap “Change” to switch quotations).',
        'Room photos: swipe left/right or tap the ‹ › arrows to see the next photo, both on the room card and in the panel on the right (a counter such as 2/3 shows which photo you are viewing).',
        'The pencil button ✏️ at the bottom right of a room card (people who can edit prices): edit the sample photos (upload several, delete, reorder — the first photo is the cover), the full-day/half-day prices, the capacity for each layout (min–max, such as 150–300 guests), the area, the banquet minimum guarantee, the description and what is in the room, then tap “Save” — there is also a “Reset to default” button.'
      ],
      shots: [{ img: 'meet', cap: 'Meeting room list with capacity and availability' }, { img: 'meet-sel', cap: 'Room selected: seating plan, conditions and equipment' }, { img: 'meet-edit', cap: 'Editing a meeting room: photos, prices and capacity' }],
      tips: ['If a booking overlaps another event, the system warns you and an administrator must approve it first.', 'Rooms that are over capacity for the chosen layout show a warning.', 'A new room price applies to items added to quotations from now on; existing quotations do not change.']
    },
    {
      id: 'rooms', group: 'sell', icon: 'bed', title: 'Guest Rooms and room rates', who: 'Sales · Administrator (rates)',
      intro: 'Rooms & Venues › Guest Rooms is where you set up a group’s room block. Prices follow the season automatically.',
      steps: [
        'Choose the check-in date and number of nights at the top, then choose With ABF (breakfast included) or Room only.',
        'On each room type card, tap − / + to set the number of rooms and choose the bed type (Twin / Double / Triple for 3 guests).',
        'Check the room block summary on the right: number of rooms, guests and free rooms for the organiser (1 free for every 20 rooms).',
        'Turn Late check-out 14:00 on or off as agreed with the customer.',
        'Check the release date (cut-off) in the yellow box. Rooms without a guest list by then go back to the system automatically.',
        'Tap “Update quote” to save the guest rooms to the quotation.',
        'Administrators: tap “Room rates” to edit the rate periods (seasons), number of rooms and the With ABF / Room only price for each type, then save.'
      ],
      shots: [{ img: 'rooms', cap: 'Room block: choose room types, quantities and bed types' }, { img: 'roomrates', cap: 'Setting room rates by season (administrator)' }],
      tips: ['Room rates are per room per night for 2 guests (Triple: 3 guests), net prices with service charge and VAT included (worked out the same way as the quotation).']
    },
    {
      id: 'food', group: 'sell', icon: 'food', title: 'Catering and menus', who: 'Sales · Administrator (menu editing)',
      intro: 'The Catering menu covers seminar packages, coffee breaks, lunch, gala dinners, cocktails and weddings.',
      steps: [
        'Choose a category on the left, or filter by price range and dietary suitability (Halal · Vegetarian · Gluten-free).',
        'Check the menu cards: photo, items in the package, price per person and minimum.',
        'Tap “+ Add to quote” to add it to the quotation you are working on — the bottom bar shows the total and a “Back to quote” button.',
        'If you can edit prices: tap “Edit” on a card to change the name, price, minimum, description, suitability labels, main photo and package items.',
        'Tap “+ Add menu / package” to add a new menu.'
      ],
      shots: [{ img: 'food', cap: 'Catering menu list and filters' }, { img: 'food-edit', cap: 'Editing a package: price, minimum, photo and food items' }],
      tips: ['New prices apply to quotations made from now on. Quotations already sent do not change.']
    },
    {
      id: 'quotes', group: 'sell', icon: 'doc', title: 'Quotations: list and editing', who: 'Sales · Sales Manager · Accounting (view)',
      intro: 'The Quotations menu lists every quotation. Search and filter by status: Draft · Sent · Tentative · Confirmed · Cancelled.',
      steps: [
        'Search by customer name or quotation number, then tap a row to open it.',
        'On the quotation page, items are grouped into Meeting Rooms · Food & beverage · Guest Rooms. Edit quantities or prices right in the fields, and tap “+ Add item…” to add more.',
        'Switch between the “Table / Timeline” views at the top right to see the daily schedule.',
        'Summary box on the right: prices on the quotation are Net prices (service charge and VAT included) — subtotal before discount · discount · price before service charge and VAT · 10% service charge · total before VAT · 7% VAT · net total, plus the average per person.',
        'Discounts over 10% need the sales manager’s approval.',
        'At the bottom, set Price valid until · Deposit (for example 30% within 7 days of confirmation) · tentative room hold until a date.',
        'Tap “Edit customer details and terms” to change the customer name, contact person, phone and email.',
        'Change the quotation status (Sent / Tentative / Confirmed / Cancelled) and see the edit history at the bottom of the page.'
      ],
      shots: [{ img: 'quotes', cap: 'Quotation list with status filters' }, { img: 'quote-edit', cap: 'Quotation edit page and summary box' }],
      tips: ['The system saves automatically every time you make a change.', 'A confirmed quotation can be turned into a function sheet.', 'How the total is broken down: price before SC/VAT = net total ÷ 1.177 → service charge is 10% of that price → total before VAT → 7% VAT. For example, 1,200 = 1,019.54 + 101.95 + 78.51.']
    },
    {
      id: 'newquote', group: 'sell', icon: 'plus', title: 'Creating a new quotation (5 steps)', who: 'Sales · Sales Manager',
      intro: 'Tap “+ New quotation” and the system walks you through it step by step. The net total (including service charge + VAT) is always shown on the right.',
      steps: [
        'Step 1 Customer: fill in the event name* and customer/company*, Tax ID (13 digits, optional), contact person, phone, email, start–end dates and number of attendees, then tap “Next”.',
        'Step 2 Meeting rooms: choose a room layout and tick the rooms (the system shows capacity and price). You can choose half day/full day and an AV package.',
        'Step 3 Food: turn each meal on or off (morning coffee break · lunch · afternoon coffee break · gala dinner · cocktail), then choose menus. You can add dietary requirements.',
        'Step 4 Guest rooms: set the number of nights and the number of rooms of each type.',
        'Step 5 Summary: check the items and total, add a discount, then tap “Create quotation”.',
        'Use the back button (it shows the previous step’s name) to change an earlier step at any time, or “Cancel” to stop.'
      ],
      shots: [{ img: 'wiz1', cap: 'Step 1 Customer and event details' }, { img: 'wiz2', cap: 'Step 2 Choosing meeting rooms' }, { img: 'wiz3', cap: 'Step 3 Food & beverage' }, { img: 'wiz4', cap: 'Step 4 Guest rooms' }, { img: 'wiz5', cap: 'Step 5 Summary before creating the quotation' }],
      tips: ['Fields marked * are required. You cannot go to the next step until they are filled in.', 'The colored bar on the total at the right shows the split between meeting rooms / food / guest rooms.']
    },
    {
      id: 'sendquote', group: 'sell', icon: 'send', title: 'Printing and sending a quotation', who: 'Sales · Sales Manager',
      intro: 'On the quotation page, tap “Send / Print”. You see a preview of the real document on the left and the sending panel on the right.',
      steps: [
        'Check the document on the left (company header, customer, items, prices, terms).',
        'Choose the quotation format in the panel on the right: “Customer copy” (hides the separate service charge and VAT lines and shows the net total) or “Full quotation” (shows every line) — it is printed and sent in the format you choose.',
        'Tap “Edit details” (sales): the text on the quotation turns into dashed boxes you can type in, such as the customer name, address, Tax ID, item names, details, quantities, terms and the signatory’s name — the ↺ button = Use the system value. Then tap “Save changes” (it applies to that quotation only).',
        'Meeting room section on the quotation: shows the room photo and price. In edit mode you can type your own room details (tap “+ Insert the room’s details from the system” to bring in the existing text), and tap the photo to upload a new one, ✕ to remove the photo, or ↺ to use the room’s photo.',
        'Tap “Print / PDF” to print or save as PDF.',
        'Send by email: enter the recipient’s email, edit the message, and choose “Let the customer confirm and sign online” and “Show unit prices”.',
        'Set “Price valid until”, then tap “Send quotation”.',
        'The sending history is at the bottom of the panel.'
      ],
      shots: [{ img: 'quote-doc', cap: 'Customer copy and the email panel' }, { img: 'quote-doc-full', cap: 'Full quotation (shows separate service charge and VAT lines)' }, { img: 'quote-doc-edit', cap: 'Editing details on the quotation' }],
      tips: ['Documents sent to customers are always in Thai, even if you use the app in another language.', 'Amounts and totals cannot be edited on the quotation — change the items on the quotation page.', 'The terms at the end include the line “Credit card payments carry an extra 3% fee on the quotation amount” · If you edited the terms field on a quotation yourself, tap ↺ to get the latest text.']
    },
    {
      id: 'beo', group: 'event', icon: 'list', title: 'Function Sheets (BEO)', who: 'Sales · Managers · Heads of Department · Accounting (view)',
      intro: 'A function sheet is the work order for every department. It is created from a confirmed quotation, and each department acknowledges it in the system.',
      steps: [
        'The list on the left is split into Upcoming · All · Draft. Labels show the status, such as “All departments” or “2 to acknowledge”.',
        'Tap “+ Create from quote” and choose a quotation. The system fills in the rooms, food and head count automatically.',
        'The header shows the event date, confirmed guests (+5% spare), rooms, contact person and salesperson.',
        'Daily schedule: time, details, place and responsible department (switch between “Day 1 / Day 2”). You can import it from an Excel file.',
        'Each department card (Kitchen, Banquet, Engineering, Front Office, Housekeeping, Accounting) shows its tasks and an “Acknowledged / Awaiting acknowledgement” status. Tap “Edit” to add tasks for that department.',
        'Every change made after sending is counted as “Revision …” and goes to an administrator for approval before it is sent to departments.',
        'Buttons at the top: Function details · Sales summary · Floor plan (administrators, F&B and Banquet departments) · Room setup (administrators) · Print / PDF.'
      ],
      shots: [{ img: 'beo', cap: 'Function sheet: daily schedule and department cards' }, { img: 'beo-list', cap: 'Function sheet list with acknowledgement status' }],
      tips: ['Heads of department should tap “Acknowledge” every time there is a new sheet or a change.']
    },
    {
      id: 'plan', group: 'event', icon: 'grid', title: 'Floor plans and room setup', who: 'Administrator · F&B department · Banquet department (heads and staff)',
      intro: 'Draw the room layout from a function sheet and print it for the setup team.',
      steps: [
        'On the function sheet, tap “Floor plan” and choose the day (Day 1 / Day 2).',
        'Choose a template on the left (From the quotation · Theatre · Meeting · Round tables, etc.). The system places the tables and chairs automatically.',
        'Tap “+ Round table / + Long table / + Chair row / + Stage …” to add items, then drag them into place on the plan.',
        'Tap an item to select, move, rotate or delete it. Tap “Undo” if you make a mistake.',
        '“Room setup”: set the room size (meters) and fixed items that cannot be moved, such as the stage, doors, pillars and power outlets — this applies to every event in that room.'
      ],
      shots: [{ img: 'plan', cap: 'Floor plan: templates, adding items and the room plan' }, { img: 'plansetup', cap: 'Room setup: room size and fixed items' }],
      tips: ['Plans can be printed as PDF and attached to the function sheet.', 'The system checks the “Department” field in the user account: Kitchen / F&B / Food & beverage / Banquet — other people can see the plan thumbnail on the function sheet but cannot open the designer.', '“Room setup” is for administrators only.']
    },
    {
      id: 'fnforms', group: 'event', icon: 'print', title: 'Function details and sales summary', who: 'Administrator · Sales Manager · Sales Staff (edit) · Accounting · Heads of Department (view)',
      intro: 'Printable forms that match the hotel’s original paper forms. Data comes from the quotation and function sheet.',
      steps: [
        'Tap “Function details” — switch between “Print preview / Edit details” and enter the Function No. (numbered automatically, or type your own) — the Edit details tab, the “＋ Add details” button and numbering are only for administrators, sales managers and sales staff.',
        'The “Accounting / Finance” section (total, cost details, billing address) is only visible to administrators, sales managers, sales staff and accounting — other roles see only the Front Office section, at full width.',
        'Tap “Sales summary” — a summary of function revenue, food, snacks, guest rooms and tax invoice details.',
        'Check it, then tap “Print / PDF”.'
      ],
      shots: [{ img: 'fn', cap: 'Function details (print form)' }, { img: 'ss', cap: 'Function summary for sales' }],
      tips: ['Print forms are always in Thai.']
    },
    {
      id: 'cal', group: 'event', icon: 'cal', title: 'Booking calendar', who: 'Sales · Managers · Heads of Department · Accounting',
      intro: 'The Calendar menu shows all events by date, color-coded by status.',
      steps: [
        'Tap ‹ › to change the month, or “Today” to return to the current month.',
        'Switch between the Month · Week · Pipeline views.',
        'Filters on the left: status (Draft · Tentative · Confirmed) with counts and totals, plus filters by room/area.',
        'Tap an event name in the calendar to open its function sheet (or the quotation if there is no function sheet yet).'
      ],
      shots: [{ img: 'cal-month', cap: 'Monthly calendar' }, { img: 'cal-week', cap: 'Weekly view' }, { img: 'cal-pipe', cap: 'Pipeline view' }],
      tips: []
    },
    {
      id: 'shifts', group: 'people', icon: 'clock', title: 'Shifts and shift extensions (QT)', who: 'Heads of Department · Managers · Administrator',
      intro: 'Staff › Shifts is where you put people on shifts for each event. The system works out how many staff are needed from the number of guests.',
      steps: [
        'Choose the event at the top (tap the quotation name ▾) and choose the day.',
        'If there is no plan yet, tap “Create plan from the quotation” — or “Apply rules to this plan” to recalculate the numbers.',
        'Check the summary boxes: Staff needed · Assigned · Short-staffed · Extend shift.',
        'Department timelines: tap a shift (Morning shift/Afternoon shift) to set the number needed and the times, tick the staff, then tap “Save”.',
        'Extend shift (QT): choose +1/+2/+3 hrs or a custom time, enter a reason, choose the staff, then send it for approval.',
        'People with approval rights tap Approve / Send back on shift extension requests.'
      ],
      shots: [{ img: 'staff-shifts', cap: 'Shifts: summary, timeline and shift extensions' }, { img: 'staff-assign', cap: 'Assigning staff to a shift' }],
      tips: ['Hours worked this week are shown next to each name (e.g. 40 / 48 hrs) to help avoid overtime.']
    },
    {
      id: 'emps', group: 'people', icon: 'users', title: 'All staff, locations, departments and positions', who: 'People who can plan staff shifts (add/edit) · everyone who sees the menu (view)',
      intro: 'Staff › All staff holds the staff list for every location: Resort · Boutique · มณีอินทร์ · ช.ดรักส์ · ช.เฮ้าส์ · ภูมณีพันธุ์.',
      steps: [
        'Choose a location from the red pin bar (or “All locations”), then filter by department or search by name, position or phone number.',
        'Tap “+ Add staff”, fill in the full name, choose the location, department and position, enter the phone number and hours worked, then tap “Add staff”.',
        'Tap ✏️ to edit or 🗑 to delete (the system also removes the person from their shifts).',
        'The + button next to the Place / Department / Position fields: add a new one, rename it (all staff using it are updated) or delete it (only when no staff use it).',
        'Import from Excel/CSV: tap “Import staff list” › download the template › choose the file › check the table › choose “Add to the current list” or “Replace all at the locations in the file” › tap Import.',
        'In the import file, the location column is optional (left blank = the location currently selected), and departments not yet in the system are added as new departments.'
      ],
      shots: [{ img: 'staff-emps', cap: 'All staff by location and department' }, { img: 'staff-plus', cap: 'The + button for managing locations, departments and positions' }, { img: 'staff-import', cap: 'Importing a staff list from an Excel / CSV file' }],
      tips: ['The 6 main departments can be renamed but not deleted · Resort cannot be deleted.', 'Adding, renaming or deleting locations, departments and positions is saved as soon as you tap.']
    },
    {
      id: 'rules', group: 'people', icon: 'sliders', title: 'Head-count rules', who: 'People who can plan staff shifts',
      intro: 'Staff › Head-count rules sets how many staff are needed per shift for a given number of guests.',
      steps: [
        'Turn on “Apply rules” for the department you want (off = the system’s automatic calculation is used).',
        'Set guest ranges, for example 1–50 guests = 1 staff, 51–100 guests = 2 staff.',
        'Tap “+ Add guest range” or 🗑 to delete a range, then save.',
        'The rules are used with “Create plan from the quotation” or “Apply rules to this plan”.'
      ],
      shots: [{ img: 'staff-rules', cap: 'Head-count rules per shift' }],
      tips: []
    },
    {
      id: 'tasks', group: 'people', icon: 'check', title: 'Assignment & Approve (assigned tasks and approvals)', who: 'Everyone (receive tasks) · Administrators and roles given permission (assign tasks) · Administrator (Approve · ⚙ permissions)',
      intro: 'Assign tasks to users or whole departments, track their status, and approve edited function sheets.',
      steps: [
        'Assignment tab: see all tasks and filter by Pending / Confirmed — urgent tasks have an “Urgent” label and overdue tasks are red.',
        'Tap “+ Assign task”, enter the task name, details, assignee (a user or a department), due date and urgency. You can link it to a quotation.',
        'Assigning to a department: tap a location button (All locations · Resort · Boutique · มณีอินทร์ · ช.ดรักส์ · ช.เฮ้าส์ · ภูมณีพันธุ์) — the number on the button is how many staff that department has at that location — then tick the staff (none ticked = the whole department). The location is shown in the task list.',
        'The ⚙ button next to the heading (administrator): Assignment permissions, 2 ways — “Org chart” (a role higher up can assign work to every role and department below it; change “Reports to” to move one; − + Fit 100% buttons to zoom; and ✏️ rename, ＋ add a role/department, 🗑 delete on each box) or “Selection form” (choose, for each role, which roles and departments it can assign work to) — tap Save and the chosen way is used.',
        'Roles given permission see the “+ Assign task” button (only the people and departments they are allowed to pick) and a “Tasks I assigned” list, where they can edit/delete only their own tasks, and they are notified when the assignee confirms.',
        'The assignee taps the task on the Home page or in the bell to see the details, then taps “Confirm” with a note, such as “All set up”.',
        'Approve tab (administrator): function sheets awaiting approval. Tap “View” to compare the changes, then “Approve” or send it back.',
        'Approved sheets can be sent to departments and move to the “Approved” section.'
      ],
      shots: [{ img: 'tasks', cap: 'Assigned task list and the ⚙ permissions button' }, { img: 'task-new', cap: 'Assigning a new task' }, { img: 'task-site', cap: 'Assigning to a department: choosing the location and staff' }, { img: 'task-perms-org', cap: 'Assignment permissions: Org chart' }, { img: 'task-perms-form', cap: 'Assignment permissions: Selection form' }, { img: 'task-view', cap: 'Details of a received task' }, { img: 'tasks-approve', cap: 'Approve: function sheets awaiting approval' }],
      tips: ['Tasks assigned to a department can be seen and confirmed by everyone in that department.', 'Until permissions are set, only administrators can assign tasks · Administrators can always assign work to anyone.', 'In the chart, you can only delete roles or departments you added yourself that have nobody in them · Changes to the chart take effect when you tap Save.']
    },
    {
      id: 'leave', group: 'people', icon: 'leave', title: 'Leave Requests (QT · Change day off · Missed fingerprint scan)', who: 'Everyone (submit) · Administrator (sign/approve) · Human Resources (details)',
      intro: 'The Leave Requests menu replaces paper form HRSD08.1. There are 3 types: QT request · Change day off · Missed fingerprint scan.',
      steps: [
        'Tap “+ New” on the card for the type you need.',
        'Fill in your full name, nickname, position, department and branch (filled in from your account; you can edit them).',
        'QT request: choose Work QT or Take QT (use QT), the start–end time and date. The system works out the hours for you.',
        'Change day off: choose Change day off or Change working time, from date… to date….',
        'Missed fingerprint scan: choose the reason, the in/out time, the date and a witness.',
        'Enter a reason, then tap “Submit”.',
        'Track the status in the My requests · Awaiting approval · All tabs. Tap a request to see the signing steps (Head of Department · General Manager · Deputy General Manager · Human Resources).',
        'While nobody has signed a request yet, its owner can edit ✏️ or delete 🗑 it · Tap “Print form” to print the form.',
        'Human Resources/administrators: the “Details” tab lets you search, filter by month, type, status and department, see a summary per person (Work QT/Take QT/QT balance) and export to Excel.'
      ],
      shots: [{ img: 'leave', cap: 'Leave Requests page and My requests list' }, { img: 'leave-qt', cap: 'QT request form' }, { img: 'leave-change', cap: 'Change day off form' }, { img: 'leave-scan', cap: 'Missed fingerprint scan form' }, { img: 'leave-view', cap: 'Request details and signing steps' }, { img: 'leave-info', cap: 'Details for Human Resources' }],
      tips: ['QT must be at least 2 hours · QT may not be worked before the start of your shift.', 'Only administrators can sign and approve leave requests.']
    },
    {
      id: 'users', group: 'admin', icon: 'shield', title: 'Users, roles and permissions', who: 'Administrator · Sales Manager (with the Manage users permission)',
      intro: 'The Users menu is where you create accounts, set roles and turn individual permissions on or off.',
      steps: [
        'Search or filter by role: Sales Staff · Sales Manager · Head of Department · Accounting · Administrator.',
        'Tap “+ Add user”, enter the name, email (used to sign in), phone, role and department, then save. Tell the user their starting password and ask them to change it the first time they sign in.',
        'Tap a user’s name to open the panel on the right and turn each permission on or off, such as Create and send quotations · Discounts up to 10% · Approve discounts over 10% · View function sheets and the booking calendar · Manage customer data · Plan staff shifts · Manage users · Edit prices and packages · View leave request details (HR).',
        'Tap ✏️ to edit details, change the status (Active / Disabled), set up two-step verification or reset the password.',
        'Roles: tap “+ Add role”, enter a name, choose the visible menus and permissions, then save — saving a role updates the permissions of everyone in that role (you can also add/rename/delete roles from the org chart in Assignment ⚙).'
      ],
      shots: [{ img: 'users', cap: 'User list and permissions panel' }, { img: 'user-edit', cap: 'Editing user details' }, { img: 'role', cap: 'Individual permissions' }],
      tips: ['When you change someone’s role, their permissions change to match the new role (you can fine-tune them afterwards).', 'Quotations already sent do not change when permissions change.']
    },
    {
      id: 'dash', group: 'admin', icon: 'chart', title: 'Dashboard sales reports', who: 'Administrator · Management',
      intro: 'The Dashboard menu summarizes revenue by event date, by year and by month.',
      steps: [
        'Choose the year (‹ year ›) and a month, or “Whole year”.',
        'Summary boxes: confirmed revenue for the period · compared with last year · awaiting confirmation (Sent + Tentative) · % closed.',
        'Monthly revenue chart (Confirmed / awaiting confirmation): point at or tap a dot to see the figures, or tap “View as table”.',
        'You can export the data as a CSV file.'
      ],
      shots: [{ img: 'dash', cap: 'Dashboard of yearly and monthly revenue' }],
      tips: []
    },
    {
      id: 'backup', group: 'admin', icon: 'save', title: 'Backups and system updates', who: 'Administrator',
      intro: 'All data is stored in a Cloudflare D1 database and syncs to every device automatically.',
      steps: [
        'Back up: tap your profile picture › “Back up data (download)” and keep the file somewhere safe.',
        'Restore: use “Import backup” and choose the file you downloaded (only when really needed).',
        'System updates: when new features are released, the app on every device updates itself within a few minutes.'
      ],
      shots: [],
      tips: ['Never send passwords or API tokens by chat or email.']
    }
  ],
  quick: [
    {
      id: 'q-sales', title: 'Sales (sales staff)', color: '#0E9F8E', menus: 'Home · Rooms & Venues · Catering · Quotations · Function Sheets · Calendar · Leave Requests · Assignment',
      can: ['Create and send quotations', 'Discounts up to 10%', 'View function sheets and the booking calendar', 'Manage customer data'],
      steps: [
        'Tap “+ New quotation” › complete the 5 steps: Customer (you can enter the Tax ID) › Meeting rooms › Food › Guest rooms › Summary.',
        'Open the quotation › adjust items, discount and terms › tap Send / Print › choose Customer copy / Full quotation › you can edit details on the quotation › email it or print a PDF.',
        'When the customer confirms › change the status to “Confirmed” › Function Sheets › “+ Create from quote”.',
        'Fill in the daily schedule and each department’s tasks › follow up on acknowledgements.',
        'Check the calendar before taking a new booking to make sure rooms are free.'
      ],
      shots: ['wiz1', 'quote-edit', 'beo'], links: ['newquote', 'quotes', 'sendquote', 'beo', 'cal']
    },
    {
      id: 'q-manager', title: 'Sales Manager', color: '#A8262D', menus: 'All Sales menus + Staff · Users',
      can: ['Everything Sales can do', 'Approve discounts over 10%', 'Plan staff shifts / approve shift extensions', 'Manage users'],
      steps: [
        'Review quotations with discounts over 10% and approve them.',
        'Check the Home page: monthly target and latest quotations.',
        'Assign work to the team: Assignment › “+ Assign task” (once an administrator gives permission in ⚙).',
        'Approve shift extension (QT) requests on the Shifts page.',
        'Add/edit users in the sales team and set their permissions.'
      ],
      shots: ['quote-edit', 'tasks', 'users'], links: ['quotes', 'tasks', 'shifts', 'users']
    },
    {
      id: 'q-ops', title: 'Operations heads of department (Kitchen/F&B · Banquet · Engineering/AV · Front Office · Housekeeping)', color: '#D4A21A', menus: 'Home · Function Sheets · Calendar · Staff · Leave Requests · Assignment',
      can: ['View function sheets and the booking calendar', 'Plan staff shifts / approve shift extensions'],
      steps: [
        'Check the bell every morning: new or edited function sheets › open the sheet › tap “Acknowledge” on your department’s card.',
        'Read the daily schedule and floor plan to prepare equipment and staff (F&B and Banquet: you can design floor plans).',
        'Staff › Shifts › choose the event › assign staff to shifts › request shift extensions when needed.',
        'Keep your department’s staff list up to date (add/edit/import).',
        'Tasks assigned to your department › open them › tap “Confirm” when everything is ready.'
      ],
      shots: ['beo', 'staff-shifts', 'task-view'], links: ['beo', 'shifts', 'emps', 'tasks']
    },
    {
      id: 'q-acc', title: 'Accounting', color: '#2B3F86', menus: 'Home · Quotations (view) · Function Sheets · Calendar · Leave Requests',
      can: ['View function sheets and the booking calendar', 'Manage customer data'],
      steps: [
        'Review confirmed quotations: net total, deposit terms and customer details.',
        'Function Sheets › “Sales summary”: check revenue, food, guest rooms and tax invoice details › Print / PDF.',
        'Function details › “Accounting / Finance” section: see the total and cost details.',
        'Acknowledge the “Accounting” card on the function sheet.',
        'Calendar › Pipeline view to see amounts awaiting confirmation.'
      ],
      shots: ['ss', 'quotes', 'cal-pipe'], links: ['quotes', 'fnforms', 'cal']
    },
    {
      id: 'q-hr', title: 'Human Resources (HR)', color: '#7B4FA0', menus: 'Leave Requests (Details) · Staff',
      can: ['View leave request details (HR)', 'View the staff list'],
      steps: [
        'Leave Requests › “Details” tab.',
        'Filter by month, type, status or department, or search by name/nickname/number.',
        'See the summary per person: Work QT · Take QT · QT balance · Change day off · Missed fingerprint scan · Awaiting approval.',
        'Tap “Export to Excel” for payroll/reports.',
        'Check each location’s staff list in the Staff menu.'
      ],
      shots: ['leave-info', 'leave', 'staff-emps'], links: ['leave', 'emps']
    },
    {
      id: 'q-admin', title: 'Administrator', color: '#17211F', menus: 'All menus',
      can: ['All permissions', 'Sign and approve leave requests', 'Approve edited function sheets and overlapping room bookings', 'Edit guest room rates and menus', 'Manage users, roles and permissions'],
      steps: [
        'Assignment & Approve › Approve tab: approve edited function sheets · ⚙ Assignment permissions (org chart / form).',
        'Leave Requests › Awaiting approval tab › open the request › Sign / Reject.',
        'Users: create accounts, set roles and permissions, enable/disable users.',
        'Set guest room rates, edit meeting rooms (✏️ photos, prices, capacity) and update catering menus when prices change.',
        'Dashboard: track sales · back up data regularly.'
      ],
      shots: ['tasks-approve', 'task-perms-org', 'leave-view', 'dash'], links: ['tasks', 'leave', 'users', 'meet', 'rooms', 'food', 'dash', 'backup']
    },
    {
      id: 'q-all', title: 'All staff', color: '#6B716E', menus: 'Home · Leave Requests · Assigned tasks',
      can: ['Submit a QT request · Change day off · Missed fingerprint scan', 'Receive and confirm assigned tasks', 'Change your own language and color theme'],
      steps: [
        'Sign in › change the language with the TH/EN/MY button.',
        'Leave Requests › choose the type › “+ New” › fill in the form › “Submit”.',
        'Check the status in the “My requests” tab. You can edit or delete a request until someone signs it.',
        'Assigned tasks: tap one on the Home page or in the bell › read it › tap “Confirm” when done.'
      ],
      shots: ['leave-qt', 'task-view', 'lang'], links: ['login', 'langtheme', 'leave', 'tasks']
    }
  ],
  changelog: [
    { v: '1.1', date: 'Oct 22, 2026', text: 'Meeting rooms: ✏️ button to edit photos, prices and capacity, and swipe through sample photos · Quotations: Net prices with the service charge/VAT worked back out, Tax ID field, Customer copy / Full quotation, editing details on the quotation, changing the meeting room photo, and the 3% credit card term · Function sheets: permissions for editing function details, the Accounting/Finance section and floor plans (F&B · Banquet) · Assignment: choose a location when assigning to a department, and ⚙ Assignment permissions (an org chart with zoom and rename/add/delete, or a form).' },
    { v: '1.0', date: 'Oct 21, 2026', text: 'First edition, covering every chapter: quotations, meeting rooms, guest rooms, catering, function sheets, floor plans, calendar, staff (locations, departments, positions, import), shifts, assigned tasks, leave requests, users, Dashboard, 3 languages and color themes.' }
  ]
};
