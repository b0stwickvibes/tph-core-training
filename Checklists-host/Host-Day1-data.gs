function getHostDay1Data() {
  return {
    role: 'Host',
    day: 1,
    sheetName: 'Host Day 1',
    title: 'Orientation & Host Foundations',
    locations: [
      'OAK: Orientation (4pm-8pm)',
      'Cantina GNV: Orientation (4pm-8pm)'
    ],
    sections: [
      { number: 1, title: 'Welcome & Introduction', items: [
        'Meet the Manager on Duty; introduce trainee to leadership on shift',
        'Complete any remaining new hire paperwork',
        'Clock in: walk trainee through the time clock process',
        '5-Day Program Overview: explain the training structure, what each day covers, and what\'s expected by Day 5',
        'Set the tone: "This is your career development, not just a job checklist. Own your training."'
      ]},
      { number: 2, title: 'Trainual Accountability Check', items: [
        'Confirm trainee has completed all pre-training Trainual modules (Onboarding Roadmap)',
        'Review: Welcome to Trainual, Welcome Packet, Employee Handbook, Role-Specific Manual',
        'If incomplete, trainee must finish before Day 2. Flag to management immediately.'
      ]},
      { number: 3, title: 'Company Policies & Procedures', items: [
        'Mission, Vision, and Core Values: what TPH stands for and how it shows up on the floor',
        'Attendance policy: call-out procedures, no-shows, tardiness expectations',
        'Uniform policy: what\'s required, what\'s not acceptable, where to get replacements',
        'Code of conduct: professionalism, phone policy, guest-first mentality'
      ]},
      { number: 4, title: 'Uniforms, Hygiene & Safety', items: [
        'Appearance standards walkthrough: hair, nails, jewelry, tattoos, fragrance',
        'Alcohol agreement review and signature: responsible service expectations',
        'Health and hygiene requirements: handwashing, illness reporting, food safety basics',
        'Emergency procedures: fire exits, first aid kit location, emergency contacts'
      ]},
      { number: 5, title: 'Payroll, Tips & Scheduling', items: [
        '7Shifts: download app, log in, show how to view schedule, request time off, swap shifts',
        'Branch: set up account for tip payouts, explain payout cadence',
        'Tip structure: how tips are calculated, pooled vs. individual, tip-out percentages per role',
        'Pay periods, direct deposit confirmation, who to contact for payroll issues'
      ]},
      { number: 6, title: 'The Grand Tour: FOH & BOH', items: [
        'Front of House: entrance, host stand, dining room layout, bar area, patio/outdoor seating, restrooms',
        'Back of House: kitchen line, dish pit, dry storage, walk-in coolers, office, staff area, lockers',
        'Emergency exits: point out every exit and the evacuation route',
        'Introduce trainee to every team member on shift by name and role'
      ]},
      { number: 7, title: 'Customer Interaction & Experience', items: [
        'First impressions: you are the first and last face every guest sees. That matters more than any other role.',
        'Greeting standards: warm, immediate, eye contact, smile. "Welcome to [venue name]" every single time.',
        'Waitlist communication: how to quote accurate wait times, what to say when it\'s longer than expected',
        'Phone scripts: how to answer (within 3 rings), standard greeting, how to handle reservations, to-go orders, and general inquiries',
        'Handling frustrated guests at the door: de-escalation, empathy, offering alternatives'
      ]},
      { number: 8, title: 'Host Operational Procedures', items: [
        'Reservation management: how reservations come in (Resy, phone, walk-in), how to confirm, modify, and cancel',
        'To-go orders: how to take them, where they\'re staged, how to hand off to the guest',
        'Phone payments: processing payments over the phone for to-go orders',
        'Large party coordination: when to flag management, how to prep the floor for big tops',
        'Waitlist management: adding guests, quoting times, notifying when table is ready'
      ]},
      { number: 9, title: 'Event & Reservation Management', items: [
        'Large party procedures: reservations for 6+, what information to collect, confirmation calls',
        'Event inquiries: what to say, who to direct them to, how to log the request',
        'Reservation confirmations: when to call, what to confirm (time, party size, special requests)',
        'No-show protocol: how long to hold a table, when to release, how to document',
        'VIP and regular guest recognition: how to flag in the system, how to communicate to the floor'
      ]},
      { number: 10, title: 'Administrative & General Info', items: [
        'Gift cards: how to sell, activate, and check balances',
        'Hours of operation: by day, by location, holiday hours',
        'Applications / hiring inquiries: "We\'d love for you to apply online at [URL]." Never take a paper resume.',
        'Lost and found procedures: where to store, how to log, how to follow up',
        'Common guest questions: parking, private events, dress code, kids menu, accessibility'
      ]},
      { number: 11, title: 'Opening Procedures', items: [
        'Jolt app walkthrough: daily opening checklist for the host stand',
        'Host stand setup: menus stocked, reservation book/iPad ready, pens, waitlist pad, to-go supplies',
        'Floor check: walk every table to confirm it\'s set and ready before doors open',
        'Communication with kitchen and bar: confirm they\'re ready for service',
        '"The host stand should be guest-ready 15 minutes before we open. No exceptions."'
      ]},
      { number: 12, title: 'Table Management', items: [
        'Floor plan walkthrough: every table number, every section, capacity per table',
        'Section assignments: how sections are divided, who has what tonight, where the rotation stands',
        'Table status awareness: open, seated, entrées fired, dessert, check dropped, bussing needed',
        '"You control the flow of the entire restaurant. If you lose track of the floor, everyone feels it."'
      ]},
      { number: 13, title: 'Guest Interaction: Observe Trainer', items: [
        'Trainer demonstrates a full guest greeting at the host stand. Walk-in and reservation.',
        'Walk through the seating process verbally: greet → check reservation → assess floor → seat → hand off to server',
        'Explain the difference between seating a deuce vs. a 6-top vs. a walk-in during a wait',
        'Trainee observes. No live guest interaction yet (that\'s Day 2)'
      ]},
      { number: 14, title: 'To-Go Orders', items: [
        'How to input a to-go order into the POS',
        'Remote payment processing: taking card info over the phone',
        'Packaging standards: how food gets bagged, what utensils/napkins/sauces go with it',
        'Handoff to guest: where to stage, how to confirm the order is correct before they leave',
        '"To-go is revenue. Treat every to-go guest like they\'re sitting at your best table."'
      ]},
      { number: 15, title: 'Server Rotation & Seating Process', items: [
        'How rotation works: who\'s next, how to track it, why even rotation matters for server tips and guest experience',
        'Double-seating: what it is, why it kills a server\'s timing, how to avoid it',
        'Rotation exceptions: large party requests, guest preferences, server in the weeds',
        'Communication with servers: "I just sat you" verbal confirmation or system notification'
      ]},
      { number: 16, title: 'Reservations: Resy App', items: [
        'Log in to Resy. Trainee gets access, navigates the interface.',
        'How to view tonight\'s reservations, upcoming reservations, and waitlist',
        'How to add a reservation: party size, time, name, phone, special notes',
        'How to modify or cancel, and when to call the guest vs. just update the system',
        'Table assignments in Resy: linking reservations to specific tables'
      ]},
      { number: 17, title: 'Familiarization: Observe the Flow', items: [
        'Walk the host stand and entrance area during live service (15-20 minutes toward end of shift)',
        'Point out real-time examples: how guests arrive, how the current host greets, how the waitlist moves',
        'Watch the seating flow: guest arrives → greeted → checked in → waited or seated → server handoff',
        'Trainee watches the rhythm. "This is what your Day 2 looks like. Tomorrow you\'re running this stand with me."'
      ]},
      { number: 18, title: 'END OF SHIFT: Recap & Close-Out', items: [
        'What did we cover today? Trainee summarizes the day in their own words.',
        'What stands out? Trainee names 2-3 things that clicked or surprised them.',
        'What questions do you have? Address anything unclear before Day 2.',
        'Day 2 Preview: "Tomorrow you\'re on the host stand with me during live service. Come ready to move."',
        'Trainual Assignment: Remind trainee to complete Day 1 Training Session & Knowledge Check before Day 2',
        'Accountability Form: Trainer completes the 3 questions + uploads photo of this signed checklist'
      ]}
    ]
  };
}
