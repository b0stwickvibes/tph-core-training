function getBartenderDay1Data() {
  return {
    role: 'Bartender',
    day: 1,
    sheetName: 'Bartender Day 1',
    title: 'Orientation & Bar Foundations',
    locations: [
      'OAK: Orientation (3pm-8pm)',
      'Cantina GNV: Orientation (4pm-8pm)',
      'White Buffalo: Orientation (7pm-10pm)'
    ],
    sections: [
      { number: 1, title: 'Welcome & Introduction', items: [
        'Meet the Manager on Duty; introduce trainee to leadership on shift',
        'Complete any remaining new hire paperwork',
        'Clock in; walk trainee through the time clock process',
        '5-Day Program Overview; explain the training structure, what each day covers, and what\'s expected by Day 5',
        'Set the tone: "This is your career development, not just a job checklist. Own your training."'
      ]},
      { number: 2, title: 'Trainual Accountability Check', items: [
        'Confirm trainee has completed all pre-training Trainual modules (Onboarding Roadmap)',
        'Review: Welcome to Trainual, Welcome Packet, Employee Handbook, Role-Specific Manual',
        'If incomplete: trainee must finish before Day 2. Flag to management immediately.'
      ]},
      { number: 3, title: 'Company Policies & Procedures', items: [
        'Mission, Vision, and Core Values; what TPH stands for and how it shows up on the floor',
        'Attendance policy; call-out procedures, no-shows, tardiness expectations',
        'Uniform policy; what\'s required, what\'s not acceptable, where to get replacements',
        'Code of conduct; professionalism, phone policy, guest-first mentality'
      ]},
      { number: 4, title: 'Uniforms, Hygiene & Safety', items: [
        'Appearance standards walkthrough; hair, nails, jewelry, tattoos, fragrance',
        'Alcohol agreement review and signature; responsible service expectations',
        'Health and hygiene requirements; handwashing, illness reporting, food safety basics',
        'Emergency procedures; fire exits, first aid kit location, emergency contacts'
      ]},
      { number: 5, title: 'Payroll, Tips & Scheduling', items: [
        '7Shifts; download app, log in, show how to view schedule, request time off, swap shifts',
        'Branch; set up account for tip payouts, explain payout cadence',
        'Tip structure; how tips are calculated, pooled vs. individual, tip-out percentages per role',
        'Pay periods, direct deposit confirmation, who to contact for payroll issues'
      ]},
      { number: 6, title: 'The Grand Tour', items: [
        'Front of House; entrance, host stand, dining room layout, bar area, patio/outdoor seating, restrooms',
        'Back of House; kitchen line, dish pit, dry storage, walk-in coolers, office, staff area, lockers',
        'Emergency exits; point out every exit and the evacuation route',
        'Introduce trainee to every team member on shift by name and role'
      ]},
      { number: 7, title: 'Spirit Categories Overview', items: [
        'Walk through the 6 spirit categories: Vodka, Gin, Rum, Tequila/Mezcal, Whiskey, Brandy/Cognac',
        'Explain how each category shows up on the menu and in guest orders',
        'Point out where each category lives on the back bar and well',
        'Quick-hit: "If a guest asks for a recommendation, you need to know what\'s in front of you."'
      ]},
      { number: 8, title: 'Product Knowledge: Flavor Profiles & Brands', items: [
        'Walk the back bar; identify every bottle by name, category, and price tier (well, call, premium, top shelf)',
        'Cover flavor profiles: what makes a bourbon different from a rye, a reposado from a blanco',
        'Highlight house pours and most-ordered spirits',
        'TPH-specific: which brands are featured, any exclusive pours or partnerships'
      ]},
      { number: 9, title: 'Introduction to Bartending', items: [
        'Brief history of the craft; why bartending is a skill, not just pouring drinks',
        'Core duties overview: drink preparation, guest interaction, bar maintenance, cash handling, opening/closing',
        'You are the face of the bar. Every drink, every conversation, every handoff. That\'s your reputation.'
      ]},
      { number: 10, title: 'SCS & The Basics', items: [
        'SCS (Sequence of Customer Service); the step-by-step framework for every guest interaction',
        'Greet within 30 seconds, identify the occasion, suggest with confidence, deliver with care, check back, close strong',
        'This is the backbone. Every shift, every guest, every time. No shortcuts.'
      ]},
      { number: 11, title: 'Building Cocktails & Basic Mixing', items: [
        'Introduction to cocktail structure: base spirit, modifier, sweetener, bitters, garnish',
        'Demonstrate a basic build; trainer makes a cocktail while explaining each step',
        'Jigger pours vs. free pours; when to use which and why accuracy matters',
        'Shaken vs. stirred; when and why',
        'Trainee watches and asks questions. No hands-on yet (that\'s Day 2).'
      ]},
      { number: 12, title: 'Essential Bar Tools', items: [
        'Identify and explain every tool: jigger, shaker tin, mixing glass, bar spoon, muddler, strainer (Hawthorne, fine mesh, julep), channel knife, peeler, pour spouts',
        'Where each tool lives behind the bar. Everything has a home. Put it back.',
        'Proper handling and care; how to clean, when to replace'
      ]},
      { number: 13, title: 'Glassware Types', items: [
        'Walk through every glass type used at the venue: rocks, highball, coupe, martini, wine (red/white), pint, snifter, shot, flute',
        'Which drinks go in which glass and why it matters (presentation, temperature, volume)',
        'Where glassware is stored, how to handle (never grab by the rim), and breakage protocol'
      ]},
      { number: 14, title: 'Bar Layout', items: [
        'Full walkthrough of the bar setup: well, speed rail, back bar, ice wells, garnish caddy, dump sink, glass racks, POS terminals',
        'This is your workspace. You need to know where everything is without looking.',
        'Point out the flow: how drinks move from order to build to garnish to serve',
        'Identify any location-specific differences (OAK vs. Cantina GNV vs. White Buffalo)'
      ]},
      { number: 15, title: 'Efficient Bar Setup', items: [
        'Walk through the opening bar setup process; what needs to happen before doors open',
        'Mise en place: garnish prep, juice prep, ice, stocking, napkins, straws, check-presenters',
        'Par levels; what needs to be full, what gets restocked during shift, what waits for closing',
        'A clean, stocked bar before service starts is non-negotiable.'
      ]},
      { number: 16, title: 'Shadow & Observe', items: [
        'Walk trainee through the bar during live service (15-20 minutes toward end of shift)',
        'Point out real-time examples: how bartenders greet guests, build drinks, handle multiple tickets',
        'Trainee watches the flow; ticket comes in, drink gets built, garnished, served, bar gets wiped',
        'This is what Day 2 looks like. Tomorrow you\'ll be next to your trainer doing this.'
      ]},
      { number: 17, title: 'Classic Cocktail & Shot Familiarization', items: [
        'Review the venue\'s cocktail menu; go drink by drink, explain what\'s in each one',
        'Highlight the top 5 most-ordered cocktails and the top 5 most-ordered shots',
        'Trainer tells the trainee: "By Day 5, you need to make all of these from memory."',
        'If time permits: flip through the recipe book or Trainual recipe section together'
      ]},
      { number: 18, title: 'END OF SHIFT: Recap & Close-Out', items: [
        'What did we cover today? Trainee summarizes the day in their own words.',
        'What stands out? Trainee names 2-3 things that clicked or surprised them.',
        'What questions do you have? Address anything unclear before Day 2.',
        'Day 2 Preview: "Tomorrow you\'re on the floor with me during live service. Come ready to move."',
        'Trainual Assignment: Remind trainee to complete Day 1 Training Session & Knowledge Check before Day 2',
        'Accountability Form: Trainer completes the 3 questions + uploads photo of this signed checklist'
      ]}
    ]
  };
}