// @ts-nocheck
/**
 * Mock Data Service
 * Provides fallback mock data for widgets when SharePoint lists don't exist yet
 * This allows the solution to demonstrate functionality before lists are provisioned
 */

// Helper function for padStart (ES2015 compatible)
function padStart(str: string, targetLength: number, padString: string): string {
  str = String(str);
  if (str.length >= targetLength) {
    return str;
  }
  const pad = padString.repeat(Math.ceil((targetLength - str.length) / padString.length));
  return pad.substring(0, targetLength - str.length) + str;
}

export class MockDataService {
  /**
   * Get mock surveys for MySurveysWidget
   */
  public static getMockSurveys(): any[] {
    return [
      {
        Id: 1,
        Title: 'Employee Engagement Survey Q4 2024',
        Description: 'Help us understand your experience and engagement',
        DueDate: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString(),
        Status: 'Active'
      },
      {
        Id: 2,
        Title: 'IT Systems Feedback Survey',
        Description: 'Share your feedback on IT tools and systems',
        DueDate: new Date(Date.now() + 14 * 24 * 60 * 60 * 1000).toISOString(),
        Status: 'Pending'
      },
      {
        Id: 3,
        Title: 'Onboarding Experience Survey',
        Description: 'Tell us about your onboarding journey',
        DueDate: new Date(Date.now() + 21 * 24 * 60 * 60 * 1000).toISOString(),
        Status: 'Active'
      }
    ];
  }

  /**
   * Get mock candidates for recruitment widgets
   */
  public static getMockCandidates(): any[] {
    const sources = ['LinkedIn', 'Indeed', 'Referral', 'Company Website', 'Recruiter'];
    const stages = ['Applied', 'Screening', 'Interview', 'Offer', 'Hired'];
    const statuses = ['Active', 'Rejected', 'Withdrawn', 'Archived'];
    const positions = ['Software Engineer', 'Product Manager', 'Data Analyst', 'UX Designer', 'Sales Manager'];

    const candidates = [];
    for (let i = 1; i <= 50; i++) {
      const status = statuses[Math.floor(Math.random() * statuses.length)];
      candidates.push({
        Id: i,
        Title: `Candidate ${i}`,
        CandidateName: `Candidate ${i}`,
        Email: `candidate${i}@example.com`,
        Phone: `(555) 000-${padStart(String(i), 4, '0')}`,
        CandidateStage: stages[Math.floor(Math.random() * stages.length)],
        CandidateStatus: status,
        JobRequisitionId: Math.floor(Math.random() * 20) + 1, // Random requisition 1-20
        Source: sources[Math.floor(Math.random() * sources.length)],
        Position: positions[Math.floor(Math.random() * positions.length)],
        Created: new Date(Date.now() - Math.random() * 90 * 24 * 60 * 60 * 1000).toISOString(),
        HiredDate: Math.random() > 0.8 ? new Date().toISOString() : null,
        Status: status // Keep both for compatibility
      });
    }
    return candidates;
  }

  /**
   * Get mock job requisitions
   */
  public static getMockRequisitions(): any[] {
    const departments = ['Engineering', 'Product', 'Sales', 'Marketing', 'HR', 'Finance'];
    const statuses = ['Open', 'Active', 'Filled', 'Closed'];
    const locations = ['New York', 'San Francisco', 'Remote', 'London', 'Austin'];
    const jobTitles = ['Senior Software Engineer', 'Product Manager', 'Data Analyst', 'UX Designer', 'Sales Manager', 'HR Specialist'];

    const requisitions = [];
    for (let i = 1; i <= 20; i++) {
      const isOpen = Math.random() > 0.5;
      const status = isOpen ? (Math.random() > 0.5 ? 'Open' : 'Active') : (Math.random() > 0.5 ? 'Filled' : 'Closed');
      requisitions.push({
        Id: i,
        Title: `REQ-2024-${padStart(String(i), 3, '0')}`,
        Position: `Position ${i}`,
        JobTitle: jobTitles[Math.floor(Math.random() * jobTitles.length)],
        Department: departments[Math.floor(Math.random() * departments.length)],
        Location: locations[Math.floor(Math.random() * locations.length)],
        Status: status,
        RequisitionStatus: status,
        Priority: Math.random() > 0.7 ? 'High' : (Math.random() > 0.5 ? 'Medium' : 'Low'),
        Created: new Date(Date.now() - Math.random() * 120 * 24 * 60 * 60 * 1000).toISOString(),
        FilledDate: !isOpen ? new Date(Date.now() - Math.random() * 30 * 24 * 60 * 60 * 1000).toISOString() : null,
        TargetHireDate: new Date(Date.now() + Math.random() * 60 * 24 * 60 * 60 * 1000).toISOString()
      });
    }
    return requisitions;
  }

  /**
   * Get mock job offers
   */
  public static getMockJobOffers(): any[] {
    const statuses = ['Pending Approval', 'Sent', 'Accepted', 'Declined'];

    const offers = [];
    for (let i = 1; i <= 15; i++) {
      const status = statuses[Math.floor(Math.random() * statuses.length)];
      const sentDate = new Date(Date.now() - Math.random() * 30 * 24 * 60 * 60 * 1000);

      offers.push({
        Id: i,
        Title: `Offer-${padStart(String(i), 3, '0')}`,
        CandidateName: `Candidate ${i}`,
        Position: `Position ${i}`,
        Status: status,
        Salary: 70000 + Math.floor(Math.random() * 80000),
        SentDate: status !== 'Pending Approval' ? sentDate.toISOString() : null,
        ResponseDate: ['Accepted', 'Declined'].includes(status) ?
          new Date(sentDate.getTime() + Math.random() * 14 * 24 * 60 * 60 * 1000).toISOString() : null,
        StartDate: status === 'Accepted' ?
          new Date(Date.now() + Math.random() * 60 * 24 * 60 * 60 * 1000).toISOString() : null
      });
    }
    return offers;
  }

  /**
   * Get mock interviews
   */
  public static getMockInterviews(): any[] {
    const types = ['Phone', 'Video', 'In-Person', 'Technical', 'Behavioral'];
    const statuses = ['Scheduled', 'Completed', 'Cancelled'];
    const times = ['09:00 AM', '10:00 AM', '11:00 AM', '01:00 PM', '02:00 PM', '03:00 PM', '04:00 PM'];

    const interviews = [];
    for (let i = 1; i <= 25; i++) {
      const status = statuses[Math.floor(Math.random() * statuses.length)];
      const interviewDate = status === 'Scheduled' ?
        new Date(Date.now() + Math.random() * 30 * 24 * 60 * 60 * 1000) :
        new Date(Date.now() - Math.random() * 30 * 24 * 60 * 60 * 1000);

      interviews.push({
        Id: i,
        Title: `Interview with Candidate ${i}`,
        CandidateName: `Candidate ${i}`,
        InterviewType: types[Math.floor(Math.random() * types.length)],
        InterviewDate: interviewDate.toISOString(),
        InterviewTime: times[Math.floor(Math.random() * times.length)],
        Duration: 30 + Math.floor(Math.random() * 4) * 15, // 30, 45, 60, or 75 minutes
        Interviewer: `Interviewer ${Math.floor(Math.random() * 10) + 1}`,
        InterviewerName: `Interviewer ${Math.floor(Math.random() * 10) + 1}`,
        Location: Math.random() > 0.5 ? 'Video Call' : `Room ${Math.floor(Math.random() * 10) + 1}`,
        Status: status,
        FeedbackSubmitted: status === 'Completed' ? Math.random() > 0.5 : false,
        Position: `Position ${Math.floor(Math.random() * 10) + 1}`
      });
    }
    return interviews;
  }

  /**
   * Get mock IT assets
   */
  public static getMockAssets(): any[] {
    const types = ['Laptop', 'Desktop', 'Monitor', 'Phone', 'Tablet', 'Keyboard', 'Mouse'];
    const statuses = ['Available', 'Assigned', 'In Repair', 'Retired'];

    const assets = [];
    for (let i = 1; i <= 100; i++) {
      const type = types[Math.floor(Math.random() * types.length)];
      const status = statuses[Math.floor(Math.random() * statuses.length)];

      assets.push({
        Id: i,
        Title: `${type}-${padStart(String(i), 4, '0')}`,
        AssetType: type,
        AssetTag: `AST-${padStart(String(i), 5, '0')}`,
        SerialNumber: `SN-${Math.random().toString(36).substring(2, 11).toUpperCase()}`,
        Status: status,
        AssetStatus: status, // Add for compatibility
        AssignedTo: status === 'Assigned' ? `Employee ${Math.floor(Math.random() * 50) + 1}` : null,
        PurchaseDate: new Date(Date.now() - Math.random() * 1095 * 24 * 60 * 60 * 1000).toISOString(),
        WarrantyExpiry: new Date(Date.now() + Math.random() * 365 * 24 * 60 * 60 * 1000).toISOString()
      });
    }
    return assets;
  }

  /**
   * Get mock software licenses
   */
  public static getMockSoftwareLicenses(): any[] {
    const software = [
      { name: 'Microsoft 365', cost: 12.50 },
      { name: 'Adobe Creative Cloud', cost: 52.99 },
      { name: 'Slack', cost: 8.00 },
      { name: 'Zoom', cost: 14.99 },
      { name: 'Salesforce', cost: 150.00 },
      { name: 'Jira', cost: 10.00 },
      { name: 'GitHub', cost: 4.00 },
      { name: 'Figma', cost: 15.00 }
    ];

    const licenses = [];
    for (let i = 0; i < software.length; i++) {
      const totalLicenses = 50 + Math.floor(Math.random() * 200);
      const assignedLicenses = Math.floor(totalLicenses * (0.6 + Math.random() * 0.3));

      licenses.push({
        Id: i + 1,
        Title: software[i].name,
        SoftwareName: software[i].name,
        LicenseType: software[i].name,
        TotalLicenses: totalLicenses,
        UsedLicenses: assignedLicenses,
        AssignedLicenses: assignedLicenses,
        AvailableLicenses: totalLicenses - assignedLicenses,
        CostPerLicense: software[i].cost,
        RenewalDate: new Date(Date.now() + Math.random() * 365 * 24 * 60 * 60 * 1000).toISOString(),
        ExpirationDate: new Date(Date.now() + Math.random() * 365 * 24 * 60 * 60 * 1000).toISOString(),
        Status: 'Active',
        Vendor: `${software[i].name.split(' ')[0]} Inc.`
      });
    }
    return licenses;
  }

  /**
   * Get mock task assignments for IT provisioning
   */
  public static getMockTaskAssignments(): any[] {
    const taskTypes = [
      'Setup Email Account',
      'Configure Laptop',
      'Install Software',
      'Create Network Access',
      'Setup Phone',
      'Configure VPN',
      'Grant System Access'
    ];
    const statuses = ['Pending', 'In Progress', 'Completed', 'Blocked'];
    const priorities = ['High', 'Medium', 'Low'];

    const tasks = [];
    for (let i = 1; i <= 50; i++) {
      const status = statuses[Math.floor(Math.random() * statuses.length)];
      const createdDate = new Date(Date.now() - Math.random() * 30 * 24 * 60 * 60 * 1000);

      tasks.push({
        Id: i,
        Title: taskTypes[Math.floor(Math.random() * taskTypes.length)],
        TaskType: taskTypes[Math.floor(Math.random() * taskTypes.length)],
        AssignedTo: `IT Staff ${Math.floor(Math.random() * 10) + 1}`,
        Employee: `Employee ${Math.floor(Math.random() * 50) + 1}`,
        Status: status,
        Priority: priorities[Math.floor(Math.random() * priorities.length)],
        DueDate: new Date(createdDate.getTime() + (5 + Math.random() * 10) * 24 * 60 * 60 * 1000).toISOString(),
        Created: createdDate.toISOString(),
        CompletedDate: status === 'Completed' ?
          new Date(createdDate.getTime() + Math.random() * 5 * 24 * 60 * 60 * 1000).toISOString() : null
      });
    }
    return tasks;
  }

  /**
   * Get mock CV/resume data
   */
  public static getMockCVData(): any[] {
    const statuses = ['New', 'Under Review', 'Shortlisted', 'Rejected', 'Interview Scheduled'];
    const candidateStages = ['New', 'Applied', 'Screening', 'Shortlisted', 'Interview'];
    const positions = ['Software Engineer', 'Product Manager', 'Data Analyst', 'UX Designer', 'Sales Manager'];

    const cvs = [];
    for (let i = 1; i <= 30; i++) {
      const candidateStage = candidateStages[Math.floor(Math.random() * candidateStages.length)];
      cvs.push({
        Id: i,
        Title: `Candidate ${i}`,
        CandidateName: `Candidate ${i}`,
        Email: `candidate${i}@example.com`,
        Phone: `(555) ${String(Math.floor(Math.random() * 900) + 100)}-${String(Math.floor(Math.random() * 9000) + 1000)}`,
        Position: `Position ${Math.floor(Math.random() * 10) + 1}`,
        PositionApplied: positions[Math.floor(Math.random() * positions.length)],
        Status: statuses[Math.floor(Math.random() * statuses.length)],
        CandidateStage: candidateStage,
        MatchScore: 60 + Math.floor(Math.random() * 40), // Score between 60-99
        YearsExperience: Math.floor(Math.random() * 15) + 1,
        Education: ['Bachelor', 'Master', 'PhD'][Math.floor(Math.random() * 3)],
        Skills: 'JavaScript, React, TypeScript, Node.js',
        Submitted: new Date(Date.now() - Math.random() * 60 * 24 * 60 * 60 * 1000).toISOString(),
        Created: new Date(Date.now() - Math.random() * 60 * 24 * 60 * 60 * 1000).toISOString(),
        ReviewedBy: Math.random() > 0.5 ? `Recruiter ${Math.floor(Math.random() * 5) + 1}` : null
      });
    }
    return cvs;
  }

  /**
   * Get mock timeline tasks
   */
  public static getMockTimelineTasks(): any[] {
    const processTypes = ['Onboarding', 'Offboarding', 'Probation', 'Training', 'Review', 'Project', 'Compliance', 'Other'];
    const taskStatuses = ['Not Started', 'In Progress', 'On Hold', 'Completed', 'Cancelled'];
    const milestoneTypes = ['Start', 'End', 'Review', 'Approval', 'Delivery'];

    const tasks = [];
    const today = new Date();

    for (let i = 1; i <= 50; i++) {
      const processType = processTypes[Math.floor(Math.random() * processTypes.length)];
      const startDate = new Date(today.getTime() + (Math.random() * 180 - 90) * 24 * 60 * 60 * 1000);
      const duration = 5 + Math.floor(Math.random() * 20); // 5-25 days
      const endDate = new Date(startDate.getTime() + duration * 24 * 60 * 60 * 1000);
      const isMilestone = Math.random() > 0.8;
      const isOnCriticalPath = Math.random() > 0.7;
      const status = taskStatuses[Math.floor(Math.random() * taskStatuses.length)];
      const progress = status === 'Completed' ? 100 : status === 'In Progress' ? Math.floor(Math.random() * 80) + 10 : 0;

      tasks.push({
        ID: i,
        Title: `${processType} Task ${i}`,
        TaskDescription: `Description for ${processType} task ${i}`,
        StartDate: startDate.toISOString(),
        EndDate: endDate.toISOString(),
        Progress: progress,
        TaskStatus: status,
        ProcessType: processType,
        ProcessInstanceID: Math.floor(Math.random() * 20) + 1,
        AssignedToList: `User ${Math.floor(Math.random() * 10) + 1};User ${Math.floor(Math.random() * 10) + 1}`,
        Owner: `Owner ${Math.floor(Math.random() * 5) + 1}`,
        DependsOn: i > 1 && Math.random() > 0.6 ? JSON.stringify([i - 1]) : null,
        BlockedBy: Math.random() > 0.9 ? JSON.stringify([Math.floor(Math.random() * i)]) : null,
        IsOnCriticalPath: isOnCriticalPath,
        ResourcesJSON: JSON.stringify([
          { name: `Resource ${Math.floor(Math.random() * 10) + 1}`, allocation: Math.floor(Math.random() * 100) }
        ]),
        EstimatedHours: duration * 8,
        ActualHours: status === 'Completed' ? duration * 8 * (0.8 + Math.random() * 0.4) : progress / 100 * duration * 8,
        IsMilestone: isMilestone,
        MilestoneType: isMilestone ? milestoneTypes[Math.floor(Math.random() * milestoneTypes.length)] : null,
        Color: '#0078d4',
        BackgroundColor: null,
        Created: new Date(Date.now() - Math.random() * 60 * 24 * 60 * 60 * 1000).toISOString(),
        Modified: new Date(Date.now() - Math.random() * 30 * 24 * 60 * 60 * 1000).toISOString(),
        Author: { Title: `Author ${Math.floor(Math.random() * 5) + 1}` },
        Editor: { Title: `Editor ${Math.floor(Math.random() * 5) + 1}` }
      });
    }
    return tasks;
  }

  /**
   * Get mock task dependencies
   */
  public static getMockTaskDependencies(): any[] {
    const dependencyTypes = ['FS', 'SS', 'FF', 'SF']; // Finish-to-Start, Start-to-Start, Finish-to-Finish, Start-to-Finish
    const dependencies = [];

    // Create some logical dependencies
    for (let i = 1; i <= 30; i++) {
      if (Math.random() > 0.5 && i > 1) {
        dependencies.push({
          ID: i,
          TaskID: i + 1,
          DependsOnTaskID: i,
          DependencyType: dependencyTypes[Math.floor(Math.random() * dependencyTypes.length)],
          LagDays: Math.random() > 0.8 ? Math.floor(Math.random() * 5) : 0
        });
      }
    }

    return dependencies;
  }

  /**
   * Get mock JML processes
   */
  public static getMockProcesses(): any[] {
    const processTypes = ['Onboarding', 'Offboarding', 'Transfer', 'Promotion', 'Leave', 'Other'];
    const processStatuses = ['Draft', 'In Progress', 'Under Review', 'Completed', 'Cancelled', 'On Hold'];
    const priorities = ['High', 'Medium', 'Low', 'Normal'];

    const processes = [];
    const today = new Date();

    for (let i = 1; i <= 30; i++) {
      const processType = processTypes[Math.floor(Math.random() * processTypes.length)];
      const processStatus = processStatuses[Math.floor(Math.random() * processStatuses.length)];
      const priority = priorities[Math.floor(Math.random() * priorities.length)];
      const startDate = new Date(today.getTime() - Math.random() * 90 * 24 * 60 * 60 * 1000);
      const completedDate = processStatus === 'Completed' ?
        new Date(startDate.getTime() + Math.random() * 60 * 24 * 60 * 60 * 1000) : null;

      processes.push({
        Id: i,
        Title: `${processType} Process ${i}`,
        ProcessType: processType,
        ProcessStatus: processStatus,
        Priority: priority,
        EmployeeName: `Employee ${Math.floor(Math.random() * 50) + 1}`,
        EmployeeEmail: `employee${i}@example.com`,
        Department: ['HR', 'IT', 'Finance', 'Sales', 'Marketing'][Math.floor(Math.random() * 5)],
        StartDate: startDate.toISOString(),
        TargetCompletionDate: new Date(startDate.getTime() + (30 + Math.random() * 60) * 24 * 60 * 60 * 1000).toISOString(),
        CompletedDate: completedDate ? completedDate.toISOString() : null,
        Progress: processStatus === 'Completed' ? 100 :
          processStatus === 'In Progress' ? Math.floor(Math.random() * 80) + 10 : 0,
        TaskCount: 5 + Math.floor(Math.random() * 10),
        CompletedTaskCount: processStatus === 'Completed' ? 15 : Math.floor(Math.random() * 10),
        Created: startDate.toISOString(),
        Modified: new Date(Date.now() - Math.random() * 30 * 24 * 60 * 60 * 1000).toISOString()
      });
    }
    return processes;
  }

  /**
   * Get mock calendar events
   */
  public static getMockCalendarEvents(userEmail?: string): any[] {
    const processTypes = ['Onboarding', 'Offboarding', 'Probation', 'Training', 'Review', 'Interview', 'Meeting', 'Deadline', 'Other'];
    const eventStatuses = ['Scheduled', 'In Progress', 'Completed', 'Cancelled', 'Postponed'];
    const locations = ['Conference Room A', 'Conference Room B', 'Virtual - Teams', 'Virtual - Zoom', 'Office', 'Remote'];

    const events = [];
    const today = new Date();
    const defaultEmail = userEmail || 'user@example.com';

    for (let i = 1; i <= 60; i++) {
      const processType = processTypes[Math.floor(Math.random() * processTypes.length)];
      const startDate = new Date(today.getTime() + (Math.random() * 90 - 30) * 24 * 60 * 60 * 1000);
      const isAllDay = Math.random() > 0.7;
      const duration = isAllDay ? 1 : Math.floor(Math.random() * 4) + 1; // 1-4 hours for regular events
      const endDate = new Date(startDate.getTime() + (isAllDay ? 24 * 60 * 60 * 1000 : duration * 60 * 60 * 1000));
      const status = eventStatuses[Math.floor(Math.random() * eventStatuses.length)];
      const isMilestone = Math.random() > 0.85;
      const isRecurring = Math.random() > 0.8;

      events.push({
        ID: i,
        Title: `${processType} - Event ${i}`,
        EventDescription: `This is a ${processType.toLowerCase()} event scheduled for ${startDate.toLocaleDateString()}`,
        StartDate: startDate.toISOString(),
        EndDate: endDate.toISOString(),
        IsAllDay: isAllDay,
        ProcessType: processType,
        ProcessInstanceID: Math.floor(Math.random() * 20) + 1,
        Location: locations[Math.floor(Math.random() * locations.length)],
        Attendees: `${defaultEmail};attendee${Math.floor(Math.random() * 10) + 1}@example.com`,
        Organizer: Math.random() > 0.5 ? defaultEmail : `organizer${Math.floor(Math.random() * 5) + 1}@example.com`,
        EventStatus: status,
        IsRecurring: isRecurring,
        RecurrenceRule: isRecurring ? 'FREQ=WEEKLY;BYDAY=MO,WE,FR' : null,
        Color: '#0078d4',
        BackgroundColor: null,
        IsMilestone: isMilestone,
        DependsOn: i > 1 && Math.random() > 0.7 ? JSON.stringify([i - 1]) : null,
        OutlookEventID: Math.random() > 0.6 ? `outlook-${i}-${Math.random().toString(36).substring(7)}` : null,
        IsSyncedWithOutlook: Math.random() > 0.5,
        Created: new Date(Date.now() - Math.random() * 60 * 24 * 60 * 60 * 1000).toISOString(),
        Modified: new Date(Date.now() - Math.random() * 30 * 24 * 60 * 60 * 1000).toISOString(),
        Author: { Title: `Author ${Math.floor(Math.random() * 5) + 1}` },
        Editor: { Title: `Editor ${Math.floor(Math.random() * 5) + 1}` }
      });
    }
    return events;
  }

  // ============================================
  // PROCUREMENT MANAGER MOCK DATA
  // ============================================

  /**
   * Get mock vendors for Procurement Manager
   */
  public static getMockVendors(): any[] {
    const vendorTypes = ['IT Services', 'Office Supplies', 'Professional Services', 'IT Security', 'Facilities', 'Cloud Services', 'Hardware', 'Legal Services', 'Travel Services', 'Data Services'];
    const statuses = ['Active', 'Active', 'Active', 'Pending Approval', 'Inactive'];
    const currencies = ['USD', 'USD', 'USD', 'EUR', 'GBP'];

    const vendors = [
      { name: 'Acme Technology Solutions', type: 'IT Services', city: 'San Francisco', state: 'CA', rating: 4.8, preferred: true },
      { name: 'Global Office Supplies Inc.', type: 'Office Supplies', city: 'Chicago', state: 'IL', rating: 4.5, preferred: true },
      { name: 'Premier Staffing Agency', type: 'Professional Services', city: 'New York', state: 'NY', rating: 4.2, preferred: false },
      { name: 'SecureNet Systems', type: 'IT Security', city: 'Austin', state: 'TX', rating: 4.9, preferred: true },
      { name: 'EcoClean Facilities Management', type: 'Facilities', city: 'Seattle', state: 'WA', rating: 4.3, preferred: false },
      { name: 'CloudFirst Solutions', type: 'Cloud Services', city: 'Denver', state: 'CO', rating: 4.7, preferred: true },
      { name: 'Precision Hardware Corp', type: 'Hardware', city: 'Phoenix', state: 'AZ', rating: 4.4, preferred: false },
      { name: 'Legal Eagles LLP', type: 'Legal Services', city: 'Boston', state: 'MA', rating: 4.6, preferred: true },
      { name: 'TravelWise Corporate', type: 'Travel Services', city: 'Miami', state: 'FL', rating: 4.1, preferred: false },
      { name: 'DataVault Storage Solutions', type: 'Data Services', city: 'Dallas', state: 'TX', rating: null, preferred: false }
    ];

    return vendors.map((v, i) => ({
      Id: i + 1,
      Title: v.name,
      VendorName: v.name,
      VendorType: v.type,
      Status: i < 9 ? 'Active' : 'Pending Approval',
      TaxId: `${10 + i}-${1000000 + i * 111111}`,
      RegistrationNumber: `REG-202${i % 5}-00${1000 + i}`,
      Email: `accounts@${v.name.toLowerCase().replace(/[^a-z]/g, '')}.com`,
      Phone: `+1 (555) ${100 + i * 111}-${4567 + i}`,
      Address: `${100 + i * 100} Business Park Drive`,
      City: v.city,
      State: v.state,
      PostalCode: `${90000 + i * 1000}`,
      Country: 'United States',
      PrimaryContactName: `Contact ${i + 1}`,
      PrimaryContactEmail: `contact${i + 1}@${v.name.toLowerCase().replace(/[^a-z]/g, '')}.com`,
      PrimaryContactPhone: `+1 (555) ${100 + i * 111}-${4570 + i}`,
      Currency: currencies[i % currencies.length],
      BankName: ['First National Bank', 'Chase Bank', 'Bank of America', 'Wells Fargo', 'US Bank'][i % 5],
      IsPreferred: v.preferred,
      Rating: v.rating,
      Notes: `Vendor ${i + 1} notes - ${v.type} provider.`,
      Created: new Date(Date.now() - (365 - i * 30) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - i * 7 * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock purchase requisitions
   */
  public static getMockPurchaseRequisitions(): any[] {
    const departments = ['Engineering', 'Operations', 'IT Security', 'Facilities', 'Marketing', 'IT Infrastructure', 'HR', 'Finance'];
    const statuses = ['Draft', 'Pending Approval', 'Approved', 'Rejected', 'Cancelled'];
    const priorities = ['High', 'Medium', 'Low'];

    const requisitions = [
      { title: 'Q4 Laptop Refresh - Engineering', dept: 'Engineering', amount: 45000, status: 'Approved', priority: 'High' },
      { title: 'Office Supplies - Monthly Restock', dept: 'Operations', amount: 2500, status: 'Approved', priority: 'Medium' },
      { title: 'Security Software Licenses', dept: 'IT Security', amount: 125000, status: 'Pending Approval', priority: 'High' },
      { title: 'Conference Room AV Equipment', dept: 'Facilities', amount: 35000, status: 'Draft', priority: 'Low' },
      { title: 'Marketing Campaign Materials', dept: 'Marketing', amount: 8500, status: 'Approved', priority: 'Medium' },
      { title: 'Cloud Infrastructure Expansion', dept: 'IT Infrastructure', amount: 75000, status: 'Under Review', priority: 'High' },
      { title: 'Training Materials - Q4', dept: 'HR', amount: 12000, status: 'Approved', priority: 'Medium' },
      { title: 'Financial Software Upgrade', dept: 'Finance', amount: 55000, status: 'Pending Approval', priority: 'High' }
    ];

    return requisitions.map((r, i) => ({
      Id: i + 1,
      Title: r.title,
      RequisitionNumber: `REQ-2024-${padStart(String(i + 1), 3, '0')}`,
      Requestor: `Requestor ${i + 1}`,
      Department: r.dept,
      RequestDate: new Date(Date.now() - (30 - i * 3) * 24 * 60 * 60 * 1000).toISOString(),
      RequiredDate: new Date(Date.now() + (30 + i * 7) * 24 * 60 * 60 * 1000).toISOString(),
      Status: r.status,
      Priority: r.priority,
      Justification: `Business justification for ${r.title}. This purchase is necessary to support ongoing operations and strategic initiatives.`,
      TotalAmount: r.amount,
      Currency: 'USD',
      CostCenter: `CC-${r.dept.substring(0, 3).toUpperCase()}-001`,
      BudgetCode: `${r.amount > 20000 ? 'CAPEX' : 'OPEX'}-${r.dept.substring(0, 3).toUpperCase()}-2024`,
      ApprovedById: r.status === 'Approved' ? 1 : null,
      ApprovedDate: r.status === 'Approved' ? new Date(Date.now() - (25 - i * 3) * 24 * 60 * 60 * 1000).toISOString() : null,
      Created: new Date(Date.now() - (30 - i * 3) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - i * 2 * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock purchase orders
   */
  public static getMockPurchaseOrders(): any[] {
    const statuses = ['Draft', 'Sent', 'Acknowledged', 'Shipped', 'Received', 'Closed', 'Cancelled'];
    const paymentTerms = ['Net 15', 'Net 30', 'Net 45', 'Due on Receipt', 'Monthly'];

    const orders = [
      { title: 'Engineering Laptops', vendor: 7, amount: 48937.50, status: 'Shipped' },
      { title: 'Office Supplies October', vendor: 2, amount: 2718.75, status: 'Received' },
      { title: 'Cloud Services Q4', vendor: 6, amount: 25000, status: 'Active' },
      { title: 'Security Services Annual', vendor: 4, amount: 180000, status: 'Active' },
      { title: 'Facilities Maintenance Q4', vendor: 5, amount: 15000, status: 'Active' },
      { title: 'Marketing Collateral', vendor: 2, amount: 8500, status: 'Sent' },
      { title: 'Server Hardware', vendor: 7, amount: 65000, status: 'Draft' }
    ];

    return orders.map((o, i) => ({
      Id: i + 1,
      Title: `PO-2024-${padStart(String(i + 1), 3, '0')} - ${o.title}`,
      PONumber: `PO-2024-${padStart(String(i + 1), 3, '0')}`,
      VendorId: o.vendor,
      RequisitionId: i < 5 ? i + 1 : null,
      OrderDate: new Date(Date.now() - (45 - i * 5) * 24 * 60 * 60 * 1000).toISOString(),
      DeliveryDate: ['Shipped', 'Received'].includes(o.status) ? new Date(Date.now() + (i * 5) * 24 * 60 * 60 * 1000).toISOString() : null,
      Status: o.status,
      PaymentTerms: paymentTerms[i % paymentTerms.length],
      ShippingMethod: o.status === 'Active' ? 'N/A - Service' : 'Ground - Standard',
      ShippingAddress: o.status === 'Active' ? null : '123 Corporate Drive, San Francisco, CA 94105',
      SubTotal: o.amount * 0.9125,
      TaxAmount: o.amount * 0.0875,
      ShippingCost: 0,
      TotalAmount: o.amount,
      Currency: 'USD',
      Notes: `Purchase order for ${o.title}`,
      TrackingNumber: o.status === 'Shipped' ? `1Z999AA1012345678${i}` : null,
      Created: new Date(Date.now() - (45 - i * 5) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - i * 3 * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock contracts
   */
  public static getMockContracts(): any[] {
    const contractTypes = ['Service Agreement', 'Framework Agreement', 'Subscription', 'Retainer', 'License Agreement'];
    const statuses = ['Active', 'Active', 'Active', 'Expiring Soon', 'Expired', 'Draft'];

    const contracts = [
      { title: 'IT Infrastructure Support Agreement', vendor: 1, value: 500000, annual: 166667, type: 'Service Agreement' },
      { title: 'Office Supplies Framework Agreement', vendor: 2, value: 50000, annual: 25000, type: 'Framework Agreement' },
      { title: 'Cybersecurity Services Contract', vendor: 4, value: 360000, annual: 180000, type: 'Service Agreement' },
      { title: 'Cloud Hosting Agreement', vendor: 6, value: 300000, annual: 100000, type: 'Subscription' },
      { title: 'Legal Retainer Agreement', vendor: 8, value: 120000, annual: 120000, type: 'Retainer' },
      { title: 'Facilities Maintenance Contract', vendor: 5, value: 90000, annual: 60000, type: 'Service Agreement' }
    ];

    return contracts.map((c, i) => ({
      Id: i + 1,
      Title: c.title,
      ContractNumber: `CON-2024-${padStart(String(i + 1), 3, '0')}`,
      VendorId: c.vendor,
      ContractType: c.type,
      Status: i < 4 ? 'Active' : (i === 4 ? 'Expiring Soon' : 'Draft'),
      StartDate: new Date(Date.now() - (365 - i * 60) * 24 * 60 * 60 * 1000).toISOString(),
      EndDate: new Date(Date.now() + (365 + i * 90) * 24 * 60 * 60 * 1000).toISOString(),
      TotalValue: c.value,
      AnnualValue: c.annual,
      Currency: 'USD',
      AutoRenewal: i % 2 === 0,
      RenewalTermMonths: i % 2 === 0 ? 12 : null,
      NoticePeriodDays: [30, 60, 90][i % 3],
      Description: `${c.title} - comprehensive agreement covering all services and deliverables.`,
      TermsAndConditions: 'Standard terms and conditions apply. See attached document for details.',
      PaymentTerms: ['Net 30', 'Monthly', 'Annual Prepaid'][i % 3],
      Notes: `Contract notes for ${c.title}`,
      Created: new Date(Date.now() - (365 - i * 60) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - i * 14 * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock invoices
   */
  public static getMockInvoices(): any[] {
    const statuses = ['Paid', 'Paid', 'Approved', 'Pending Approval', 'Disputed', 'Overdue'];
    const paymentMethods = ['ACH Transfer', 'Wire Transfer', 'Check', 'Credit Card'];

    const invoices = [
      { title: 'IT Support - October', vendor: 1, amount: 13888.89, status: 'Paid' },
      { title: 'Office Supplies - October', vendor: 2, amount: 2718.75, status: 'Paid' },
      { title: 'Security Services - Q4', vendor: 4, amount: 45000, status: 'Paid' },
      { title: 'Cloud Services - October', vendor: 6, amount: 8750, status: 'Approved' },
      { title: 'Facilities - November', vendor: 5, amount: 5000, status: 'Pending Approval' },
      { title: 'Engineering Laptops', vendor: 7, amount: 48937.50, status: 'Pending Approval' },
      { title: 'Legal Services - October', vendor: 8, amount: 12500, status: 'Disputed' }
    ];

    return invoices.map((inv, i) => ({
      Id: i + 1,
      Title: `INV-${['ACM', 'GLO', 'SEC', 'CFS', 'ECO', 'PRE', 'LEG'][i]}-2024-${1000 + i}`,
      InvoiceNumber: `INV-${['ACM', 'GLO', 'SEC', 'CFS', 'ECO', 'PRE', 'LEG'][i]}-2024-${1000 + i}`,
      VendorId: inv.vendor,
      PurchaseOrderId: i < 5 ? i + 1 : null,
      ContractId: i < 6 ? i + 1 : null,
      InvoiceDate: new Date(Date.now() - (30 - i * 3) * 24 * 60 * 60 * 1000).toISOString(),
      DueDate: new Date(Date.now() + (i * 5) * 24 * 60 * 60 * 1000).toISOString(),
      Status: inv.status,
      SubTotal: inv.amount * 0.9125,
      TaxAmount: inv.amount * 0.0875,
      TotalAmount: inv.amount,
      PaidAmount: inv.status === 'Paid' ? inv.amount : 0,
      Currency: 'USD',
      PaymentMethod: inv.status === 'Paid' ? paymentMethods[i % paymentMethods.length] : null,
      PaymentDate: inv.status === 'Paid' ? new Date(Date.now() - (20 - i * 2) * 24 * 60 * 60 * 1000).toISOString() : null,
      PaymentReference: inv.status === 'Paid' ? `PAY-2024-${1000 + i}` : null,
      Description: inv.title,
      Notes: `Invoice for ${inv.title}`,
      DisputeReason: inv.status === 'Disputed' ? 'Unauthorized charges - under review' : null,
      DisputeDate: inv.status === 'Disputed' ? new Date(Date.now() - 5 * 24 * 60 * 60 * 1000).toISOString() : null,
      Created: new Date(Date.now() - (30 - i * 3) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - i * 2 * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock budgets
   */
  public static getMockBudgets(): any[] {
    const departments = ['IT', 'IT Security', 'Operations', 'Marketing', 'Legal', 'IT', 'IT', 'Facilities'];
    const categories = ['Infrastructure', 'Security', 'General Operations', 'Marketing', 'Professional Services', 'Cloud', 'Capital Expenditure', 'Facilities'];

    const budgets = [
      { title: 'IT Infrastructure 2024', total: 500000, allocated: 450000, spent: 380000 },
      { title: 'Security Operations 2024', total: 400000, allocated: 385000, spent: 315000 },
      { title: 'Office Operations 2024', total: 150000, allocated: 125000, spent: 98000 },
      { title: 'Marketing Programs 2024', total: 250000, allocated: 220000, spent: 175000 },
      { title: 'Legal & Compliance 2024', total: 200000, allocated: 180000, spent: 142000 },
      { title: 'Cloud Services 2024', total: 300000, allocated: 275000, spent: 210000 },
      { title: 'Capital Equipment 2024', total: 200000, allocated: 165000, spent: 95000 },
      { title: 'Facilities Management 2024', total: 180000, allocated: 160000, spent: 135000 }
    ];

    return budgets.map((b, i) => ({
      Id: i + 1,
      Title: b.title,
      BudgetCode: `BUD-${departments[i].substring(0, 3).toUpperCase()}-2024`,
      FiscalYear: '2024',
      Department: departments[i],
      Category: categories[i],
      TotalBudget: b.total,
      AllocatedAmount: b.allocated,
      SpentAmount: b.spent,
      RemainingAmount: b.total - b.spent,
      Currency: 'USD',
      Status: 'Active',
      StartDate: '2024-01-01',
      EndDate: '2024-12-31',
      Notes: `Annual budget for ${departments[i]} - ${categories[i]}`,
      Created: new Date(2024, 0, 1).toISOString(),
      Modified: new Date(Date.now() - i * 7 * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock catalog items
   */
  public static getMockCatalogItems(): any[] {
    const items = [
      { title: 'Dell XPS 15 Laptop', code: 'HW-LAP-001', cat: 'Hardware', subCat: 'Laptops', price: 2500, vendor: 7, asset: true },
      { title: 'Dell Docking Station WD19TBS', code: 'HW-DOC-001', cat: 'Hardware', subCat: 'Accessories', price: 300, vendor: 7, asset: true },
      { title: 'Professional Laptop Bag', code: 'HW-BAG-001', cat: 'Hardware', subCat: 'Accessories', price: 200, vendor: 7, asset: false },
      { title: 'Dell 27" 4K Monitor', code: 'HW-MON-001', cat: 'Hardware', subCat: 'Monitors', price: 650, vendor: 7, asset: true },
      { title: 'Logitech MX Master 3 Mouse', code: 'HW-MOU-001', cat: 'Hardware', subCat: 'Peripherals', price: 100, vendor: 7, asset: false },
      { title: 'Copy Paper - Premium', code: 'OS-PAP-001', cat: 'Office Supplies', subCat: 'Paper Products', price: 45, vendor: 2, asset: false },
      { title: 'HP LaserJet Toner - Black', code: 'OS-TON-001', cat: 'Office Supplies', subCat: 'Printer Supplies', price: 120, vendor: 2, asset: false },
      { title: 'Office Supplies Starter Kit', code: 'OS-KIT-001', cat: 'Office Supplies', subCat: 'Kits', price: 80, vendor: 2, asset: false },
      { title: 'CrowdStrike Falcon - Annual License', code: 'SW-SEC-001', cat: 'Software', subCat: 'Security', price: 150, vendor: 4, asset: false },
      { title: 'Microsoft 365 E5 - Annual', code: 'SW-M365-001', cat: 'Software', subCat: 'Productivity', price: 57, vendor: 1, asset: false },
      { title: 'Azure Reserved Instance - D4s v3', code: 'CLD-AZ-001', cat: 'Cloud Services', subCat: 'Compute', price: 115, vendor: 6, asset: false },
      { title: 'Professional Services - IT Consulting', code: 'SVC-CON-001', cat: 'Services', subCat: 'Consulting', price: 175, vendor: 1, asset: false },
      { title: 'Ergonomic Office Chair', code: 'FUR-CHR-001', cat: 'Furniture', subCat: 'Seating', price: 1395, vendor: 2, asset: true },
      { title: 'Standing Desk - Electric', code: 'FUR-DSK-001', cat: 'Furniture', subCat: 'Desks', price: 850, vendor: 2, asset: true },
      { title: 'Janitorial Services - Monthly', code: 'SVC-JAN-001', cat: 'Services', subCat: 'Facilities', price: 5000, vendor: 5, asset: false }
    ];

    return items.map((item, i) => ({
      Id: i + 1,
      Title: item.title,
      ItemCode: item.code,
      Category: item.cat,
      SubCategory: item.subCat,
      Description: `${item.title} - Standard catalog item for procurement.`,
      DefaultPrice: item.price,
      Currency: 'USD',
      UnitOfMeasure: item.cat === 'Services' ? (item.subCat === 'Consulting' ? 'Hour' : 'Month') : 'Each',
      PreferredVendorId: item.vendor,
      LeadTimeDays: item.cat === 'Hardware' ? 5 : (item.cat === 'Services' ? 0 : 2),
      MinOrderQuantity: 1,
      IsActive: true,
      CreateAssetOnReceipt: item.asset,
      AssetCategory: item.asset ? (item.cat === 'Furniture' ? 'Office Furniture' : 'Computer Equipment') : null,
      Notes: `Catalog item: ${item.title}`,
      Created: new Date(Date.now() - (180 - i * 10) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - i * 5 * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock vendor performance reviews
   */
  public static getMockVendorPerformance(): any[] {
    const reviews = [
      { vendorId: 1, quality: 4.8, delivery: 4.7, comm: 4.9, pricing: 4.5, overall: 4.7, onTime: 96, defect: 1.2 },
      { vendorId: 1, quality: 4.9, delivery: 4.8, comm: 4.9, pricing: 4.6, overall: 4.8, onTime: 98, defect: 0.8 },
      { vendorId: 2, quality: 4.5, delivery: 4.3, comm: 4.4, pricing: 4.8, overall: 4.5, onTime: 92, defect: 2.1 },
      { vendorId: 4, quality: 5.0, delivery: 4.9, comm: 4.8, pricing: 4.2, overall: 4.7, onTime: 99, defect: 0.1 },
      { vendorId: 6, quality: 4.6, delivery: 4.7, comm: 4.8, pricing: 4.4, overall: 4.6, onTime: 95, defect: 1.5 }
    ];

    return reviews.map((r, i) => ({
      Id: i + 1,
      Title: `Performance Review - Vendor ${r.vendorId}`,
      VendorId: r.vendorId,
      ReviewDate: new Date(Date.now() - (90 - i * 30) * 24 * 60 * 60 * 1000).toISOString(),
      ReviewPeriodStart: new Date(Date.now() - (180 - i * 30) * 24 * 60 * 60 * 1000).toISOString(),
      ReviewPeriodEnd: new Date(Date.now() - (90 - i * 30) * 24 * 60 * 60 * 1000).toISOString(),
      QualityRating: r.quality,
      DeliveryRating: r.delivery,
      CommunicationRating: r.comm,
      PricingRating: r.pricing,
      OverallRating: r.overall,
      OnTimeDeliveryPercent: r.onTime,
      DefectRate: r.defect,
      ResponseTime: 2 + Math.random() * 3,
      Comments: `Performance review comments for vendor ${r.vendorId}. Overall performance meets expectations.`,
      Recommendations: 'Continue partnership and review in next quarter.',
      ReviewerId: 1,
      Created: new Date(Date.now() - (90 - i * 30) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - (90 - i * 30) * 24 * 60 * 60 * 1000).toISOString()
    }));
  }

  /**
   * Get mock budget transactions
   */
  public static getMockBudgetTransactions(): any[] {
    const transactions = [
      { budgetId: 1, type: 'Allocation', amount: 50000, desc: 'Q4 allocation for hardware refresh', ref: 'ALO-2024-001' },
      { budgetId: 1, type: 'Spend', amount: 48937.50, desc: 'PO-2024-001 - Engineering Laptops', ref: 'PO-2024-001' },
      { budgetId: 2, type: 'Spend', amount: 45000, desc: 'Q4 Security Services Prepayment', ref: 'INV-SEC-2024-Q4' },
      { budgetId: 3, type: 'Spend', amount: 2718.75, desc: 'October Office Supplies', ref: 'INV-GLO-2024-5678' },
      { budgetId: 4, type: 'Allocation', amount: 15000, desc: 'Trade show budget allocation', ref: 'ALO-2024-002' },
      { budgetId: 5, type: 'Spend', amount: 10000, desc: 'October Legal Retainer', ref: 'INV-LEG-2024-OCT' },
      { budgetId: 6, type: 'Spend', amount: 8750, desc: 'October Cloud Usage', ref: 'INV-CFS-2024-1015' },
      { budgetId: 8, type: 'Spend', amount: 5000, desc: 'November Facilities Maintenance', ref: 'INV-ECO-2024-NOV' }
    ];

    return transactions.map((t, i) => ({
      Id: i + 1,
      Title: `${t.type} - ${t.ref}`,
      BudgetId: t.budgetId,
      TransactionDate: new Date(Date.now() - (30 - i * 3) * 24 * 60 * 60 * 1000).toISOString(),
      TransactionType: t.type,
      Amount: t.amount,
      Description: t.desc,
      ReferenceNumber: t.ref,
      CreatedBy: t.type === 'Allocation' ? 'Finance Team' : 'Accounts Payable',
      Created: new Date(Date.now() - (30 - i * 3) * 24 * 60 * 60 * 1000).toISOString(),
      Modified: new Date(Date.now() - (30 - i * 3) * 24 * 60 * 60 * 1000).toISOString()
    }));
  }
}
