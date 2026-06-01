/* ========================================
   Mock Geotab SDK for Work Order Print Add-in
   Simulates api.call(), api.multiCall(), and state
   ======================================== */

const MOCK_SCENARIOS = {
  'brake-overhaul': {
    companyDetails: {
      companyName: 'Northstar Fleet Services Inc.',
      phoneNumber: '(905) 555-1234',
      country: 'CA'
    },
    loggedInUser: {
      id: 'aUser100',
      name: 'Maria Santos',
      firstName: 'Maria',
      lastName: 'Santos',
      companyName: 'Northstar Fleet Services Inc.',
      companyAddress: '2440 Winston Park Dr, Oakville, ON L6H 7V2',
      phoneNumber: '(905) 555-1234'
    },
    workOrder: {
      id: 'aWO001',
      reference: 'WO-2026-0342',
      description: 'Front and rear brake inspection and replacement',
      notes: 'Driver reported squealing noise when braking. Check rotors for warping. Replace pads if below 3mm. Vehicle pulling to the left under braking — inspect calipers.',
      status: 'InProgress',
      priority: 'Medium',
      dateOpened: '2026-03-15T09:00:00Z',
      dateClosed: null,
      device: { id: 'b1234' },
      openedByUser: { id: 'aUser100' },
      assignedToUser: { id: 'aUser200' },
      laborCost: 280.00,
      partsCost: 347.50,
      otherCost: 15.00,
      taxCost: 51.40,
      totalCost: 693.90
    },
    jobs: [
      {
        id: 'aJob001',
        workOrder: { id: 'aWO001' },
        name: 'Front Brake Pad Replacement',
        description: 'Remove front wheels, inspect rotors, replace brake pads. Torque lug nuts to spec.',
        notes: 'Rotors measured at 28.2mm — within spec, no resurfacing needed.',
        status: 'Completed',
        laborCost: 160.00,
        partsCost: 189.50,
        otherCost: 0,
        jobType: { id: 'aJT001', name: 'Brake Service' }
      },
      {
        id: 'aJob002',
        workOrder: { id: 'aWO001' },
        name: 'Rear Brake Pad Replacement',
        description: 'Remove rear wheels, inspect drums/rotors, replace brake pads or shoes as needed.',
        notes: 'Rear drums in good condition. Shoes replaced — adjusted to proper clearance.',
        status: 'InProgress',
        laborCost: 120.00,
        partsCost: 158.00,
        otherCost: 15.00,
        jobType: { id: 'aJT001', name: 'Brake Service' }
      }
    ],
    device: {
      id: 'b1234',
      name: 'Fleet-Truck-201',
      vehicleIdentificationNumber: '1GCGG25K081234567',
      licensePlate: 'ON · CVXR 482',
      make: 'Chevrolet',
      model: 'Silverado 2500HD',
      year: '2022',
      odometer: 87432.6,
      engineHours: 3241.5
    },
    openedByUser: {
      id: 'aUser100',
      name: 'Maria Santos',
      firstName: 'Maria',
      lastName: 'Santos',
      companyName: 'Northstar Fleet Services Inc.',
      companyAddress: '2440 Winston Park Dr, Oakville, ON L6H 7V2',
      phoneNumber: '(905) 555-1234'
    },
    assignedToUser: {
      id: 'aUser200',
      name: 'James Wright',
      firstName: 'James',
      lastName: 'Wright'
    }
  },

  'oil-change': {
    companyDetails: {
      companyName: 'Northstar Fleet Services Inc.',
      phoneNumber: '(905) 555-1234',
      country: 'CA'
    },
    loggedInUser: {
      id: 'aUser100',
      name: 'Maria Santos',
      firstName: 'Maria',
      lastName: 'Santos',
      companyName: 'Northstar Fleet Services Inc.',
      companyAddress: '2440 Winston Park Dr, Oakville, ON L6H 7V2',
      phoneNumber: '(905) 555-1234'
    },
    workOrder: {
      id: 'aWO002',
      reference: 'WO-2026-0343',
      description: 'Routine oil and filter change',
      notes: '',
      status: 'Pending',
      priority: 'Low',
      dateOpened: '2026-03-17T14:30:00Z',
      dateClosed: null,
      device: { id: 'b5678' },
      openedByUser: { id: 'aUser100' },
      assignedToUser: null,
      laborCost: 45.00,
      partsCost: 62.00,
      otherCost: 0,
      taxCost: 13.91,
      totalCost: 120.91
    },
    jobs: [
      {
        id: 'aJob003',
        workOrder: { id: 'aWO002' },
        name: 'Oil and Filter Change',
        description: 'Drain old oil, replace oil filter, fill with 5W-30 synthetic (8 quarts). Reset oil life monitor.',
        notes: '',
        status: 'Pending',
        laborCost: 45.00,
        partsCost: 62.00,
        otherCost: 0,
        jobType: { id: 'aJT002', name: 'Preventive Maintenance' }
      }
    ],
    device: {
      id: 'b5678',
      name: 'Delivery-Van-44',
      vehicleIdentificationNumber: '1FTFW1ET5DFC10042',
      licensePlate: 'ON · BKML 917',
      make: 'Ford',
      model: 'Transit 250',
      year: '2024',
      odometer: 23105.2,
      engineHours: 891.3
    },
    openedByUser: {
      id: 'aUser100',
      name: 'Maria Santos',
      firstName: 'Maria',
      lastName: 'Santos',
      companyName: 'Northstar Fleet Services Inc.',
      companyAddress: '2440 Winston Park Dr, Oakville, ON L6H 7V2',
      phoneNumber: '(905) 555-1234'
    },
    assignedToUser: null
  }
};

function createMockApi(scenario) {
  const data = MOCK_SCENARIOS[scenario] || MOCK_SCENARIOS['brake-overhaul'];

  const api = {
    _isMock: true,
    _scenario: scenario,
    _callLog: [],

    call(method, params) {
      return new Promise((resolve) => {
        const delay = 50 + Math.random() * 100;
        setTimeout(() => {
          this._callLog.push({ method, params, timestamp: Date.now() });
          console.log(`[Mock API] ${method}(${params.typeName})`, params);

          if (method === 'Get') {
            switch (params.typeName) {
              case 'MaintenanceWorkOrder':
                resolve([data.workOrder]);
                break;
              case 'MaintenanceWorkOrderJob':
                resolve(data.jobs);
                break;
              case 'Device':
                resolve([data.device]);
                break;
              case 'CompanyDetails':
                resolve([data.companyDetails]);
                break;
              case 'User':
                // Return correct user based on search ID
                if (params.search && params.search.id === data.openedByUser.id) {
                  resolve([data.openedByUser]);
                } else if (data.assignedToUser && params.search && params.search.id === data.assignedToUser.id) {
                  resolve([data.assignedToUser]);
                } else {
                  resolve([data.openedByUser]);
                }
                break;
              default:
                resolve([]);
            }
          } else if (method === 'GetSession') {
            resolve({ userName: data.loggedInUser.name, database: 'mock-db' });
          } else {
            resolve([]);
          }
        }, delay);
      });
    },

    multiCall(calls) {
      console.log(`[Mock API] multiCall with ${calls.length} calls`);
      return Promise.all(
        calls.map(([method, params]) => this.call(method, params))
      );
    },

    getSession(callback) {
      const session = { userName: data.loggedInUser.name, database: 'mock-db' };
      if (callback) callback(session);
      return Promise.resolve(session);
    }
  };

  const state = {
    device: data.device
  };

  return { api, state, mockData: data };
}
