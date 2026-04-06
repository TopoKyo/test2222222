export type UserRole = 'admin' | 'worker';

export interface UserProfile {
  uid: string;
  email: string;
  name?: string;
  role: UserRole;
  createdAt?: any;
}

export interface Product {
  id: string;
  name: string;
  category: string;
  quantity: number;
  unit: string;
  imageUrl?: string;
  location?: string;
  entryDate?: any;
  updatedAt?: any;
  status: 'available' | 'ordered';
  useQR: boolean;
}

export interface Requester {
  id: string;
  name: string;
  createdAt: any;
}

export interface Movement {
  id: string;
  productId: string;
  productName: string;
  type: 'entry' | 'exit';
  quantity: number;
  date: any;
  userId: string;
  userName?: string;
  givenBy?: string;
  receivedBy?: string;
  notes?: string;
}

export interface Attendance {
  id: string;
  userId: string;
  userName: string;
  date: string; // YYYY-MM-DD
  arrival?: any;
  breakStart?: any;
  breakEnd?: any;
  departure?: any;
  totalHours?: number;
  status: 'not_started' | 'working' | 'on_break' | 'finished';
}

export interface ProductRequest {
  id: string;
  productName: string;
  category: string;
  quantity: number;
  unit: string;
  priority: 'low' | 'medium' | 'high';
  comment?: string;
  status: 'pending' | 'in_process' | 'purchased' | 'rejected';
  userId: string;
  userName: string;
  createdAt: any;
  updatedAt?: any;
}

export interface DailyReport {
  id: string;
  date: any; // serverTimestamp
  userId: string;
  userName: string;
  advances: string;
  issues: string;
  delays: string;
  nextDayPlan: string;
  createdAt: any;
}
