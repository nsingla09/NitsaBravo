/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, useCallback, Component, useRef } from 'react';
import React from 'react';
import { useQuery, useQueryClient } from '@tanstack/react-query';
import * as XLSX from 'xlsx';
import { 
  auth, 
  db, 
  googleProvider, 
  signInWithPopup, 
  signOut, 
  onAuthStateChanged,
  doc, 
  getDoc, 
  setDoc, 
  updateDoc, 
  deleteDoc,
  collection, 
  query, 
  where,
  getDocs,
  orderBy, 
  limit, 
  onSnapshot, 
  addDoc, 
  serverTimestamp,
  runTransaction,
  loadBundle,
  namedQuery,
  handleFirestoreError,
  OperationType,
  or,
  and,
  increment,
  getDocFromCache,
  getDocsFromCache,
  getAggregateFromServer,
  sum,
  count
} from './firebase';
import { 
  BarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip, Legend, ResponsiveContainer, 
  PieChart, Pie, Cell, LineChart, Line, AreaChart, Area, LabelList
} from 'recharts';
import { backfillSalesSummaries } from './scripts/backfillSalesSummaries';
import { 
  Trophy, 
  Send, 
  History, 
  LogOut, 
  Search, 
  Award, 
  Star, 
  TrendingUp,
  ArrowUpRight,
  ArrowDownRight,
  Info,
  Download,
  Shield,
  User as UserIcon,
  ChevronRight,
  Loader2,
  CheckCircle2,
  AlertCircle,
  Users,
  Plus,
  Filter,
  Calendar as CalendarIcon,
  Edit2,
  Trash2,
  IndianRupee,
  MapPin,
  Calculator,
  Menu,
  Clock,
  ExternalLink,
  BarChart3,
  PieChart as PieChartIcon,
  DollarSign,
  Target,
  Briefcase,
  MessageSquare,
  X,
  LayoutGrid,
  Lock,
  PhoneCall,
  Database,
  RefreshCw
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { formatDistanceToNow, format, subDays } from 'date-fns';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { AIChatBot } from './components/AIChatBot';

// --- Utilities ---
const generateTalktimeOptions = () => {
  const options = [];
  // Increased to 15 hours (900 minutes) to cover all possible shift lengths
  for (let i = 0; i <= 900; i += 5) {
    const hours = Math.floor(i / 60);
    const mins = i % 60;
    let label = '';
    if (hours > 0) label += `${hours}h `;
    if (mins > 0 || hours === 0) label += `${mins}m`;
    options.push({ value: label.trim(), minutes: i });
  }
  return options;
};

const talktimeToMinutes = (str: string): number => {
  if (!str) return 0;
  let totalMinutes = 0;
  const hoursMatch = str.match(/(\d+)h/);
  const minutesMatch = str.match(/(\d+)m/);
  if (hoursMatch) totalMinutes += parseInt(hoursMatch[1]) * 60;
  if (minutesMatch) totalMinutes += parseInt(minutesMatch[1]);
  if (!hoursMatch && !minutesMatch && str.includes('m')) {
    const m = str.replace('m', '').trim();
    if (!isNaN(parseInt(m))) totalMinutes = parseInt(m);
  }
  return totalMinutes;
};

const minutesToTalktime = (mins: number): string => {
  if (mins === 0) return '0m';
  const hours = Math.floor(mins / 60);
  const minutes = mins % 60;
  let result = '';
  if (hours > 0) result += `${hours}h `;
  if (minutes > 0 || result === '') result += `${minutes}m`;
  return result.trim();
};

function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Components ---
const Last7DaysSalesOverview = ({ sales, employees, sourceFilter = 'All', bdeFilter = 'All' }: { sales: any[], employees: any[], sourceFilter?: string, bdeFilter?: string }) => {
  const last7Days = Array.from({ length: 7 }, (_, i) => {
    const d = new Date();
    d.setDate(d.getDate() - i);
    return d.toISOString().split('T')[0];
  });

  const employeeBdeMap = React.useMemo(() => employees.reduce((acc, emp) => {
    acc[emp.name.toLowerCase().trim()] = emp.bde;
    return acc;
  }, {} as Record<string, string>), [employees]);

  const last7DaysDataRaw = sales.reduce((acc: any, sale) => {
    const saleDate = sale.date || 'Unknown';
    if (last7Days.includes(saleDate)) {
      const matchesSource = sourceFilter === 'All' || sale.source === sourceFilter;
      
      if (matchesSource) {
        const salesBy = sale.salesBy || 'Inhouse';
        
        // Calculate credits for each person involved
        const credits = [];
        
        // Agent's credit
        const agentName = sale.agent?.toLowerCase().trim();
        if (agentName) {
          const agentBde = employeeBdeMap[agentName] || sale.bde || 'Unknown';
          const agentCredit = (sale.packageValue || 0) * ((sale.agentPercentage ?? 100) / 100);
          if (agentCredit > 0 || sale.agentPercentage === 0) { // Include even if 0% if they are the primary agent
            credits.push({ bde: agentBde, value: agentCredit, isPrimary: true });
          }
        }
        
        // Associate BDE's credit
        const assocName = sale.associateBde?.toLowerCase().trim();
        if (assocName) {
          const assocBde = employeeBdeMap[assocName] || sale.bde || 'Unknown';
          const assocCredit = (sale.packageValue || 0) * ((sale.assocBdePercentage ?? 0) / 100);
          if (assocCredit > 0) {
            credits.push({ bde: assocBde, value: assocCredit, isPrimary: false });
          }
        }

        credits.forEach(credit => {
          const matchesBde = bdeFilter === 'All' || credit.bde === bdeFilter;
          if (matchesBde) {
            const key = `${saleDate}_${credit.bde}`;
            
            if (!acc[key]) {
              acc[key] = {
                date: saleDate,
                bde: credit.bde,
                inhouseCount: 0,
                inhouseValue: 0,
                branchCount: 0,
                branchValue: 0,
                franchiseeCount: 0,
                franchiseeValue: 0,
                totalCount: 0,
                totalValue: 0
              };
            }
            
            // Only count the sale once per BDE if possible, but here we are splitting values
            // If we split the value, we should probably split the count too or only count primary?
            // Let's count primary as 1 and associate as 0 for "Count" to avoid inflation,
            // but use the split value for "Value".
            if (credit.isPrimary) {
              acc[key].totalCount += 1;
              if (salesBy === 'Inhouse') acc[key].inhouseCount += 1;
              else if (salesBy === 'Branch') acc[key].branchCount += 1;
              else if (salesBy === 'Franchisee Sales') acc[key].franchiseeCount += 1;
            }

            acc[key].totalValue += credit.value;
            
            if (salesBy === 'Inhouse') {
              acc[key].inhouseValue += credit.value;
            } else if (salesBy === 'Branch') {
              acc[key].branchValue += credit.value;
            } else if (salesBy === 'Franchisee Sales') {
              acc[key].franchiseeValue += credit.value;
            }
          }
        });
      }
    }
    return acc;
  }, {});

  const last7DaysData = Object.values(last7DaysDataRaw).sort((a: any, b: any) => {
    if (a.date !== b.date) return a.date.localeCompare(b.date);
    return a.bde.localeCompare(b.bde);
  });

  const dayTotals: Record<string, number> = {};
  last7DaysData.forEach((row: any) => {
    if (!dayTotals[row.date]) dayTotals[row.date] = 0;
    dayTotals[row.date] += row.totalValue;
  });

  const grandTotals: any = last7DaysData.reduce((acc: any, row: any) => {
    acc.inhouseCount += row.inhouseCount;
    acc.inhouseValue += row.inhouseValue;
    acc.branchCount += row.branchCount;
    acc.branchValue += row.branchValue;
    acc.franchiseeCount += row.franchiseeCount;
    acc.franchiseeValue += row.franchiseeValue;
    acc.totalCount += row.totalCount;
    acc.totalValue += row.totalValue;
    return acc;
  }, {
    inhouseCount: 0, inhouseValue: 0,
    branchCount: 0, branchValue: 0,
    franchiseeCount: 0, franchiseeValue: 0,
    totalCount: 0, totalValue: 0
  });

  const maxVal = Math.max(...last7DaysData.map((r: any) => r.totalValue), 1);

  const getGreenTone = (val: number) => {
    if (val <= 0) return '';
    const ratio = val / maxVal;
    if (ratio > 0.8) return 'bg-emerald-600 text-white';
    if (ratio > 0.6) return 'bg-emerald-500 text-white';
    if (ratio > 0.4) return 'bg-emerald-400 text-white';
    if (ratio > 0.2) return 'bg-emerald-200 text-emerald-900';
    return 'bg-emerald-50 text-emerald-900';
  };

  return (
    <Card className="p-6 overflow-hidden">
      <div className="flex items-center gap-2 mb-6">
        <CalendarIcon className="w-4 h-4 text-orange-600" />
        <h3 className="text-sm font-bold text-zinc-900 uppercase">Last 7 Days Sales Overview</h3>
      </div>
      <div className="overflow-x-auto -mx-6 px-6">
        <table className="w-full text-left text-xs border-collapse">
          <thead className="bg-zinc-50 text-zinc-500 uppercase font-bold border-y border-zinc-100">
            <tr>
              <th className="px-4 py-2 border-r border-zinc-100">Date</th>
              <th className="px-4 py-2 border-r border-zinc-100">BDE</th>
              <th className="px-4 py-2 border-r border-zinc-100 text-center" colSpan={3}>Inhouse (S/A/%)</th>
              <th className="px-4 py-2 border-r border-zinc-100 text-center" colSpan={3}>Branch (S/A/%)</th>
              <th className="px-4 py-2 border-r border-zinc-100 text-center" colSpan={3}>Franchisee (S/A/%)</th>
              <th className="px-4 py-2 text-center bg-orange-50 text-orange-700" colSpan={2}>Total (S/A)</th>
            </tr>
            <tr className="bg-zinc-50/50 text-[11px]">
              <th className="px-4 py-1 border-r border-zinc-100"></th>
              <th className="px-4 py-1 border-r border-zinc-100"></th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">S</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">A</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">%</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">S</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">A</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">%</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">S</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">A</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center">%</th>
              <th className="px-2 py-1 border-r border-zinc-100 text-center bg-orange-50/50">S</th>
              <th className="px-2 py-1 text-center bg-orange-50/50">A</th>
            </tr>
          </thead>
          <tbody className="divide-y divide-zinc-100">
            {last7DaysData.length > 0 ? last7DaysData.map((row: any, idx: number) => (
              <tr key={`${row.date}_${row.bde}`} className="hover:bg-zinc-50 transition-colors">
                <td className="px-4 py-2 border-r border-zinc-100 font-medium text-zinc-600 whitespace-nowrap">
                  {idx === 0 || (last7DaysData[idx-1] as any).date !== row.date ? format(new Date(row.date), 'dd MMM') : ''}
                </td>
                <td className="px-4 py-2 border-r border-zinc-100 font-bold text-zinc-900 whitespace-nowrap">{row.bde}</td>
                
                {/* Inhouse */}
                <td className="px-2 py-2 border-r border-zinc-100 text-center font-bold">{row.inhouseCount}</td>
                <td className={cn("px-2 py-2 border-r border-zinc-100 text-center font-medium", getGreenTone(row.inhouseValue))}>₹{(row.inhouseValue/100000).toFixed(2)}L</td>
                <td className="px-2 py-2 border-r border-zinc-100 text-center text-zinc-400">{(dayTotals[row.date] > 0 ? (row.inhouseValue/dayTotals[row.date])*100 : 0).toFixed(0)}%</td>
                
                {/* Branch */}
                <td className="px-2 py-2 border-r border-zinc-100 text-center font-bold">{row.branchCount}</td>
                <td className={cn("px-2 py-2 border-r border-zinc-100 text-center font-medium", getGreenTone(row.branchValue))}>₹{(row.branchValue/100000).toFixed(2)}L</td>
                <td className="px-2 py-2 border-r border-zinc-100 text-center text-zinc-400">{(dayTotals[row.date] > 0 ? (row.branchValue/dayTotals[row.date])*100 : 0).toFixed(0)}%</td>
                
                {/* Franchisee */}
                <td className="px-2 py-2 border-r border-zinc-100 text-center font-bold">{row.franchiseeCount}</td>
                <td className={cn("px-2 py-2 border-r border-zinc-100 text-center font-medium", getGreenTone(row.franchiseeValue))}>₹{(row.franchiseeValue/100000).toFixed(2)}L</td>
                <td className="px-2 py-2 border-r border-zinc-100 text-center text-zinc-400">{(dayTotals[row.date] > 0 ? (row.franchiseeValue/dayTotals[row.date])*100 : 0).toFixed(0)}%</td>
                
                {/* Total */}
                <td className="px-2 py-2 border-r border-zinc-100 text-center font-bold bg-orange-50/30 text-orange-700">{row.totalCount}</td>
                <td className={cn("px-2 py-2 text-center font-bold", getGreenTone(row.totalValue))}>₹{(row.totalValue/100000).toFixed(2)}L</td>
              </tr>
            )) : (
              <tr>
                <td colSpan={13} className="px-4 py-8 text-center text-zinc-400 italic">No sales data found for the last 7 days</td>
              </tr>
            )}
          </tbody>
          {last7DaysData.length > 0 && (
            <tfoot className="bg-zinc-100 font-black text-zinc-900 border-t-2 border-zinc-200">
              <tr>
                <td className="px-4 py-3 border-r border-zinc-200" colSpan={2}>GRAND TOTAL (7 DAYS)</td>
                
                {/* Inhouse Total */}
                <td className="px-2 py-3 border-r border-zinc-200 text-center">{grandTotals.inhouseCount}</td>
                <td className="px-2 py-3 border-r border-zinc-200 text-center">₹{(grandTotals.inhouseValue/100000).toFixed(2)}L</td>
                <td className="px-2 py-3 border-r border-zinc-200 text-center text-zinc-500">{(grandTotals.totalValue > 0 ? (grandTotals.inhouseValue/grandTotals.totalValue)*100 : 0).toFixed(1)}%</td>
                
                {/* Branch Total */}
                <td className="px-2 py-3 border-r border-zinc-200 text-center">{grandTotals.branchCount}</td>
                <td className="px-2 py-3 border-r border-zinc-200 text-center">₹{(grandTotals.branchValue/100000).toFixed(2)}L</td>
                <td className="px-2 py-3 border-r border-zinc-200 text-center text-zinc-500">{(grandTotals.totalValue > 0 ? (grandTotals.branchValue/grandTotals.totalValue)*100 : 0).toFixed(1)}%</td>
                
                {/* Franchisee Total */}
                <td className="px-2 py-3 border-r border-zinc-200 text-center">{grandTotals.franchiseeCount}</td>
                <td className="px-2 py-3 border-r border-zinc-200 text-center">₹{(grandTotals.franchiseeValue/100000).toFixed(2)}L</td>
                <td className="px-2 py-3 border-r border-zinc-200 text-center text-zinc-500">{(grandTotals.totalValue > 0 ? (grandTotals.franchiseeValue/grandTotals.totalValue)*100 : 0).toFixed(1)}%</td>
                
                {/* Grand Total */}
                <td className="px-2 py-3 border-r border-zinc-200 text-center bg-orange-100 text-orange-900">{grandTotals.totalCount}</td>
                <td className="px-2 py-3 text-center bg-orange-100 text-orange-900">₹{(grandTotals.totalValue/100000).toFixed(2)}L</td>
              </tr>
            </tfoot>
          )}
        </table>
      </div>
    </Card>
  );
};

// --- Types ---
interface Role {
  id: string;
  name: string;
  permissions: Record<string, 'Limited' | 'Complete'>;
  defaultPage?: string;
  createdAt: any;
}

interface UserProfile {
  uid: string;
  email: string;
  displayName: string;
  photoURL: string;
  pointsBalance: number;
  role?: string;
  employeeName?: string;
  employeeId?: string;
  bdeName?: string;
  permissions?: string[] | Record<string, 'Limited' | 'Complete' | string>;
  lastReadRemarks?: Record<string, number>;
}

interface Transaction {
  id: string;
  fromUid: string;
  fromName: string;
  toUid: string;
  toName: string;
  points: number;
  reason: string;
  createdAt: any;
}

interface TrainingMaterial {
  id?: string;
  remarks: string;
  module: ('Sales' | 'Operations' | 'Lead' | 'Both')[] | string;
  points: number;
  day: number;
  time: number;
  isActive: boolean;
  link: string;
  createdAt: any;
  priority?: number;
}

interface Employee {
  id?: string;
  employeeCode: string;
  name: string;
  email: string;
  target: number;
  position: string;
  joiningDate: string;
  joiningWeek: string;
  bde: string;
  status: 'Active' | 'Deactive';
  isAssocBDE?: boolean;
  week1Target: number;
  week2Target: number;
  week3Target: number;
  week4Target: number;
  week5Target: number;
  week6Target: number;
  salary: number;
  completedTrainings?: string[];
  trainingPoints?: Record<string, number>;
  trainingModuleStatus?: Record<string, 'Active' | 'Complete'>;
}

interface Week {
  id?: string;
  weekName: string;
  startDate: string;
  endDate: string;
  month: string;
  createdAt?: any;
}

interface BDE {
  id?: string;
  name: string;
  phone?: string;
  email?: string;
  createdAt?: any;
}

import { SearchableSelect } from './components/SearchableSelect';

const allPermissions = [
  { id: 'dashboard', label: 'Dashboard' },
  { id: 'employees', label: 'Employee Master' },
  { id: 'weeks', label: 'Week Master' },
  { id: 'bdes', label: 'BDE Master' },
  { id: 'sales', label: 'Sales Master' },
  { id: 'weeklyOverview', label: 'Weekly Overview' },
  { id: 'agentOverview', label: 'Agent Overview' },
  { id: 'userManagement', label: 'User Management' },
  { id: 'training', label: 'Training Module' },
  { id: 'trainingOverview', label: 'Training Overview' },
  { id: 'bdeOverview', label: 'BDE Overview' },
  { id: 'statsOverview', label: 'Stats Overview' },
  { id: 'issueOverview', label: 'Issue Overview' },
  { id: 'matrixMaster', label: 'Matrix Master' },
  { id: 'incentiveMaster', label: 'Incentive Master' },
  { id: 'passwordManager', label: 'Password Manager' },
];

interface Remark {
  id?: string;
  date: string;
  time: string;
  userName: string;
  userUid: string;
  text: string;
  createdAt: any;
}

interface Incentive {
  id?: string;
  employeeId: string;
  employeeName: string;
  type: 'Daily' | 'Weekly' | 'Monthly' | 'Quarterly' | 'Annually';
  amount: number;
  date: string;
  remarks: string;
  isEligible?: boolean;
  isPaid?: boolean;
  recordedBy?: string;
  recordedAt?: string;
  createdAt: any;
}

interface IncentivePayment {
  id?: string;
  employeeId: string;
  employeeName: string;
  amount: number;
  date: string;
  paymentMethod: string;
  remarks: string;
  recordedBy?: string;
  recordedAt?: string;
  createdAt: any;
}

interface MatrixReport {
  id?: string;
  employeeId: string;
  employeeName: string;
  date: string; // YYYY-MM-DD
  performancePoints: number;
  totalCalls: number;
  totalTalktime: string;
  topTalktimes: {
    tripId: string;
    phoneNo: string;
    talktime: string;
  }[];
  loginTime: string;
  logoutTime: string;
  breakTime: string;
  createdBy?: string;
  createdByName?: string;
  updatedBy?: string;
  updatedByName?: string;
  createdAt?: any;
  updatedAt?: any;
}

interface Sales {
  id?: string;
  week: string;
  date: string;
  guestName: string;
  agent: string;
  agentEmail?: string;
  agentPercentage?: number;
  bde: string;
  bdeEmail?: string;
  associateBde: string;
  associateBdeEmail?: string;
  assocBdePercentage?: number;
  destination: string;
  tripId: string;
  tripDate: string;
  packageValue?: number;
  lessThan10Percent20k: string;
  advanceCN?: string;
  salesBy?: 'Inhouse' | 'Branch' | 'Franchisee Sales';
  ppMargin?: number;
  noOfPax?: number;
  totalMargin?: number;
  flight?: number;
  source: string;
  converted: string;
  tasksPending: string;
  pHotel: string;
  pFlight: string;
  land: string;
  hflIssue: string;
  workPending: string;
  remarks: string;
  remarksCount?: number;
  createdAt?: any;
}

// --- Components ---

interface ErrorBoundaryProps {
  children: React.ReactNode;
}

interface ErrorBoundaryState {
  hasError: boolean;
  error: any;
}

export class ErrorBoundary extends Component<any, any> {
  constructor(props: any) {
    super(props);
    (this as any).state = { hasError: false, error: null };
  }

  static getDerivedStateFromError(error: any) {
    return { hasError: true, error };
  }

  componentDidCatch(error: any, errorInfo: any) {
    console.error("Uncaught error:", error, errorInfo);
  }

  render() {
    const self = this as any;
    if (self.state.hasError) {
      let message = "Something went wrong.";
      try {
        const errObj = JSON.parse(self.state.error.message);
        if (errObj.error) message = `Firestore Error: ${errObj.error}`;
      } catch (e) {
        message = self.state.error?.message || message;
      }
      return (
        <div className="min-h-screen flex items-center justify-center p-4 bg-red-50">
          <Card className="max-w-md w-full p-6 space-y-4 border-red-200">
            <div className="flex items-center gap-2 text-red-600">
              <AlertCircle className="w-6 h-6" />
              <h2 className="text-lg font-bold">Application Error</h2>
            </div>
            <p className="text-sm text-zinc-600">{message}</p>
            <Button onClick={() => window.location.reload()} variant="danger" className="w-full">
              Reload Application
            </Button>
          </Card>
        </div>
      );
    }

    return self.props.children;
  }
}

const Button = ({ 
  children, 
  className, 
  variant = 'primary', 
  isLoading = false, 
  size = 'md',
  ...props 
}: React.ButtonHTMLAttributes<HTMLButtonElement> & { 
  variant?: 'primary' | 'secondary' | 'outline' | 'ghost' | 'danger';
  isLoading?: boolean;
  size?: string;
}) => {
  const variants = {
    primary: 'bg-orange-600 text-white hover:bg-orange-700 shadow-md',
    secondary: 'bg-zinc-800 text-white hover:bg-zinc-900',
    outline: 'border border-zinc-300 text-zinc-700 hover:bg-zinc-50',
    ghost: 'text-zinc-600 hover:bg-zinc-100',
    danger: 'bg-red-500 text-white hover:bg-red-600',
  };

  return (
    <button 
      className={cn(
        'px-4 py-2 rounded-lg font-medium transition-all flex items-center justify-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed',
        variants[variant],
        className
      )}
      disabled={isLoading || props.disabled}
      {...props}
    >
      {isLoading && <Loader2 className="w-4 h-4 animate-spin" />}
      {children}
    </button>
  );
};

const Card = ({ children, className, id, onClick }: { children: React.ReactNode; className?: string; id?: string; onClick?: () => void; key?: string | number }) => (
  <div id={id} onClick={onClick} className={cn('bg-white border border-zinc-200 rounded-2xl shadow-sm overflow-hidden', className)}>
    {children}
  </div>
);

const Badge = ({ children, className }: { children: React.ReactNode; className?: string; key?: React.Key }) => (
  <span className={cn('px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider', className)}>
    {children}
  </span>
);

const Modal = ({ isOpen, onClose, title, children, maxWidth = "max-w-2xl" }: { isOpen: boolean; onClose: () => void; title: string; children: React.ReactNode; maxWidth?: string }) => (
  <AnimatePresence>
    {isOpen && (
      <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
        <motion.div
          initial={{ opacity: 0 }}
          animate={{ opacity: 1 }}
          exit={{ opacity: 0 }}
          onClick={onClose}
          className="absolute inset-0 bg-zinc-900/60 backdrop-blur-sm"
        />
        <motion.div
          initial={{ opacity: 0, scale: 0.95, y: 20 }}
          animate={{ opacity: 1, scale: 1, y: 0 }}
          exit={{ opacity: 0, scale: 0.95, y: 20 }}
          className={cn("relative w-full bg-white rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]", maxWidth)}
        >
          <div className="px-6 py-4 border-b border-zinc-100 flex items-center justify-between bg-white sticky top-0 z-10">
            <h3 className="text-xl font-bold tracking-tight">{title}</h3>
            <button 
              onClick={onClose}
              className="p-2 hover:bg-zinc-100 rounded-full transition-colors"
            >
              <Plus className="w-6 h-6 rotate-45 text-zinc-400" />
            </button>
          </div>
          <div className="p-6 overflow-y-auto">
            {children}
          </div>
        </motion.div>
      </div>
    )}
  </AnimatePresence>
);

// --- Main App ---

export default function App() {
  const [isQuotaExceeded, setIsQuotaExceeded] = useState(false);
  const [isSalesLoaded, setIsSalesLoaded] = useState(false);
  const [tripIdSearch, setTripIdSearch] = useState('');
  const [isSearchingTripId, setIsSearchingTripId] = useState(false);
  const [dashboardAggregates, setDashboardAggregates] = useState<any>(null);
  const [isLoadingAggregates, setIsLoadingAggregates] = useState(false);
  const [user, setUser] = useState<any>(null);
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [allUsers, setAllUsers] = useState<UserProfile[]>([]);
  const [leaderboardUsers, setLeaderboardUsers] = useState<UserProfile[]>([]);
  const [roles, setRoles] = useState<Role[]>([]);
  const [isRolesModalOpen, setIsRolesModalOpen] = useState(false);
  const [selectedRole, setSelectedRole] = useState<Role | null>(null);
  const [isRoleEditModalOpen, setIsRoleEditModalOpen] = useState(false);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [weeks, setWeeks] = useState<Week[]>([]);
  const [bdes, setBDES] = useState<BDE[]>([]);
  const [sales, setSales] = useState<Sales[]>([]);
  const [incentives, setIncentives] = useState<Incentive[]>([]);
  const [incentivePayments, setIncentivePayments] = useState<IncentivePayment[]>([]);
  const [matrixReports, setMatrixReports] = useState<MatrixReport[]>([]);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'employees' | 'weeks' | 'bdes' | 'sales' | 'weeklyOverview' | 'agentOverview' | 'bdeOverview' | 'statsOverview' | 'issueOverview' | 'userManagement' | 'training' | 'trainingOverview' | 'incentiveMaster' | 'passwordManager' | 'matrixMaster'>('dashboard');
  const [pagePasswords, setPagePasswords] = useState<Record<string, { password: string, isOpen: boolean }>>({});
  const [unlockedPages, setUnlockedPages] = useState<Set<string>>(new Set());
  const [isPasswordPromptOpen, setIsPasswordPromptOpen] = useState(false);
  const [pendingTab, setPendingTab] = useState<string | null>(null);
  const [passwordInput, setPasswordInput] = useState('');
  const [isMasterPasswordPromptOpen, setIsMasterPasswordPromptOpen] = useState(false);
  const [masterPasswordInput, setMasterPasswordInput] = useState('');
  const [passwordError, setPasswordError] = useState<string | null>(null);
  const [selectedBDEForOverview, setSelectedBDEForOverview] = useState<string>('');
  const [bdeOverviewTimeframe, setBdeOverviewTimeframe] = useState<'Daily' | 'Weekly' | '4-Weekly' | '8-Weekly' | 'Monthly'>('Weekly');
  const [selectedMonthForBDEOverview, setSelectedMonthForBDEOverview] = useState<string>('');
  const [agentOverviewTimeframe, setAgentOverviewTimeframe] = useState<'Daily' | 'Weekly' | '4-Weekly' | '8-Weekly' | 'Monthly'>('Monthly');
  const [selectedMonthForAgentOverview, setSelectedMonthForAgentOverview] = useState<string>('');
  const [bdeOverviewSummaryModal, setBdeOverviewSummaryModal] = useState<{ 
    isOpen: boolean, 
    type: 'sales' | 'margin' | 'expenses' | 'profit' | 'agent', 
    title: string,
    data?: any 
  }>({ isOpen: false, type: 'sales', title: '' });
  const [trainingAgentFilter, setTrainingAgentFilter] = useState<string>('');
  const [trainingModuleFilter, setTrainingModuleFilter] = useState<string>('');
  const [trainingStatusFilter, setTrainingStatusFilter] = useState<'Active' | 'Complete' | 'All'>('Active');
  const [trainingEmployeeStatusFilter, setTrainingEmployeeStatusFilter] = useState<'Active' | 'Deactive' | 'All'>('Active');
  const [isTrainingUpdateModalOpen, setIsTrainingUpdateModalOpen] = useState(false);
  const [selectedAgentForTraining, setSelectedAgentForTraining] = useState<Employee | null>(null);
  const [selectedModuleForTraining, setSelectedModuleForTraining] = useState<string | null>(null);
  const [selectedAgentOverview, setSelectedAgentOverview] = useState<string>('');
  const [selectedAgentOverviewDay, setSelectedAgentOverviewDay] = useState<number>(1);
  const [trainingStatusOverviewFilter, setTrainingStatusOverviewFilter] = useState<'All' | 'Pending' | 'Complete'>('All');
  const [isLoading, setIsLoading] = useState(true);
  const [isAwarding, setIsAwarding] = useState(false);
  const [isSavingEmployee, setIsSavingEmployee] = useState(false);
  const [isSavingWeek, setIsSavingWeek] = useState(false);
  const [isSavingBDE, setIsSavingBDE] = useState(false);
  const [isSavingSales, setIsSavingSales] = useState(false);
  const [isSavingUser, setIsSavingUser] = useState(false);
  const [isEditingEmployee, setIsEditingEmployee] = useState(false);
  const [isEditingWeek, setIsEditingWeek] = useState(false);
  const [isEditingBDE, setIsEditingBDE] = useState(false);
  const [isEditingSales, setIsEditingSales] = useState(false);
  const [isEditingUser, setIsEditingUser] = useState(false);
  const [isEmployeeModalOpen, setIsEmployeeModalOpen] = useState(false);
  const [isWeekModalOpen, setIsWeekModalOpen] = useState(false);
  const [isBDEModalOpen, setIsBDEModalOpen] = useState(false);
  const [isSalesModalOpen, setIsSalesModalOpen] = useState(false);
  const [isUserModalOpen, setIsUserModalOpen] = useState(false);
  const [isIncentiveModalOpen, setIsIncentiveModalOpen] = useState(false);
  const [isPaymentModalOpen, setIsPaymentModalOpen] = useState(false);
  const [isSavingIncentive, setIsSavingIncentive] = useState(false);
  const [isSavingPayment, setIsSavingPayment] = useState(false);
  const [isEditingIncentive, setIsEditingIncentive] = useState(false);
  const [isEditingPayment, setIsEditingPayment] = useState(false);
  const [selectedIncentive, setSelectedIncentive] = useState<Incentive | null>(null);
  const [selectedPayment, setSelectedPayment] = useState<IncentivePayment | null>(null);
  const [isOutstandingDuesModalOpen, setIsOutstandingDuesModalOpen] = useState(false);
  const [agentLedgerModal, setAgentLedgerModal] = useState<{ isOpen: boolean; employeeId: string | null }>({ isOpen: false, employeeId: null });
  const [isMatrixModalOpen, setIsMatrixModalOpen] = useState(false);
  const [matrixForm, setMatrixForm] = useState<Partial<MatrixReport>>({
    date: format(new Date(), 'yyyy-MM-dd'),
    performancePoints: 0,
    totalCalls: 0,
    totalTalktime: '',
    topTalktimes: [
      { tripId: '', phoneNo: '', talktime: '' },
      { tripId: '', phoneNo: '', talktime: '' },
      { tripId: '', phoneNo: '', talktime: '' },
      { tripId: '', phoneNo: '', talktime: '' },
    ],
    loginTime: '',
    logoutTime: '',
    breakTime: '',
    employeeId: '',
    employeeName: '',
  });
  const [isSavingMatrix, setIsSavingMatrix] = useState(false);
  const [matrixEditId, setMatrixEditId] = useState<string | null>(null);
  const [matrixView, setMatrixView] = useState<'reports' | 'trips'>('reports');
  const [matrixStartDate, setMatrixStartDate] = useState<string>(format(new Date(), 'yyyy-MM-dd'));
  const [matrixEndDate, setMatrixEndDate] = useState<string>(format(new Date(), 'yyyy-MM-dd'));
  const [matrixAgentFilter, setMatrixAgentFilter] = useState<string>('all');
  const [matrixBdeFilter, setMatrixBdeFilter] = useState<string>('all');
  const [incentiveForm, setIncentiveForm] = useState<Partial<Incentive>>({
    type: 'Daily',
    date: new Date().toISOString().split('T')[0],
    amount: 0,
    remarks: '',
    isEligible: false,
    isPaid: false
  });
  const [paymentForm, setPaymentForm] = useState<Partial<IncentivePayment>>({
    date: new Date().toISOString().split('T')[0],
    amount: 0,
    paymentMethod: 'Cash',
    remarks: ''
  });
  const [summaryDetailModal, setSummaryDetailModal] = useState<{
    isOpen: boolean;
    title: string;
    sales: Sales[];
  }>({
    isOpen: false,
    title: '',
    sales: []
  });
  const [bravoCalculationModal, setBravoCalculationModal] = useState<{
    isOpen: boolean;
    title: string;
    agentName: string;
    numWeeks: number;
    weeksData: {
      weekName: string;
      sales: number;
      target: number;
      percentage: number;
      tenureWeek: number;
      isReference?: boolean;
    }[];
    totalSales: number;
    totalTarget: number;
    achievement: number;
  }>({
    isOpen: false,
    title: '',
    agentName: '',
    numWeeks: 0,
    weeksData: [],
    totalSales: 0,
    totalTarget: 0,
    achievement: 0,
  });
  const isAdmin = useMemo(() => profile?.role === 'Admin' || (profile?.email === 'nsingla09@gmail.com' && user?.emailVerified), [profile, user]);
  const isManager = useMemo(() => profile?.role === 'Manager', [profile]);
  const isTrainer = useMemo(() => profile?.role?.toLowerCase() === 'trainer', [profile]);
  const isBDE = useMemo(() => profile?.role?.toLowerCase() === 'bde', [profile]);
  const isSalesAgent = useMemo(() => profile?.role?.toLowerCase() === 'sales agent', [profile]);
  const isAgent = useMemo(() => profile?.role?.toLowerCase() === 'agent' || isSalesAgent, [profile, isSalesAgent]);
  
  const hasTrainingOverviewPermission = useMemo(() => profile?.permissions ? (
    Array.isArray(profile.permissions) 
      ? profile.permissions.includes('trainingOverview') 
      : profile.permissions['trainingOverview'] === 'Complete'
  ) : false, [profile]);

  const hasIssueOverviewPermission = useMemo(() => profile?.permissions ? (
    Array.isArray(profile.permissions) 
      ? profile.permissions.includes('issueOverview') 
      : profile.permissions['issueOverview'] === 'Complete'
  ) : false, [profile]);

  const hasSalesComplete = useMemo(() => profile?.permissions ? (
    !Array.isArray(profile.permissions) && profile.permissions['sales'] === 'Complete'
  ) : false, [profile]);

  const hasAgentOverviewComplete = useMemo(() => profile?.permissions ? (
    !Array.isArray(profile.permissions) && profile.permissions['agentOverview'] === 'Complete'
  ) : false, [profile]);

  const isPrivileged = useMemo(() => isAdmin || isManager || isTrainer || isBDE || isSalesAgent || hasTrainingOverviewPermission || hasIssueOverviewPermission || hasSalesComplete || hasAgentOverviewComplete, 
    [isAdmin, isManager, isTrainer, isBDE, isSalesAgent, hasTrainingOverviewPermission, hasIssueOverviewPermission, hasSalesComplete, hasAgentOverviewComplete]);

  const [contributionCalculationModal, setContributionCalculationModal] = useState<{
    isOpen: boolean;
    title: string;
    numWeeks: number;
    totalContribution: number;
    weeksData: any[];
  }>({
    isOpen: false,
    title: '',
    numWeeks: 0,
    totalContribution: 0,
    weeksData: []
  });
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);
  const lastUserRef = useRef<any>(null);

  // Query for Employees
  const { data: employeesData, refetch: refetchEmployees } = useQuery({
    queryKey: ['employees', user?.uid, isPrivileged, profile?.bdeName],
    queryFn: async () => {
      if (!user) return [];
      const employeeFilters = [where('email', '==', user.email || '')];
      if (profile?.bdeName) employeeFilters.push(where('bde', '==', profile.bdeName));

      const employeesQuery = isPrivileged
        ? query(collection(db, 'employees'), limit(200))
        : (employeeFilters.length > 1 
            ? query(collection(db, 'employees'), or(...employeeFilters), limit(50))
            : query(collection(db, 'employees'), employeeFilters[0], limit(50)));
      
      const snap = await getDocs(employeesQuery);
      return snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Employee));
    },
    enabled: !!user && !!profile,
  });

  useEffect(() => {
    if (employeesData) setEmployees(employeesData);
  }, [employeesData]);

  // Query for Incentives
  const { data: incentivesData, refetch: refetchIncentives } = useQuery({
    queryKey: ['incentives', user?.uid, isPrivileged, profile?.uid],
    queryFn: async () => {
      if (!user) return [];
      const incentiveQuery = isPrivileged
        ? query(collection(db, 'incentives'), limit(100))
        : query(collection(db, 'incentives'), where('employeeId', '==', user.uid), limit(50));
      
      const snap = await getDocs(incentiveQuery);
      return snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Incentive));
    },
    enabled: !!user && !!profile,
  });

  useEffect(() => {
    if (incentivesData) setIncentives(incentivesData);
  }, [incentivesData]);

  // Query for Incentive Payments
  const { data: paymentsData, refetch: refetchPayments } = useQuery({
    queryKey: ['incentivePayments', user?.uid, isPrivileged, profile?.uid],
    queryFn: async () => {
      if (!user) return [];
      const paymentQuery = isPrivileged
        ? query(collection(db, 'incentivePayments'), limit(100))
        : query(collection(db, 'incentivePayments'), where('employeeId', '==', user.uid), limit(50));
      
      const snap = await getDocs(paymentQuery);
      return snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as IncentivePayment));
    },
    enabled: !!user && !!profile,
  });

  useEffect(() => {
    if (paymentsData) setIncentivePayments(paymentsData);
  }, [paymentsData]);

  // Query for Matrix Reports
  const { data: matrixData, refetch: refetchMatrix } = useQuery({
    queryKey: ['matrixReports', user?.uid, isPrivileged],
    queryFn: async () => {
      if (!user) return [];
      const matrixDaysAgo = subDays(new Date(), 14).toISOString().split('T')[0];
      const matrixQuery = query(collection(db, 'matrixReports'), where('date', '>=', matrixDaysAgo), limit(200));
      
      const snap = await getDocs(matrixQuery);
      return snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as MatrixReport));
    },
    enabled: !!user && !!profile,
  });

  useEffect(() => {
    if (matrixData) setMatrixReports(matrixData);
  }, [matrixData]);

  const refreshAllData = async () => {
    setIsRefreshingStatic(true);
    try {
      await Promise.all([
        refreshStaticData(),
        refetchEmployees(),
        refetchIncentives(),
        refetchPayments(),
        refetchMatrix(),
        queryClient.invalidateQueries({ queryKey: ['sales'] })
      ]);
    } finally {
      setIsRefreshingStatic(false);
    }
  };

  useEffect(() => {
    const handleClickOutside = (event: MouseEvent) => {
      if (menuRef.current && !menuRef.current.contains(event.target as Node)) {
        setIsMenuOpen(false);
      }
    };

    document.addEventListener('mousedown', handleClickOutside);
    return () => {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, []);
  const [searchQuery, setSearchQuery] = useState('');
  const [employeeSearchQuery, setEmployeeSearchQuery] = useState('');
  const [employeeStatusFilter, setEmployeeStatusFilter] = useState<'Active' | 'Deactive' | 'All'>('Active');
  const [employeeAssocBdeFilter, setEmployeeAssocBdeFilter] = useState<'All' | 'Yes' | 'No'>('All');
  const [salesAgentFilter, setSalesAgentFilter] = useState('');
  const [salesBdeFilter, setSalesBdeFilter] = useState('');
  const [salesWeekFilter, setSalesWeekFilter] = useState('');
  const [salesMonthFilter, setSalesMonthFilter] = useState('');
  const [salesSearchQuery, setSalesSearchQuery] = useState('');
  const [salesByFilter, setSalesByFilter] = useState<'All' | 'Inhouse' | 'Branch' | 'Franchisee Sales'>('All');
  const [weeklyOverviewSearch, setWeeklyOverviewSearch] = useState('');
  const [bdeOverviewSalesSearch, setBdeOverviewSalesSearch] = useState('');
  const [bdeOverviewWorkPendingFilter, setBdeOverviewWorkPendingFilter] = useState<string>('All');
  const [bdeOverviewStatusFilter, setBdeOverviewStatusFilter] = useState<string>('All');
  const [statsOverviewSourceFilter, setStatsOverviewSourceFilter] = useState<string>('All');
  const [statsOverviewStatusFilter, setStatsOverviewStatusFilter] = useState<string>('All');
  const [statsOverviewWeekFilter, setStatsOverviewWeekFilter] = useState<string>('All');
  const [statsOverviewMonthFilter, setStatsOverviewMonthFilter] = useState<string>('All');
  const [statsOverviewBdeFilter, setStatsOverviewBdeFilter] = useState<string>('All');
  const [issueOverviewWeekFilter, setIssueOverviewWeekFilter] = useState<string>('All');
  const [issueOverviewActiveTab, setIssueOverviewActiveTab] = useState<'All' | 'Hotel' | 'Flight' | 'Land' | 'HFL' | 'Work Pending' | 'Credit Note'>('All');
  const [remarksModalSale, setRemarksModalSale] = useState<Sales | null>(null);
  const [remarksLog, setRemarksLog] = useState<Remark[]>([]);
  const [newRemarkText, setNewRemarkText] = useState('');
  const [isAddingRemark, setIsAddingRemark] = useState(false);
  const [statsOverviewDestMetric, setStatsOverviewDestMetric] = useState<'totalValue' | 'salesCount' | 'totalPax' | 'avgSalesValue'>('totalValue');
  const [statsOverviewSourceMetric, setStatsOverviewSourceMetric] = useState<'count' | 'value'>('count');
  const [statsOverviewStatusMetric, setStatsOverviewStatusMetric] = useState<'count' | 'value'>('count');
  const [statsOverviewPackageAnalysisDimension, setStatsOverviewPackageAnalysisDimension] = useState<'Weekly' | 'Monthly' | 'BDE' | 'Destination' | 'Source'>('BDE');
  const [hasSetDefaultWeek, setHasSetDefaultWeek] = useState(false);
  const [hasSetDefaultMonth, setHasSetDefaultMonth] = useState(false);
  const [hasSetDefaultStatsMonth, setHasSetDefaultStatsMonth] = useState(false);
  const [selectedUser, setSelectedUser] = useState<UserProfile | null>(null);
  const [pointsToAward, setPointsToAward] = useState(10);
  const [reason, setReason] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [fatalError, setFatalError] = useState<any>(null);
  const [incentiveCalcSales, setIncentiveCalcSales] = useState<number>(0);
  const [incentiveCalcMonthlySales, setIncentiveCalcMonthlySales] = useState<number>(0);
  const [assumeWeeklyEligible, setAssumeWeeklyEligible] = useState<boolean>(false);
  const [userRoleFilter, setUserRoleFilter] = useState<string>('Agent');
  const [ledgerEmployeeFilter, setLedgerEmployeeFilter] = useState<string>('');

  const monthlyIncentiveData = useMemo(() => {
    if (!ledgerEmployeeFilter) return [];
    
    const empIncentives = incentives.filter(i => i.employeeId === ledgerEmployeeFilter);
    const empPayments = incentivePayments.filter(p => p.employeeId === ledgerEmployeeFilter);
    
    const months: Record<string, any> = {};
    
    empIncentives.forEach(i => {
      const monthKey = format(new Date(i.date), 'MMM yyyy');
      if (!months[monthKey]) {
        months[monthKey] = { 
          month: monthKey, 
          Daily: 0, 
          Weekly: 0, 
          Monthly: 0, 
          Quarterly: 0, 
          Annually: 0, 
          totalDue: 0, 
          totalPaid: 0,
          isEligible: false,
          isPaid: true,
          hasIncentives: true,
          typeStatuses: {
            Daily: { isEligible: false, isPaid: true, hasData: false },
            Weekly: { isEligible: false, isPaid: true, hasData: false },
            Monthly: { isEligible: false, isPaid: true, hasData: false },
            Quarterly: { isEligible: false, isPaid: true, hasData: false },
            Annually: { isEligible: false, isPaid: true, hasData: false },
          }
        };
      }
      months[monthKey][i.type] += i.amount;
      months[monthKey].totalDue += i.amount;
      if (i.isEligible) months[monthKey].isEligible = true;
      if (!i.isPaid) months[monthKey].isPaid = false;

      const ts = months[monthKey].typeStatuses[i.type];
      ts.hasData = true;
      if (i.isEligible) ts.isEligible = true;
      if (!i.isPaid) ts.isPaid = false;
    });
    
    empPayments.forEach(p => {
      const monthKey = format(new Date(p.date), 'MMM yyyy');
      if (!months[monthKey]) {
        months[monthKey] = { 
          month: monthKey, 
          Daily: 0, 
          Weekly: 0, 
          Monthly: 0, 
          Quarterly: 0, 
          Annually: 0, 
          totalDue: 0, 
          totalPaid: 0,
          isEligible: false,
          isPaid: false,
          hasIncentives: false,
          typeStatuses: {
            Daily: { isEligible: false, isPaid: false, hasData: false },
            Weekly: { isEligible: false, isPaid: false, hasData: false },
            Monthly: { isEligible: false, isPaid: false, hasData: false },
            Quarterly: { isEligible: false, isPaid: false, hasData: false },
            Annually: { isEligible: false, isPaid: false, hasData: false },
          }
        };
      }
      months[monthKey].totalPaid += p.amount;
    });
    
    return Object.values(months).map(m => ({
      ...m,
      isPaid: m.hasIncentives ? m.isPaid : false,
      isEligible: m.hasIncentives ? (m.isEligible && !m.isPaid) : false,
      typeStatuses: Object.fromEntries(
        Object.entries(m.typeStatuses).map(([type, status]: [string, any]) => [
          type,
          {
            ...status,
            isPaid: status.hasData ? status.isPaid : false,
            isEligible: status.hasData ? (status.isEligible && !status.isPaid) : false
          }
        ])
      )
    })).sort((a, b) => {
      const dateA = new Date(a.month);
      const dateB = new Date(b.month);
      return dateA.getTime() - dateB.getTime();
    }).slice(-6);
  }, [incentives, incentivePayments, ledgerEmployeeFilter]);

  const agentOverviewMonthlyIncentiveData = useMemo(() => {
    if (!selectedAgentOverview) return [];
    
    const empIncentives = incentives.filter(i => i.employeeId === selectedAgentOverview);
    const empPayments = incentivePayments.filter(p => p.employeeId === selectedAgentOverview);
    
    const months: Record<string, any> = {};
    
    empIncentives.forEach(i => {
      const monthKey = format(new Date(i.date), 'MMM yyyy');
      if (!months[monthKey]) {
        months[monthKey] = { 
          month: monthKey, 
          Daily: 0, 
          Weekly: 0, 
          Monthly: 0, 
          Quarterly: 0, 
          Annually: 0, 
          totalDue: 0, 
          totalPaid: 0,
          isEligible: false,
          isPaid: true,
          hasIncentives: true,
          typeStatuses: {
            Daily: { isEligible: false, isPaid: true, hasData: false },
            Weekly: { isEligible: false, isPaid: true, hasData: false },
            Monthly: { isEligible: false, isPaid: true, hasData: false },
            Quarterly: { isEligible: false, isPaid: true, hasData: false },
            Annually: { isEligible: false, isPaid: true, hasData: false },
          }
        };
      }
      months[monthKey][i.type] += i.amount;
      months[monthKey].totalDue += i.amount;
      if (i.isEligible) months[monthKey].isEligible = true;
      if (!i.isPaid) months[monthKey].isPaid = false;

      const ts = months[monthKey].typeStatuses[i.type];
      ts.hasData = true;
      if (i.isEligible) ts.isEligible = true;
      if (!i.isPaid) ts.isPaid = false;
    });
    
    empPayments.forEach(p => {
      const monthKey = format(new Date(p.date), 'MMM yyyy');
      if (!months[monthKey]) {
        months[monthKey] = { 
          month: monthKey, 
          Daily: 0, 
          Weekly: 0, 
          Monthly: 0, 
          Quarterly: 0, 
          Annually: 0, 
          totalDue: 0, 
          totalPaid: 0,
          isEligible: false,
          isPaid: false,
          hasIncentives: false,
          typeStatuses: {
            Daily: { isEligible: false, isPaid: false, hasData: false },
            Weekly: { isEligible: false, isPaid: false, hasData: false },
            Monthly: { isEligible: false, isPaid: false, hasData: false },
            Quarterly: { isEligible: false, isPaid: false, hasData: false },
            Annually: { isEligible: false, isPaid: false, hasData: false },
          }
        };
      }
      months[monthKey].totalPaid += p.amount;
    });
    
    return Object.values(months).map(m => ({
      ...m,
      isPaid: m.hasIncentives ? m.isPaid : false,
      isEligible: m.hasIncentives ? (m.isEligible && !m.isPaid) : false,
      typeStatuses: Object.fromEntries(
        Object.entries(m.typeStatuses).map(([type, status]: [string, any]) => [
          type,
          {
            ...status,
            isPaid: status.hasData ? status.isPaid : false,
            isEligible: status.hasData ? (status.isEligible && !status.isPaid) : false
          }
        ])
      )
    })).sort((a, b) => {
      const dateA = new Date(a.month);
      const dateB = new Date(b.month);
      return dateA.getTime() - dateB.getTime();
    }).slice(-6);
  }, [incentives, incentivePayments, selectedAgentOverview]);

  const modalMonthlyIncentiveData = useMemo(() => {
    if (!agentLedgerModal.employeeId) return [];
    
    const empIncentives = incentives.filter(i => i.employeeId === agentLedgerModal.employeeId);
    const empPayments = incentivePayments.filter(p => p.employeeId === agentLedgerModal.employeeId);
    
    const months: Record<string, any> = {};
    
    empIncentives.forEach(i => {
      const monthKey = format(new Date(i.date), 'MMM yyyy');
      if (!months[monthKey]) {
        months[monthKey] = { 
          month: monthKey, 
          Daily: 0, 
          Weekly: 0, 
          Monthly: 0, 
          Quarterly: 0, 
          Annually: 0, 
          totalDue: 0, 
          totalPaid: 0,
          isEligible: false,
          isPaid: true,
          hasIncentives: true,
          typeStatuses: {
            Daily: { isEligible: false, isPaid: true, hasData: false },
            Weekly: { isEligible: false, isPaid: true, hasData: false },
            Monthly: { isEligible: false, isPaid: true, hasData: false },
            Quarterly: { isEligible: false, isPaid: true, hasData: false },
            Annually: { isEligible: false, isPaid: true, hasData: false },
          }
        };
      }
      months[monthKey][i.type] += i.amount;
      months[monthKey].totalDue += i.amount;
      if (i.isEligible) months[monthKey].isEligible = true;
      if (!i.isPaid) months[monthKey].isPaid = false;

      const ts = months[monthKey].typeStatuses[i.type];
      ts.hasData = true;
      if (i.isEligible) ts.isEligible = true;
      if (!i.isPaid) ts.isPaid = false;
    });
    
    empPayments.forEach(p => {
      const monthKey = format(new Date(p.date), 'MMM yyyy');
      if (!months[monthKey]) {
        months[monthKey] = { 
          month: monthKey, 
          Daily: 0, 
          Weekly: 0, 
          Monthly: 0, 
          Quarterly: 0, 
          Annually: 0, 
          totalDue: 0, 
          totalPaid: 0,
          isEligible: false,
          isPaid: false,
          hasIncentives: false,
          typeStatuses: {
            Daily: { isEligible: false, isPaid: false, hasData: false },
            Weekly: { isEligible: false, isPaid: false, hasData: false },
            Monthly: { isEligible: false, isPaid: false, hasData: false },
            Quarterly: { isEligible: false, isPaid: false, hasData: false },
            Annually: { isEligible: false, isPaid: false, hasData: false },
          }
        };
      }
      months[monthKey].totalPaid += p.amount;
    });
    
    return Object.values(months).map(m => ({
      ...m,
      isPaid: m.hasIncentives ? m.isPaid : false,
      isEligible: m.hasIncentives ? (m.isEligible && !m.isPaid) : false,
      typeStatuses: Object.fromEntries(
        Object.entries(m.typeStatuses).map(([type, status]: [string, any]) => [
          type,
          {
            ...status,
            isPaid: status.hasData ? status.isPaid : false,
            isEligible: status.hasData ? (status.isEligible && !status.isPaid) : false
          }
        ])
      )
    })).sort((a, b) => {
      const dateA = new Date(a.month);
      const dateB = new Date(b.month);
      return dateA.getTime() - dateB.getTime();
    }).slice(-6);
  }, [incentives, incentivePayments, agentLedgerModal.employeeId]);

  const allEmployeesIncentiveSummary = useMemo(() => {
    if (ledgerEmployeeFilter !== 'all') return [];
    
    const summary: Record<string, any> = {};
    
    employees.forEach(emp => {
      summary[emp.id] = {
        employeeId: emp.id,
        employeeName: emp.name,
        Daily: { due: 0, eligible: 0 },
        Weekly: { due: 0, eligible: 0 },
        Monthly: { due: 0, eligible: 0 },
        Quarterly: { due: 0, eligible: 0 },
        Annually: { due: 0, eligible: 0 },
        totalDue: { due: 0, eligible: 0 },
        totalPaid: 0,
        totalIncentives: 0,
        balance: 0
      };
    });
    
    incentives.forEach(i => {
      if (summary[i.employeeId]) {
        summary[i.employeeId].totalIncentives += i.amount;
        if (!i.isPaid) {
          if (i.isEligible) {
            summary[i.employeeId][i.type].eligible += i.amount;
            summary[i.employeeId].totalDue.eligible += i.amount;
          } else {
            summary[i.employeeId][i.type].due += i.amount;
            summary[i.employeeId].totalDue.due += i.amount;
          }
        }
      }
    });
    
    incentivePayments.forEach(p => {
      if (summary[p.employeeId]) {
        summary[p.employeeId].totalPaid += p.amount;
      }
    });
    
    return Object.values(summary).map((s: any) => ({
      ...s,
      balance: s.totalIncentives - s.totalPaid
    })).sort((a, b) => a.employeeName.localeCompare(b.employeeName));
  }, [employees, incentives, incentivePayments, ledgerEmployeeFilter]);

  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    id: string;
    type: 'employee' | 'week' | 'bde' | 'sales' | 'user' | 'trainingMaterials' | 'role';
    name: string;
    passwordInput?: string;
    error?: string | null;
  } | null>(null);
  const unsubsRef = useRef<(() => void)[]>([]);
  const hasRedirectedRef = useRef(false);

  const [isRefreshingStatic, setIsRefreshingStatic] = useState(false);
  const queryClient = useQueryClient();

  const refreshStaticData = async () => {
    if (isRefreshingStatic) return;
    setIsRefreshingStatic(true);
    try {
      // Try loading Data Bundle with cache busting
      const response = await fetch(`/api/bundle?t=${Date.now()}`);
      if (response.ok) {
        const bundleData = await response.arrayBuffer();
        await loadBundle(db, bundleData);
        
        const rolesSnap = await getDocs(await namedQuery(db, 'roles-query'));
        setRoles(rolesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Role)));
        
        const trainingSnap = await getDocs(await namedQuery(db, 'training-query'));
        setTrainingMaterials(trainingSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as TrainingMaterial)));
        
        const weeksSnap = await getDocs(await namedQuery(db, 'weeks-query'));
        setWeeks(weeksSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Week)));
        
        const bdesSnap = await getDocs(await namedQuery(db, 'bdes-query'));
        setBDES(bdesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as BDE)));
        
        const passwordsSnap = await getDocs(await namedQuery(db, 'passwords-query'));
        const passwords: Record<string, { password: string, isOpen: boolean }> = {};
        passwordsSnap.docs.forEach(doc => {
          passwords[doc.id] = doc.data() as { password: string, isOpen: boolean };
        });
        setPagePasswords(passwords);
        
        console.log("Static data refreshed successfully from Data Bundle");
        return;
      }
    } catch (err) {
      console.warn("Failed to refresh via bundle, falling back to direct reads:", err);
    }

    try {
      // Fallback: Fetch Roles
      const rolesSnap = await getDocs(query(collection(db, 'roles'), limit(50)));
      setRoles(rolesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Role)));

      // Fetch Training Materials
      const trainingSnap = await getDocs(query(collection(db, 'trainingMaterials'), limit(100)));
      setTrainingMaterials(trainingSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as TrainingMaterial)));

      // Fetch Weeks
      const weeksSnap = await getDocs(query(collection(db, 'weeks'), limit(100)));
      setWeeks(weeksSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Week)));

      // Fetch BDEs
      const bdesSnap = await getDocs(query(collection(db, 'bdes'), limit(100)));
      setBDES(bdesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as BDE)));

      // Fetch Page Passwords
      const passwordsSnap = await getDocs(collection(db, 'pagePasswords'));
      const passwords: Record<string, { password: string, isOpen: boolean }> = {};
      passwordsSnap.docs.forEach(doc => {
        passwords[doc.id] = doc.data() as { password: string, isOpen: boolean };
      });
      setPagePasswords(passwords);

      console.log("Static data refreshed successfully");
    } catch (err) {
      console.error("Error refreshing static data:", err);
    } finally {
      setIsRefreshingStatic(false);
    }
  };

  const cleanupListeners = () => {
    unsubsRef.current.forEach(unsub => {
      try {
        unsub();
      } catch (e) {
        console.error("Error cleaning up listener:", e);
      }
    });
    unsubsRef.current = [];
  };

  // Employee Form State
  const [employeeForm, setEmployeeForm] = useState<Partial<Employee>>({
    status: 'Active',
    position: 'Sales Agent',
    joiningWeek: 'Week 1',
    bde: 'None',
    email: '',
    isAssocBDE: false,
    week1Target: 25,
    week2Target: 25,
    week3Target: 50,
    week4Target: 50,
    week5Target: 50,
    week6Target: 50,
    completedTrainings: [],
  });

  const [trainingMaterials, setTrainingMaterials] = useState<TrainingMaterial[]>([]);
  const [trainingForm, setTrainingForm] = useState<Partial<TrainingMaterial>>({
    module: ['Sales'],
    points: 1,
    day: 1,
    time: 15,
    isActive: true,
    link: '',
    priority: 0
  });
  const [isTrainingModalOpen, setIsTrainingModalOpen] = useState(false);

  const [weekForm, setWeekForm] = useState<Partial<Week>>({
    weekName: 'Week 1',
    month: new Date().toLocaleString('default', { month: 'long', year: 'numeric' }),
  });

  const [bdeForm, setBDEForm] = useState<Partial<BDE>>({
    name: '',
    phone: '',
    email: ''
  });

  const [salesForm, setSalesForm] = useState<Partial<Sales>>({
    week: '',
    date: new Date().toISOString().split('T')[0],
    guestName: '',
    agent: '',
    agentEmail: '',
    agentPercentage: 100,
    bde: '',
    associateBde: '',
    assocBdePercentage: 0,
    destination: '',
    tripId: '',
    tripDate: '',
    packageValue: undefined,
    lessThan10Percent20k: '',
    advanceCN: '',
    salesBy: 'Inhouse',
    ppMargin: undefined,
    noOfPax: undefined,
    totalMargin: undefined,
    flight: undefined,
    source: '',
    converted: 'Yes',
    tasksPending: '',
    pHotel: '',
    pFlight: '',
    land: '',
    hflIssue: '',
    workPending: '',
    remarks: '',
  });

  const [userForm, setUserForm] = useState<Partial<UserProfile>>({
    role: 'Agent',
    permissions: ['dashboard'],
  });

  const isEmployee = useMemo(() => {
    return isAgent || employees.some(emp => emp.email && emp.email.toLowerCase() === profile?.email?.toLowerCase());
  }, [employees, profile, isAgent]);

  const bdeUser = useMemo(() => {
    return bdes.find(b => b.email && b.email.toLowerCase() === profile?.email?.toLowerCase());
  }, [bdes, profile]);

  const currentEmployee = useMemo(() => {
    return employees.find(e => e.email?.toLowerCase() === profile?.email?.toLowerCase());
  }, [employees, profile]);

  const isBDEUser = isBDE || !!bdeUser;
  const currentBDEName = bdeUser?.name || '';
  const currentEmployeeName = currentEmployee?.name || '';

  const findEmployeeBySale = useCallback((name: string, email?: string) => {
    if (!name && !email) return null;
    return employees.find(e => 
      (email && e.email?.toLowerCase() === email.toLowerCase()) ||
      (name && e.name.toLowerCase().trim() === name.toLowerCase().trim())
    );
  }, [employees]);

  const findBDEBySale = useCallback((name: string, email?: string) => {
    if (!name && !email) return null;
    return bdes.find(b => 
      (email && b.email?.toLowerCase() === email.toLowerCase()) ||
      (name && b.name.toLowerCase().trim() === name.toLowerCase().trim())
    );
  }, [bdes]);

  const getAgentCreditForSale = useCallback((s: Sales, agent: Employee, field: 'packageValue' | 'totalMargin' = 'packageValue') => {
    let credit = 0;
    const agentName = agent.name.toLowerCase().trim();
    const agentEmail = agent.email?.toLowerCase().trim();

    const saleAgentName = s.agent?.toLowerCase().trim();
    const saleAgentEmail = s.agentEmail?.toLowerCase().trim();
    const saleAssocBdeName = s.associateBde?.toLowerCase().trim();
    const saleAssocBdeEmail = s.associateBdeEmail?.toLowerCase().trim();

    const value = s[field] || 0;

    if ((saleAgentName && saleAgentName === agentName) || (agentEmail && saleAgentEmail === agentEmail)) {
      credit += value * ((s.agentPercentage ?? 100) / 100);
    }
    if ((saleAssocBdeName && saleAssocBdeName === agentName) || (agentEmail && saleAssocBdeEmail === agentEmail)) {
      credit += value * ((s.assocBdePercentage ?? 0) / 100);
    }
    return credit;
  }, []);

  const getPermissionLevel = (permission: string): 'None' | 'Limited' | 'Complete' => {
    if (isAdmin) return 'Complete';
    
    // Check custom permissions first
    if (profile?.permissions) {
      if (Array.isArray(profile.permissions)) {
        if (profile.permissions.includes(permission)) return 'Complete';
      } else if (profile.permissions[permission]) {
        return profile.permissions[permission] as 'None' | 'Limited' | 'Complete';
      }
    }

    // Legacy/Default logic
    if ((permission === 'trainingOverview' || permission === 'issueOverview') && (isManager || isTrainer)) return 'Complete';
    
    // Sales Agent specific overrides
    if (isSalesAgent && (permission === 'sales' || permission === 'agentOverview' || permission === 'matrixMaster')) return 'Complete';

    if (permission === 'dashboard' || permission === 'agentOverview' || permission === 'sales' || permission === 'bdeOverview' || permission === 'issueOverview' || permission === 'matrixMaster' || permission === 'training') {
      if (isEmployee || isBDEUser || isBDE || isAgent) return 'Limited';
    }
    
    return 'None';
  };

  const hasPermission = (permission: string) => {
    return getPermissionLevel(permission) !== 'None';
  };

  useEffect(() => {
    if (isEmployee && employees.length > 0) {
      const currentEmp = employees.find(e => e.email?.toLowerCase() === profile?.email?.toLowerCase());
      if (currentEmp) {
        if (!selectedAgentOverview && currentEmp.id) {
          setSelectedAgentOverview(currentEmp.id);
        }
        if (!salesAgentFilter && currentEmp.name) {
          setSalesAgentFilter(currentEmp.name);
        }
      }
    }
    if (isAdmin && !selectedBDEForOverview) {
      setSelectedBDEForOverview('All');
    } else if (isBDEUser && bdes.length > 0) {
      const currentBDE = bdes.find(b => b.email?.toLowerCase() === profile?.email?.toLowerCase());
      if (currentBDE && !selectedBDEForOverview) {
        setSelectedBDEForOverview(currentBDE.name);
      }
    }
  }, [isEmployee, isBDEUser, isAdmin, employees, bdes, profile, selectedAgentOverview, salesAgentFilter, selectedBDEForOverview]);

  const filteredEmployees = useMemo(() => {
    const level = getPermissionLevel('employees');
    if (level === 'None') return [];

    return employees.filter(emp => {
      const matchesSearch = 
        emp.name.toLowerCase().includes(employeeSearchQuery.toLowerCase()) || 
        emp.employeeCode.toLowerCase().includes(employeeSearchQuery.toLowerCase()) ||
        (emp.email && emp.email.toLowerCase().includes(employeeSearchQuery.toLowerCase()));
      
      const matchesStatus = 
        employeeStatusFilter === 'All' || 
        emp.status === employeeStatusFilter;
      
      const canSeeEmployee = 
        level === 'Complete' || 
        !emp.email || 
        emp.email.toLowerCase() === profile?.email?.toLowerCase();
      
      const matchesAssocBde = 
        employeeAssocBdeFilter === 'All' || 
        (employeeAssocBdeFilter === 'Yes' ? emp.isAssocBDE === true : emp.isAssocBDE !== true);
      
      return matchesSearch && matchesStatus && canSeeEmployee && matchesAssocBde;
    }).sort((a, b) => a.name.localeCompare(b.name));
  }, [employees, employeeSearchQuery, employeeStatusFilter, employeeAssocBdeFilter, profile, getPermissionLevel]);

  const fetchWithTimeout = async <T,>(promise: Promise<T>, timeoutMs: number = 30000, errorMsg: string = 'Request Timeout'): Promise<T> => {
    return Promise.race([
      promise,
      new Promise<T>((_, reject) => setTimeout(() => reject(new Error(errorMsg)), timeoutMs))
    ]);
  };

  // --- Auth & Profile Sync ---
  useEffect(() => {
    console.log("Auth sync useEffect started");
    const timer = setTimeout(() => {
      console.log("Auth sync timeout, forcing loading state to false");
      setIsLoading(false);
    }, 15000);

    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      clearTimeout(timer);
      console.log("onAuthStateChanged called", currentUser);
      if (currentUser?.uid === lastUserRef.current?.uid) {
        setIsLoading(false);
        return;
      }
      lastUserRef.current = currentUser;
      cleanupListeners();
      try {
        setUser(currentUser);
        console.log("Auth state changed:", currentUser?.email, currentUser?.uid);
        if (currentUser) {
          const userRef = doc(db, 'users', currentUser.uid);
          console.log("Fetching user profile for:", currentUser.uid);
          let userSnap;
          try {
            userSnap = await fetchWithTimeout(getDoc(userRef), 30000, 'Profile Fetch Timeout');
            console.log("User profile fetch completed");
          } catch (err) {
            console.error("Error fetching user profile:", err);
            const errorString = err instanceof Error ? err.message : String(err);
            if (errorString.includes('Quota exceeded') || errorString.includes('quota limit exceeded')) {
              try {
                userSnap = await getDocFromCache(userRef);
                console.log("Using cached profile due to quota error");
              } catch (cacheErr) {
                throw err;
              }
            } else {
              throw err;
            }
          }
          
          let currentProfile: UserProfile;
          if (!userSnap || !userSnap.exists()) {
            console.log("User profile not found by UID, checking by email...");
            const q = query(collection(db, 'users'), where('email', '==', currentUser.email), limit(1));
            let querySnap;
            try {
              querySnap = await fetchWithTimeout(getDocs(q), 30000, 'Email Profile Fetch Timeout');
            } catch (err) {
              const errorString = err instanceof Error ? err.message : String(err);
              if (errorString.includes('Quota exceeded') || errorString.includes('quota limit exceeded')) {
                try {
                  querySnap = await getDocsFromCache(q);
                  console.log("Using cached email search due to quota error");
                } catch (cacheErr) {
                  throw err;
                }
              } else {
                throw err;
              }
            }
            
            if (querySnap && !querySnap.empty) {
              console.log("Found pre-provisioned profile by email, migrating to UID...");
              const preDoc = querySnap.docs[0];
              const preData = preDoc.data() as UserProfile;
              
              const defaultPermissions: Record<string, 'Limited' | 'Complete'> = {
                dashboard: 'Complete',
                sales: 'Limited',
                training: 'Limited',
                bdeOverview: 'Limited'
              };
              
              currentProfile = {
                ...preData,
                uid: currentUser.uid,
                displayName: currentUser.displayName || preData.displayName,
                photoURL: currentUser.photoURL || preData.photoURL,
                pointsBalance: preData.pointsBalance ?? 100,
                role: preData.role || 'Agent',
                permissions: preData.permissions || defaultPermissions,
              };
              
              await setDoc(userRef, currentProfile);
              if (preDoc.id !== currentUser.uid) {
                await deleteDoc(preDoc.ref);
              }
              setProfile(currentProfile);
            } else {
              console.log("No profile found, creating new profile...");
              const isFirstUser = currentUser.email === 'nsingla09@gmail.com' && currentUser.emailVerified;
              const adminPermissions: Record<string, 'Limited' | 'Complete'> = {};
              allPermissions.forEach(p => { adminPermissions[p.id] = 'Complete'; });
              
              const defaultPermissions: Record<string, 'Limited' | 'Complete'> = {
                dashboard: 'Complete',
                sales: 'Limited',
                training: 'Limited',
                bdeOverview: 'Limited'
              };

              currentProfile = {
                uid: currentUser.uid,
                email: currentUser.email || '',
                displayName: currentUser.displayName || 'Anonymous User',
                photoURL: currentUser.photoURL || '',
                pointsBalance: 100, // Starting points
                role: isFirstUser ? 'Admin' : 'Agent',
                permissions: isFirstUser 
                  ? adminPermissions
                  : defaultPermissions,
              };
              await setDoc(userRef, currentProfile);
              setProfile(currentProfile);
            }
            
            if (!(currentProfile.permissions?.['dashboard']) && currentProfile.permissions?.['employees']) {
              setActiveTab('employees');
            }
          } else {
            console.log("User profile found:", userSnap.data());
            currentProfile = userSnap.data() as UserProfile;
            
            // Ensure role and permissions exist
            const defaultKeys = ['dashboard', 'sales', 'training', 'bdeOverview', 'agentOverview', 'issueOverview', 'matrixMaster'];
            let needsUpdate = false;
            
            if (!currentProfile.role) {
              const role = (currentProfile.email === 'nsingla09@gmail.com' && currentUser.emailVerified) ? 'Admin' : 'Agent';
              currentProfile.role = role;
              needsUpdate = true;
            }

            if (!currentProfile.permissions || typeof currentProfile.permissions !== 'object' || Array.isArray(currentProfile.permissions)) {
              const adminPermissions: Record<string, 'Limited' | 'Complete'> = {};
              allPermissions.forEach(p => { adminPermissions[p.id] = 'Complete'; });
              
              const defaultPermissions: Record<string, 'Limited' | 'Complete'> = {
                dashboard: 'Complete',
                sales: 'Limited',
                training: 'Limited',
                bdeOverview: 'Limited',
                agentOverview: 'Limited',
                issueOverview: 'Limited',
                matrixMaster: 'Limited'
              };

              currentProfile.permissions = currentProfile.role === 'Admin' ? adminPermissions : defaultPermissions;
              needsUpdate = true;
            } else {
              // Add missing default permissions
              defaultKeys.forEach(key => {
                if (!(key in currentProfile.permissions)) {
                  currentProfile.permissions[key] = 'Limited';
                  needsUpdate = true;
                }
              });
            }

            if (needsUpdate) {
              try {
                await updateDoc(userRef, { 
                  role: currentProfile.role, 
                  permissions: currentProfile.permissions 
                });
              } catch (err) {
                console.warn("Could not update profile in DB", err);
              }
            }
            setProfile(currentProfile);
            if (!(currentProfile.permissions?.['dashboard']) && currentProfile.permissions?.['employees']) {
              setActiveTab('employees');
            }
          }

          // Check if user is a BDE to get their name for queries
          let bdeName = currentProfile.bdeName || '';
          if (!bdeName) {
            try {
              const bdeQuery = query(collection(db, 'bdes'), where('email', '==', currentUser.email));
              let bdeSnap;
              try {
                bdeSnap = await fetchWithTimeout(getDocs(bdeQuery), 30000, 'BDE Status Fetch Timeout');
              } catch (err) {
                const errorString = err instanceof Error ? err.message : String(err);
                if (errorString.includes('Quota exceeded') || errorString.includes('quota limit exceeded')) {
                  bdeSnap = await getDocsFromCache(bdeQuery);
                } else {
                  throw err;
                }
              }

              if (bdeSnap && !bdeSnap.empty) {
                bdeName = bdeSnap.docs[0].data().name;
                currentProfile.bdeName = bdeName;
                
                // Assign BDE role if not already privileged
                if (!currentProfile.role || currentProfile.role === 'Agent' || currentProfile.role === 'Sales Agent') {
                  currentProfile.role = 'BDE';
                  currentProfile.permissions = {
                    ...(currentProfile.permissions || {}),
                    dashboard: 'Complete',
                    sales: 'Limited',
                    agentOverview: 'Limited',
                    bdeOverview: 'Limited',
                    training: 'Limited'
                  };
                }
                
                try {
                  await updateDoc(userRef, { 
                    bdeName,
                    role: currentProfile.role, 
                    permissions: currentProfile.permissions 
                  });
                  setProfile({ ...currentProfile });
                } catch (err) {
                  console.warn("Failed to update BDE info in DB", err);
                }
              }
            } catch (err) {
              console.error("Error checking BDE status:", err);
            }
          }

          let employeeName = currentProfile.employeeName || '';
          let employeeId = currentProfile.employeeId || '';
          if (!employeeName || !employeeId) {
            try {
              const empQuery = query(collection(db, 'employees'), where('email', '==', currentUser.email));
              let empSnap;
              try {
                empSnap = await fetchWithTimeout(getDocs(empQuery), 30000, 'Employee Status Fetch Timeout');
              } catch (err) {
                const errorString = err instanceof Error ? err.message : String(err);
                if (errorString.includes('Quota exceeded') || errorString.includes('quota limit exceeded')) {
                  empSnap = await getDocsFromCache(empQuery);
                } else {
                  throw err;
                }
              }

              if (empSnap && !empSnap.empty) {
                employeeName = empSnap.docs[0].data().name;
                employeeId = empSnap.docs[0].id;
                currentProfile.employeeName = employeeName;
                currentProfile.employeeId = employeeId;
                try {
                  await updateDoc(userRef, { employeeName, employeeId });
                  setProfile({ ...currentProfile });
                } catch (err) {
                  console.warn("Could not update employee info in DB", err);
                }
              }
            } catch (err) {
              console.error("Error checking employee status:", err);
            }
          }

          const isUserAdmin = currentProfile.role?.toLowerCase() === 'admin' || (currentProfile.email === 'nsingla09@gmail.com' && currentUser.emailVerified);
          const isUserManager = currentProfile.role?.toLowerCase() === 'manager';
          const isUserTrainer = currentProfile.role?.toLowerCase() === 'trainer';
          const isUserBDE = currentProfile.role?.toLowerCase() === 'bde';
          const isUserSalesAgent = currentProfile.role?.toLowerCase() === 'sales agent';
          const hasTrainingOverviewPermission = currentProfile.permissions ? (
            Array.isArray(currentProfile.permissions)
              ? currentProfile.permissions.includes('trainingOverview')
              : currentProfile.permissions['trainingOverview'] === 'Complete'
          ) : false;
          const hasIssueOverviewPermission = currentProfile.permissions ? (
            Array.isArray(currentProfile.permissions)
              ? currentProfile.permissions.includes('issueOverview')
              : currentProfile.permissions['issueOverview'] === 'Complete'
          ) : false;
          const hasSalesComplete = currentProfile.permissions ? (
            !Array.isArray(currentProfile.permissions) && currentProfile.permissions['sales'] === 'Complete'
          ) : false;
          const hasAgentOverviewComplete = currentProfile.permissions ? (
            !Array.isArray(currentProfile.permissions) && currentProfile.permissions['agentOverview'] === 'Complete'
          ) : false;
          const isUserPrivileged = isUserAdmin || isUserManager || isUserTrainer || isUserBDE || isUserSalesAgent || hasTrainingOverviewPermission || hasIssueOverviewPermission || hasSalesComplete || hasAgentOverviewComplete;

          console.log("Starting Firestore listeners...");
          const unsubProfile = onSnapshot(userRef, (doc) => {
            if (doc.exists()) {
              setProfile(doc.data() as UserProfile);
            }
          }, (err) => {
            console.error("Profile listener error:", err);
            // Only set fatal error if we don't have a profile yet and user is still logged in
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, `users/${currentUser.uid}`);
            }
          });
          unsubsRef.current.push(unsubProfile);

          // One-time fetch for static/infrequently updated data
          const fetchStaticData = async () => {
            try {
              // Try loading Data Bundle first to save read costs
              const response = await fetch('/api/bundle');
              if (response.ok) {
                const bundleData = await response.arrayBuffer();
                await loadBundle(db, bundleData);
                
                // Load from named queries in the bundle
                const rolesSnap = await getDocs(await namedQuery(db, 'roles-query'));
                setRoles(rolesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Role)));
                
                const trainingSnap = await getDocs(await namedQuery(db, 'training-query'));
                setTrainingMaterials(trainingSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as TrainingMaterial)));
                
                const weeksSnap = await getDocs(await namedQuery(db, 'weeks-query'));
                setWeeks(weeksSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Week)));
                
                const bdesSnap = await getDocs(await namedQuery(db, 'bdes-query'));
                setBDES(bdesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as BDE)));
                
                const passwordsSnap = await getDocs(await namedQuery(db, 'passwords-query'));
                const passwords: Record<string, { password: string, isOpen: boolean }> = {};
                passwordsSnap.docs.forEach(doc => {
                  passwords[doc.id] = doc.data() as { password: string, isOpen: boolean };
                });
                setPagePasswords(passwords);
                
                console.log("Static data loaded successfully from Data Bundle");
                return; // Exit if bundle loaded successfully
              }
            } catch (err) {
              console.warn("Failed to load Data Bundle, falling back to direct reads:", err);
            }

            try {
              // Fallback: Fetch Roles
              const rolesSnap = await getDocs(query(collection(db, 'roles'), limit(50)));
              const rolesList = rolesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Role));
              setRoles(rolesList);
              
              // Initialize default roles if none exist
              if (rolesList.length === 0 && isUserAdmin) {
                console.log("Initializing default roles...");
                const defaultRoles = [
                  { name: 'Admin', defaultPage: 'dashboard', permissions: Object.fromEntries(allPermissions.map(p => [p.id, 'Complete'])) },
                  { name: 'Manager', defaultPage: 'dashboard', permissions: Object.fromEntries(['dashboard', 'employees', 'weeks', 'bdes', 'sales', 'weeklyOverview', 'agentOverview', 'training', 'trainingOverview', 'issueOverview'].map(p => [p, 'Complete'])) },
                  { name: 'Trainer', defaultPage: 'training', permissions: Object.fromEntries(['dashboard', 'training', 'trainingOverview'].map(p => [p, 'Complete'])) },
                  { name: 'BDE', defaultPage: 'bdeOverview', permissions: Object.fromEntries(['dashboard', 'sales', 'agentOverview', 'bdeOverview', 'training'].map(p => [p, 'Complete'])) },
                  { name: 'Agent', defaultPage: 'dashboard', permissions: Object.fromEntries(['dashboard', 'sales', 'training', 'trainingOverview'].map(p => [p, 'Complete'])) },
                  { name: 'Operations', defaultPage: 'issueOverview', permissions: Object.fromEntries(['dashboard', 'issueOverview'].map(p => [p, 'Complete'])) },
                  { name: 'Extra Team', defaultPage: 'dashboard', permissions: Object.fromEntries(['dashboard', 'training', 'trainingOverview'].map(p => [p, 'Complete'])) },
                  { name: 'Trainee', defaultPage: 'training', permissions: Object.fromEntries(['dashboard', 'training', 'trainingOverview'].map(p => [p, 'Complete'])) },
                ];
                
                for (const role of defaultRoles) {
                  await addDoc(collection(db, 'roles'), {
                    ...role,
                    createdAt: serverTimestamp(),
                  });
                }
                const updatedRolesSnap = await getDocs(query(collection(db, 'roles'), limit(50)));
                setRoles(updatedRolesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Role)));
              }

              // Fetch Training Materials
              const trainingSnap = await getDocs(query(collection(db, 'trainingMaterials'), limit(100)));
              setTrainingMaterials(trainingSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as TrainingMaterial)));

              // Fetch Weeks
              const weeksSnap = await getDocs(query(collection(db, 'weeks'), limit(100)));
              setWeeks(weeksSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Week)));

              // Fetch BDEs
              const bdesSnap = await getDocs(query(collection(db, 'bdes'), limit(100)));
              setBDES(bdesSnap.docs.map(doc => ({ id: doc.id, ...doc.data() } as BDE)));

              // Fetch All Users (for management)
              if (isUserAdmin || isUserManager) {
                const usersSnap = await getDocs(query(collection(db, 'users'), limit(100)));
                setAllUsers(usersSnap.docs.map(doc => doc.data() as UserProfile));
              }

              // Fetch Page Passwords
              const passwordsSnap = await getDocs(collection(db, 'pagePasswords'));
              const passwords: Record<string, { password: string, isOpen: boolean }> = {};
              passwordsSnap.docs.forEach(doc => {
                passwords[doc.id] = doc.data() as { password: string, isOpen: boolean };
              });
              setPagePasswords(passwords);

            } catch (err) {
              console.error("Error fetching static data:", err);
            }
          };

          fetchStaticData();

          // Listen for leaderboard (Top 10)
          const leaderboardQuery = query(collection(db, 'users'), orderBy('pointsBalance', 'desc'), limit(10));
          const unsubLeaderboard = onSnapshot(leaderboardQuery, (snap) => {
            const users = snap.docs.map(doc => doc.data() as UserProfile);
            setLeaderboardUsers(users);
          }, (err) => {
            console.error("Leaderboard listener error:", err);
          });
          unsubsRef.current.push(unsubLeaderboard);

          // Listen for transactions
          const transQuery = isUserPrivileged 
            ? query(collection(db, 'transactions'), orderBy('createdAt', 'desc'), limit(20))
            : query(collection(db, 'transactions'), 
                or(
                  where('fromUid', '==', currentUser.uid),
                  where('toUid', '==', currentUser.uid)
                ),
                limit(20));

          const unsubTrans = onSnapshot(transQuery, (snap) => {
            const trans = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Transaction));
            // Sort client-side for non-privileged users since we removed orderBy to avoid index requirement
            if (!isUserPrivileged) {
              trans.sort((a, b) => {
                const dateA = a.createdAt?.toMillis ? a.createdAt.toMillis() : 0;
                const dateB = b.createdAt?.toMillis ? b.createdAt.toMillis() : 0;
                return dateB - dateA;
              });
            }
            setTransactions(trans);
          }, (err) => {
            console.error("Transactions listener error:", err);
            if (err.message.includes('Quota exceeded') || err.message.includes('quota limit exceeded')) {
              setIsQuotaExceeded(true);
            }
            if (auth.currentUser && !err.message.includes('Quota exceeded')) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'transactions');
            }
          });
          unsubsRef.current.push(unsubTrans);

          if (isUserAdmin) {
            console.log("User is admin, starting admin listeners...");
            // Admin-only listeners (none currently, but kept for structure)
          }
        } else {
          setProfile(null);
        }
      } catch (err) {
        console.error("Auth sync error:", err);
        const errorString = err instanceof Error ? err.message : String(err);
        if (errorString.includes('Quota exceeded') || errorString.includes('quota limit exceeded')) {
          setFatalError({
            message: "Firestore daily quota exceeded. The free tier limit has been reached for today. The quota will reset tomorrow. You can check detailed quota info at https://firebase.google.com/pricing#cloud-firestore"
          });
        } else {
          setFatalError(err);
        }
      } finally {
        console.log("Auth sync finally block, setting loading to false");
        setIsLoading(false);
      }
    });

    return () => {
      clearTimeout(timer);
      unsubscribe();
      cleanupListeners();
    };
  }, []);

  useEffect(() => {
    if (!user || !profile) {
      setSales([]);
      return;
    }

    // By default, load only last 5 days to save quota. 
    // If isSalesLoaded is true, load full window (60 or 30 days).
    const loadingWindowDays = isSalesLoaded ? (isPrivileged ? 60 : 30) : 5;
    const startDateLimit = subDays(new Date(), loadingWindowDays).toISOString().split('T')[0];
    
    const salesFilters = [
      where('agentEmail', '==', user.email || ''),
      where('agent', '==', profile.displayName || '')
    ];

    if (profile.employeeName) {
      salesFilters.push(where('agent', '==', profile.employeeName));
    }
    
    if (profile.bdeName) {
      salesFilters.push(where('bde', '==', profile.bdeName));
      salesFilters.push(where('associateBde', '==', profile.bdeName));
    }

    const salesQuery = isPrivileged
      ? query(collection(db, 'sales'), where('date', '>=', startDateLimit), limit(500))
      : query(collection(db, 'sales'), or(...salesFilters), limit(200));

    const unsubSales = onSnapshot(salesQuery, (snap) => {
      const salesList = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Sales));
      setSales(salesList);
    }, (err) => {
      console.error("Sales listener error:", err);
      if (err.message.includes('Quota exceeded') || err.message.includes('quota limit exceeded')) {
        setIsQuotaExceeded(true);
      }
      if (auth.currentUser && !err.message.includes('Quota exceeded')) {
        setFatalError(err);
        handleFirestoreError(err, OperationType.GET, 'sales');
      }
    });

    return () => unsubSales();
  }, [user, profile, isSalesLoaded, isPrivileged]);

  useEffect(() => {
    let timeoutId: any;

    const fetchAggregates = async () => {
      // If sales are fully loaded in memory, we don't need server-side aggregation
      if (!user || !profile || !['dashboard', 'sales'].includes(activeTab) || isSalesLoaded || isQuotaExceeded) {
        setDashboardAggregates(null);
        return;
      }
      
      setIsLoadingAggregates(true);
      try {
        let level = getPermissionLevel('sales');
        if (level === 'None') {
          setDashboardAggregates(null);
          setIsLoadingAggregates(false);
          return;
        }

        // Determine scope and period to match Cloud Function IDs
        let scope = 'global';
        if (salesAgentFilter) {
          scope = `agent_${salesAgentFilter}`;
        } else if (salesBdeFilter) {
          scope = `bde_${salesBdeFilter}`;
        }

        let type = 'lifetime';
        let periodId = 'total';
        if (salesWeekFilter) {
          type = 'weekly';
          periodId = salesWeekFilter;
        } else if (salesMonthFilter) {
          type = 'monthly';
          periodId = salesMonthFilter;
        }

        const summaryId = `${type}_${scope}_${periodId}`;
        const summaryRef = doc(db, 'salesSummaries', summaryId);
        const summarySnap = await getDoc(summaryRef);

        if (summarySnap.exists()) {
          const data = summarySnap.data();
          // Map summary data to the format expected by the UI
          // Since summaries are currently global/bde/agent totals, 
          // we might need to handle the 'Inhouse/Branch/Franchisee' split.
          // For now, if the summary doesn't have category breakdown, we fallback or use it as 'Inhouse'
          
          const results: any = {
            inhouse: { 
              total: data.totalSalesCount || 0, 
              value: data.totalRevenue || 0,
              advance: data.statusCounts?.['Advance'] || 0,
              advanceValue: data.statusValues?.['Advance'] || 0,
              cn: data.statusCounts?.['Credit Note'] || 0,
              cnValue: data.statusValues?.['Credit Note'] || 0,
              confirmed: (data.statusCounts?.['Confirmed'] || 0) + (data.statusCounts?.['Done'] || 0),
              confirmedValue: (data.statusValues?.['Confirmed'] || 0) + (data.statusValues?.['Done'] || 0),
              cancel: data.statusCounts?.['Cancel'] || 0,
              cancelValue: data.statusValues?.['Cancel'] || 0
            },
            branch: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
            franchisee: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
          };
          
          setDashboardAggregates(results);
          setIsLoadingAggregates(false);
          return;
        }

        // Fallback to getAggregateFromServer if summary doesn't exist
        const baseQuery = collection(db, 'sales');
        const filters: any[] = [];
        
        level = getPermissionLevel('sales');
        if (level === 'Limited') {
          const salesFilters = [
            where('agentEmail', '==', user.email || ''),
            where('agent', '==', profile.displayName || '')
          ];
          if (profile.employeeName) salesFilters.push(where('agent', '==', profile.employeeName));
          if (profile.bdeName) {
            salesFilters.push(where('bde', '==', profile.bdeName));
            salesFilters.push(where('associateBde', '==', profile.bdeName));
          }
          filters.push(or(...salesFilters));
        }

        if (salesAgentFilter) {
          const filterEmp = employees.find(e => e.name === salesAgentFilter);
          if (filterEmp?.email) {
            filters.push(or(where('agent', '==', salesAgentFilter), where('agentEmail', '==', filterEmp.email)));
          } else {
            filters.push(where('agent', '==', salesAgentFilter));
          }
        }
        if (salesBdeFilter) {
          const filterBde = bdes.find(b => b.name === salesBdeFilter);
          if (filterBde?.email) {
            filters.push(or(where('bde', '==', salesBdeFilter), where('bdeEmail', '==', filterBde.email)));
          } else {
            filters.push(where('bde', '==', salesBdeFilter));
          }
        }
        if (salesWeekFilter) {
          filters.push(where('week', '==', salesWeekFilter));
        } else if (salesMonthFilter) {
          const monthWeeks = weeks.filter(w => w.month === salesMonthFilter).map(w => w.weekName);
          if (monthWeeks.length > 0) {
            filters.push(where('week', 'in', monthWeeks));
          } else {
            setDashboardAggregates({
              inhouse: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
              branch: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
              franchisee: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
            });
            setIsLoadingAggregates(false);
            return;
          }
        }
        
        const categories = ['Inhouse', 'Branch', 'Franchisee Sales'];
        const statuses = ['Advance', 'Credit Note', 'Done', 'Confirmed', 'Cancel'];
        
        const results: any = {
          inhouse: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
          branch: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
          franchisee: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
        };

        // We'll do queries for each category and status to get full breakdown
        for (const cat of categories) {
          const catKey = cat === 'Inhouse' ? 'inhouse' : cat === 'Branch' ? 'branch' : 'franchisee';
          
          // If salesByFilter is set and doesn't match this category, skip it
          if (salesByFilter !== 'All' && salesByFilter !== cat) {
            continue;
          }

          const catFilters = [...filters];
          if (cat === 'Inhouse') {
            catFilters.push(or(where('salesBy', '==', 'Inhouse'), where('salesBy', '==', '')));
          } else {
            catFilters.push(where('salesBy', '==', cat));
          }
          
          // Total for category
          const qTotal = query(baseQuery, and(...catFilters));
          const snapTotal = await getAggregateFromServer(qTotal, {
            count: count(),
            value: sum('packageValue')
          });
          results[catKey].total = snapTotal.data().count;
          results[catKey].value = snapTotal.data().value || 0;

          // Status breakdowns
          const statusMap: any = {
            'Advance': 'advance',
            'Credit Note': 'cn',
            'Done': 'confirmed',
            'Confirmed': 'confirmed',
            'Cancel': 'cancel'
          };

          for (const status of statuses) {
            const statusFilters = [...catFilters, where('advanceCN', '==', status)];
            const qStatus = query(baseQuery, and(...statusFilters));
            const snapStatus = await getAggregateFromServer(qStatus, {
              count: count(),
              value: sum('packageValue')
            });
            
            const key = statusMap[status];
            if (key) {
              results[catKey][key] += snapStatus.data().count;
              results[catKey][key + 'Value'] += snapStatus.data().value || 0;
            }
          }
        }

        setDashboardAggregates(results);
      } catch (err: any) {
        console.error("Aggregation error:", err);
        if (err.message?.includes('Quota exceeded') || err.message?.includes('quota limit exceeded')) {
          setIsQuotaExceeded(true);
        } else if (err.message?.includes('permissions')) {
          setError('Dashboard data is being optimized. Please ask an Admin to run "Backfill Sales Summaries" in User Management.');
        } else {
          setError('Aggregation error: ' + (err.message || String(err)));
        }
      } finally {
        setIsLoadingAggregates(false);
      }
    };

    // Debounce the aggregation call to prevent quota exhaustion
    timeoutId = setTimeout(fetchAggregates, 1000);

    return () => clearTimeout(timeoutId);
  }, [user, profile, activeTab, salesAgentFilter, salesBdeFilter, salesWeekFilter, salesMonthFilter, salesByFilter, weeks, isSalesLoaded, isQuotaExceeded]);

  useEffect(() => {
    if (profile && roles.length > 0 && !hasRedirectedRef.current) {
      const userRole = roles.find(r => r.name === profile.role);
      if (userRole && userRole.defaultPage) {
        // Check if user has permission for this page
        if (profile.permissions?.[userRole.defaultPage] && profile.permissions[userRole.defaultPage] !== 'None') {
          setActiveTab(userRole.defaultPage as any);
        }
      } else if (profile.permissions && !profile.permissions['dashboard'] && profile.permissions['employees']) {
        // Fallback for legacy users or roles without defaultPage
        setActiveTab('employees');
      }
      hasRedirectedRef.current = true;
    }
  }, [profile, roles]);

  useEffect(() => {
    if (remarksModalSale?.id) {
      const unsub = onSnapshot(
        query(collection(db, 'sales', remarksModalSale.id, 'remarks'), orderBy('createdAt', 'desc')),
        (snap) => {
          const remarks = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Remark));
          setRemarksLog(remarks);
        },
        (err) => {
          console.error("Remarks listener error:", err);
          handleFirestoreError(err, OperationType.GET, `sales/${remarksModalSale.id}/remarks`);
        }
      );
      return () => unsub();
    } else {
      setRemarksLog([]);
    }
  }, [remarksModalSale]);

  const isRemarkUnread = (sale: Sales) => {
    if (!profile || !sale.remarksCount) return false;
    const lastRead = profile.lastReadRemarks?.[sale.id] || 0;
    return sale.remarksCount > lastRead;
  };

  const handleOpenRemarks = async (sale: Sales) => {
    setRemarksModalSale(sale);
    if (profile && sale.remarksCount && (profile.lastReadRemarks?.[sale.id] || 0) < sale.remarksCount) {
      try {
        const updatedLastRead = {
          ...(profile.lastReadRemarks || {}),
          [sale.id]: sale.remarksCount
        };
        await updateDoc(doc(db, 'users', profile.uid), {
          lastReadRemarks: updatedLastRead,
          updatedAt: serverTimestamp()
        });
        setProfile({ ...profile, lastReadRemarks: updatedLastRead });
      } catch (err) {
        console.error("Error updating last read remarks:", err);
      }
    }
  };

  const handleOpenBravoCalculation = (emp: Employee, numWeeks: number) => {
    const allTargetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks + 1);
    const targetWeeks = allTargetWeeks.slice(0, numWeeks);
    const referenceWeek = allTargetWeeks[numWeeks];

    let totalSales = 0;
    let totalTarget = 0;
    const weeklyBaseTarget = emp.target / 4;
    const agentSales = sales.filter(s => 
      (s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim()) ||
      (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase()) ||
      (s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim()) ||
      (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())
    );

    const calculateWeekData = (wk: any) => {
      const wkSales = agentSales.filter(s => s.week === wk.weekName).reduce((sum, s) => sum + getAgentCreditForSale(s, emp, 'packageValue'), 0);

      const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
      const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
      
      let percentageVal = 0;
      let tenureWeek = 0;
      if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
        tenureWeek = wkIndex - joiningWkIndex + 1;
        if (tenureWeek <= 6) {
          percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
        } else {
          percentageVal = 100;
        }
      }
      const wkTarget = (weeklyBaseTarget * percentageVal) / 100;

      return {
        weekName: wk.weekName,
        sales: wkSales,
        target: wkTarget,
        percentage: percentageVal,
        tenureWeek
      };
    };

    const weeksData: {
      weekName: string;
      sales: number;
      target: number;
      percentage: number;
      tenureWeek: number;
      isReference?: boolean;
    }[] = targetWeeks.map(wk => {
      const data = calculateWeekData(wk);
      totalSales += data.sales;
      totalTarget += data.target;
      return data;
    });

    if (referenceWeek) {
      const refData = calculateWeekData(referenceWeek);
      weeksData.push({ ...refData, isReference: true });
    }

    const achievement = totalTarget > 0 ? Math.round((totalSales / totalTarget) * 100) : 0;

    setBravoCalculationModal({
      isOpen: true,
      title: `${emp.name} - Bravo ${numWeeks} Achievement Calculation`,
      agentName: emp.name,
      numWeeks,
      weeksData,
      totalSales,
      totalTarget,
      achievement
    });
  };

  const handleOpenContributionCalculation = (emp: Employee, numWeeks: number) => {
    const targetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks);
    let totalContribution = 0;
    const weeklyBaseTarget = emp.target / 4;
    const agentSales = sales.filter(s => 
      (s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim()) ||
      (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase()) ||
      (s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim()) ||
      (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())
    );

    const weeksData = targetWeeks.map(wk => {
      const weekSalesRaw = agentSales.filter(s => s.week === wk.weekName);
      const weekSales = weekSalesRaw.filter(s => s.advanceCN !== 'Cancel' && s.advanceCN !== 'Credit Note').reduce((sum, s) => sum + getAgentCreditForSale(s, emp, 'packageValue'), 0);

      const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
      const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
      const isJoined = joiningWkIndex !== -1 && wkIndex >= joiningWkIndex;
      
      const workstationCost = isJoined ? (emp.salary * 3) / 4 : 0;
      const contribution = isJoined ? (weekSales * 0.12) - workstationCost : 0;
      totalContribution += contribution;

      return {
        weekName: wk.weekName,
        sales: weekSales,
        workstationCost,
        contribution
      };
    });

    setContributionCalculationModal({
      isOpen: true,
      title: `${emp.name} - Total Contribution (${numWeeks} Weeks) Calculation`,
      numWeeks,
      totalContribution,
      weeksData
    });
  };

  const handleAddRemark = async () => {
    if (!remarksModalSale?.id || !newRemarkText.trim() || !profile) return;

    setIsAddingRemark(true);
    try {
      const now = new Date();
      const dateStr = now.toLocaleDateString('en-GB'); // DD/MM/YYYY
      const timeStr = now.toLocaleTimeString('en-GB', { hour: '2-digit', minute: '2-digit' });

      const remarkData: Omit<Remark, 'id'> = {
        date: dateStr,
        time: timeStr,
        userName: profile.displayName || 'Unknown',
        userUid: profile.uid,
        text: newRemarkText.trim(),
        createdAt: serverTimestamp()
      };

      await addDoc(collection(db, 'sales', remarksModalSale.id, 'remarks'), remarkData);
      
      // Update remarksCount on the sale document
      const saleRef = doc(db, 'sales', remarksModalSale.id);
      await updateDoc(saleRef, {
        remarksCount: increment(1)
      });
      
      // Update lastReadRemarks for the current user so they don't see their own remark as unread
      if (profile && remarksModalSale.id) {
        const newCount = (remarksModalSale.remarksCount || 0) + 1;
        const updatedLastRead = {
          ...(profile.lastReadRemarks || {}),
          [remarksModalSale.id]: newCount
        };
        await updateDoc(doc(db, 'users', profile.uid), {
          lastReadRemarks: updatedLastRead,
          updatedAt: serverTimestamp()
        });
        setProfile({ ...profile, lastReadRemarks: updatedLastRead });
      }

      setNewRemarkText('');
    } catch (err) {
      console.error("Error adding remark:", err);
      handleFirestoreError(err, OperationType.WRITE, `sales/${remarksModalSale.id}/remarks`);
    } finally {
      setIsAddingRemark(false);
    }
  };

  const handleCreateIncentive = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!incentiveForm.employeeId || !incentiveForm.amount || !incentiveForm.type) return;

    setIsSavingIncentive(true);
    try {
      const employee = employees.find(emp => emp.id === incentiveForm.employeeId);
      const now = new Date();
      const recordedAt = `${now.toLocaleDateString()} ${now.toLocaleTimeString()}`;
      const data: Omit<Incentive, 'id'> = {
        employeeId: incentiveForm.employeeId,
        employeeName: employee?.name || 'Unknown',
        type: incentiveForm.type as any,
        amount: Number(incentiveForm.amount),
        date: incentiveForm.date || new Date().toISOString().split('T')[0],
        remarks: incentiveForm.remarks || '',
        isEligible: incentiveForm.isEligible || false,
        isPaid: incentiveForm.isPaid || false,
        recordedBy: profile?.displayName || 'Unknown',
        recordedAt: recordedAt,
        createdAt: serverTimestamp(),
      };

      if (isEditingIncentive && selectedIncentive?.id) {
        await updateDoc(doc(db, 'incentives', selectedIncentive.id), {
          ...data,
          createdAt: selectedIncentive.createdAt // Keep original createdAt
        });
        setSuccess('Incentive updated successfully');
      } else {
        await addDoc(collection(db, 'incentives'), data);
        setSuccess('Incentive recorded successfully');
      }

      setIsIncentiveModalOpen(false);
      setIsEditingIncentive(false);
      setSelectedIncentive(null);
      setIncentiveForm({
        type: 'Daily',
        date: new Date().toISOString().split('T')[0],
        amount: 0,
        remarks: '',
        isEligible: false,
        isPaid: false
      });
      setTimeout(() => setSuccess(null), 3000);
    } catch (err) {
      console.error("Error saving incentive:", err);
      handleFirestoreError(err, OperationType.WRITE, 'incentives');
    } finally {
      setIsSavingIncentive(false);
    }
  };

  const handleCreatePayment = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!paymentForm.employeeId || !paymentForm.amount) return;

    setIsSavingPayment(true);
    try {
      const employee = employees.find(emp => emp.id === paymentForm.employeeId);
      const now = new Date();
      const recordedAt = `${now.toLocaleDateString()} ${now.toLocaleTimeString()}`;
      const data: Omit<IncentivePayment, 'id'> = {
        employeeId: paymentForm.employeeId,
        employeeName: employee?.name || 'Unknown',
        amount: Number(paymentForm.amount),
        date: paymentForm.date || new Date().toISOString().split('T')[0],
        paymentMethod: paymentForm.paymentMethod || 'Cash',
        remarks: paymentForm.remarks || '',
        recordedBy: profile?.displayName || 'Unknown',
        recordedAt: recordedAt,
        createdAt: serverTimestamp(),
      };

      if (isEditingPayment && selectedPayment?.id) {
        await updateDoc(doc(db, 'incentivePayments', selectedPayment.id), {
          ...data,
          createdAt: selectedPayment.createdAt // Keep original createdAt
        });
        setSuccess('Payment updated successfully');
      } else {
        await addDoc(collection(db, 'incentivePayments'), data);
        setSuccess('Payment recorded successfully');
      }

      setIsPaymentModalOpen(false);
      setIsEditingPayment(false);
      setSelectedPayment(null);
      setPaymentForm({
        date: new Date().toISOString().split('T')[0],
        amount: 0,
        paymentMethod: 'Cash',
        remarks: ''
      });
      setTimeout(() => setSuccess(null), 3000);
    } catch (err) {
      console.error("Error saving payment:", err);
      handleFirestoreError(err, OperationType.WRITE, 'incentivePayments');
    } finally {
      setIsSavingPayment(false);
    }
  };

  const handleEditIncentive = (incentive: Incentive) => {
    setSelectedIncentive(incentive);
    setIncentiveForm({
      employeeId: incentive.employeeId,
      type: incentive.type,
      amount: incentive.amount,
      date: incentive.date,
      remarks: incentive.remarks,
      isEligible: incentive.isEligible || false,
      isPaid: incentive.isPaid || false
    });
    setIsEditingIncentive(true);
    setIsIncentiveModalOpen(true);
  };

  const handleEditPayment = (payment: IncentivePayment) => {
    setSelectedPayment(payment);
    setPaymentForm({
      employeeId: payment.employeeId,
      amount: payment.amount,
      date: payment.date,
      paymentMethod: payment.paymentMethod,
      remarks: payment.remarks
    });
    setIsEditingPayment(true);
    setIsPaymentModalOpen(true);
  };

  const handleLogin = async () => {
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (err) {
      console.error('Login error:', err);
    }
  };

  const handleLogout = async () => {
    hasRedirectedRef.current = false;
    await signOut(auth);
  };

  const handleTripIdSearch = async () => {
    if (!tripIdSearch.trim()) return;
    
    setIsSearchingTripId(true);
    try {
      const q = query(collection(db, 'sales'), where('tripId', '==', tripIdSearch.trim()), limit(1));
      const snap = await getDocs(q);
      
      if (snap.empty) {
        alert("No sale found with this Trip ID.");
      } else {
        const foundSale = { id: snap.docs[0].id, ...snap.docs[0].data() } as Sales;
        // Add to sales if not already there
        setSales(prev => {
          if (prev.some(s => s.id === foundSale.id)) return prev;
          return [foundSale, ...prev];
        });
        // Switch to sales tab if not already there
        setActiveTab('sales');
        // Clear search
        setTripIdSearch('');
      }
    } catch (err) {
      console.error("Trip ID search error:", err);
      handleFirestoreError(err, OperationType.GET, 'sales');
    } finally {
      setIsSearchingTripId(false);
    }
  };

  const handleAwardPoints = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedUser || !profile || pointsToAward <= 0 || !reason.trim()) return;
    if (selectedUser.uid === profile.uid) {
      setError("You cannot award points to yourself!");
      return;
    }

    setIsAwarding(true);
    setError(null);
    setSuccess(null);

    try {
      await runTransaction(db, async (transaction) => {
        const fromRef = doc(db, 'users', profile.uid);
        const toRef = doc(db, 'users', selectedUser.uid);
        const transRef = doc(collection(db, 'transactions'));

        const fromSnap = await transaction.get(fromRef);
        if (!fromSnap.exists()) throw new Error("Sender profile missing");
        
        const fromData = fromSnap.data() as UserProfile;
        if (fromData.pointsBalance < pointsToAward) {
          throw new Error("Insufficient points balance!");
        }

        // Update balances
        transaction.update(fromRef, { pointsBalance: fromData.pointsBalance - pointsToAward });
        
        const toSnap = await transaction.get(toRef);
        const toData = toSnap.data() as UserProfile;
        transaction.update(toRef, { pointsBalance: (toData.pointsBalance || 0) + pointsToAward });

        // Record transaction
        transaction.set(transRef, {
          fromUid: profile.uid,
          fromName: profile.displayName,
          toUid: selectedUser.uid,
          toName: selectedUser.displayName,
          points: pointsToAward,
          reason: reason.trim(),
          createdAt: serverTimestamp()
        });
      });

      setSuccess(`Successfully awarded ${pointsToAward} points to ${selectedUser.displayName}!`);
      setSelectedUser(null);
      setReason('');
      setSearchQuery('');
    } catch (err: any) {
      setError(err.message || "Failed to award points.");
      handleFirestoreError(err, OperationType.WRITE, 'transactions');
    } finally {
      setIsAwarding(false);
    }
  };

  const handleTabChange = (tabId: typeof activeTab) => {
    if (tabId === 'passwordManager') {
      setIsMasterPasswordPromptOpen(true);
      setMasterPasswordInput('');
      setPasswordError(null);
      return;
    }

    const pageConfig = pagePasswords[tabId];
    if (pageConfig && !pageConfig.isOpen && !unlockedPages.has(tabId)) {
      setPendingTab(tabId);
      setIsPasswordPromptOpen(true);
      setPasswordInput('');
      setPasswordError(null);
    } else {
      setActiveTab(tabId);
      setIsMenuOpen(false);
    }
  };

  const verifyPagePassword = () => {
    if (!pendingTab) return;
    const pageConfig = pagePasswords[pendingTab];
    if (pageConfig && passwordInput === pageConfig.password) {
      setUnlockedPages(prev => new Set(prev).add(pendingTab));
      setActiveTab(pendingTab as any);
      setIsPasswordPromptOpen(false);
      setPendingTab(null);
      setPasswordInput('');
      setPasswordError(null);
      setIsMenuOpen(false);
    } else {
      setPasswordError('Invalid password');
    }
  };

  const verifyMasterPassword = () => {
    if (masterPasswordInput === '123456') {
      setActiveTab('passwordManager');
      setIsMasterPasswordPromptOpen(false);
      setMasterPasswordInput('');
      setPasswordError(null);
      setIsMenuOpen(false);
    } else {
      setPasswordError('Invalid master password');
    }
  };

  const handleUpdatePagePassword = async (pageId: string, password: string, isOpen: boolean) => {
    if (!isAdmin) return;
    try {
      await setDoc(doc(db, 'pagePasswords', pageId), {
        pageId,
        password,
        isOpen,
        updatedAt: serverTimestamp()
      });
      setSuccess(`Password for ${pageId} updated successfully`);
    } catch (err: any) {
      handleFirestoreError(err, OperationType.WRITE, `pagePasswords/${pageId}`);
      setError(`Failed to update password for ${pageId}`);
    }
  };

  const handleCreateEmployee = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!hasPermission('employees')) return;
    
    setIsSavingEmployee(true);
    setError(null);
    setSuccess(null);

    try {
      if (isEditingEmployee && employeeForm.id) {
        const docRef = doc(db, 'employees', employeeForm.id);
        await updateDoc(docRef, {
          ...employeeForm,
          updatedAt: serverTimestamp()
        });
        setSuccess('Employee record updated successfully!');
      } else {
        const employeeCode = `EMP-${Math.floor(1000 + Math.random() * 9000)}`;
        await addDoc(collection(db, 'employees'), {
          ...employeeForm,
          employeeCode,
          createdAt: serverTimestamp()
        });
        setSuccess(`Employee ${employeeForm.name} created successfully with code ${employeeCode}`);
      }
      
      setEmployeeForm({
        status: 'Active',
        position: 'Sales Agent',
        joiningWeek: 'Week 1',
        bde: 'None',
        email: '',
        week1Target: 25,
        week2Target: 25,
        week3Target: 50,
        week4Target: 50,
        week5Target: 50,
        week6Target: 50,
        salary: 0,
        completedTrainings: [],
      });
      setIsEditingEmployee(false);
      setIsEmployeeModalOpen(false);
    } catch (err: any) {
      setError(err.message || "Failed to save employee.");
      handleFirestoreError(err, OperationType.WRITE, 'employees');
    } finally {
      setIsSavingEmployee(false);
    }
  };

  const handleEditEmployee = (emp: Employee) => {
    setEmployeeForm(emp);
    setIsEditingEmployee(true);
    setIsEmployeeModalOpen(true);
    setError(null);
    setSuccess(null);
  };

  const handleDeleteEmployee = (id: string, name: string) => {
    if (!isAdmin) return;
    setDeleteConfirmation({ id, type: 'employee', name });
  };

  const confirmDelete = async () => {
    if (!deleteConfirmation) return;
    const { id, type, passwordInput } = deleteConfirmation;

    const deleteConfig = pagePasswords['globalDelete'] || { password: '', isOpen: false };
    if (!deleteConfig.isOpen) {
      if (passwordInput !== deleteConfig.password) {
        setDeleteConfirmation(prev => prev ? { ...prev, error: 'Incorrect password' } : null);
        return;
      }
    }
    
    if (type === 'trainingMaterials' && !hasPermission('training')) return;
    if (type !== 'trainingMaterials' && !isAdmin) return;
    
    try {
      const collectionName = 
        type === 'employee' ? 'employees' : 
        type === 'week' ? 'weeks' : 
        type === 'bde' ? 'bdes' : 
        type === 'trainingMaterials' ? 'trainingMaterials' : 
        type === 'user' ? 'users' :
        type === 'role' ? 'roles' :
        'sales';
        
      await deleteDoc(doc(db, collectionName, id));
      setSuccess(`${type.charAt(0).toUpperCase() + type.slice(1)} deleted successfully`);
      setDeleteConfirmation(null);
    } catch (err: any) {
      const collectionName = 
        type === 'employee' ? 'employees' : 
        type === 'week' ? 'weeks' : 
        type === 'bde' ? 'bdes' : 
        type === 'trainingMaterials' ? 'trainingMaterials' : 
        type === 'user' ? 'users' :
        type === 'role' ? 'roles' :
        'sales';
      handleFirestoreError(err, OperationType.DELETE, `${collectionName}/${id}`);
      setError(`Failed to delete ${type}`);
    }
  };

  const handleDeleteUser = (id: string, name: string) => {
    if (!isAdmin) return;
    setDeleteConfirmation({ id, type: 'user', name });
  };

  const handleDeleteRole = (id: string, name: string) => {
    if (!isAdmin) return;
    setDeleteConfirmation({ id, type: 'role', name });
  };

  const handleCreateWeek = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!hasPermission('weeks')) return;
    
    setIsSavingWeek(true);
    setError(null);
    setSuccess(null);

    try {
      if (isEditingWeek && weekForm.id) {
        const docRef = doc(db, 'weeks', weekForm.id);
        await updateDoc(docRef, {
          ...weekForm,
          updatedAt: serverTimestamp()
        });
        setSuccess('Week record updated successfully!');
      } else {
        await addDoc(collection(db, 'weeks'), {
          ...weekForm,
          createdAt: serverTimestamp()
        });
        setSuccess(`Week ${weekForm.weekName} created successfully!`);
      }
      
      setWeekForm({
        weekName: '',
        month: 'January',
        startDate: '',
        endDate: ''
      });
      setIsEditingWeek(false);
      setIsWeekModalOpen(false);
    } catch (err: any) {
      setError(err.message || "Failed to save week.");
      handleFirestoreError(err, OperationType.WRITE, 'weeks');
    } finally {
      setIsSavingWeek(false);
    }
  };

  const handleEditWeek = (wk: Week) => {
    setWeekForm(wk);
    setIsEditingWeek(true);
    setIsWeekModalOpen(true);
    setError(null);
    setSuccess(null);
  };

  const handleDeleteWeek = (id: string, name: string) => {
    if (!isAdmin) return;
    setDeleteConfirmation({ id, type: 'week', name });
  };

  const handleCreateBDE = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!hasPermission('bdes')) return;
    
    setIsSavingBDE(true);
    setError(null);
    setSuccess(null);

    try {
      if (isEditingBDE && bdeForm.id) {
        const docRef = doc(db, 'bdes', bdeForm.id);
        await updateDoc(docRef, {
          ...bdeForm,
          updatedAt: serverTimestamp()
        });
        setSuccess('BDE record updated successfully!');
      } else {
        await addDoc(collection(db, 'bdes'), {
          ...bdeForm,
          createdAt: serverTimestamp()
        });
        setSuccess(`BDE ${bdeForm.name} created successfully!`);
      }
      
      setBDEForm({ name: '' });
      setIsEditingBDE(false);
      setIsBDEModalOpen(false);
    } catch (err: any) {
      setError(err.message || "Failed to save BDE.");
      handleFirestoreError(err, OperationType.WRITE, 'bdes');
    } finally {
      setIsSavingBDE(false);
    }
  };

  const handleDeleteBDE = (id: string, name: string) => {
    if (!isAdmin) return;
    setDeleteConfirmation({ id, type: 'bde', name });
  };

  const handleSaveMatrixReport = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!hasPermission('matrixMaster')) {
      setError("You do not have permission to submit matrix reports.");
      return;
    }
    
    setIsSavingMatrix(true);
    setError(null);
    setSuccess(null);

    try {
      if (!profile) {
        throw new Error("User profile not loaded. Please refresh.");
      }
      const currentEmp = employees.find(e => e.email?.toLowerCase() === profile?.email?.toLowerCase());
      const employeeId = matrixForm.employeeId || currentEmp?.id || '';
      const employeeName = matrixForm.employeeName || currentEmp?.name || '';
      
      if ((isAdmin || isBDEUser) && !employeeId && !matrixEditId) {
        throw new Error("Please select an agent.");
      }

      if (!employeeId && !isAdmin && !isBDEUser) {
        throw new Error("Only employees or authorized users can submit matrix reports.");
      }

      // Check for duplicate report for the same date and employee
      const existingReport = matrixReports.find(r => 
        r.employeeId === employeeId && 
        r.date === matrixForm.date && 
        r.id !== matrixEditId
      );
      if (existingReport) {
        throw new Error(`A report for ${employeeName} on ${matrixForm.date} already exists.`);
      }

      // Validate all fields are compulsory
      if (!matrixForm.date || !matrixForm.totalTalktime || !matrixForm.loginTime || !matrixForm.logoutTime || !matrixForm.breakTime) {
        throw new Error("All fields are compulsory. Please fill in all details.");
      }

      if (matrixForm.performancePoints === undefined || matrixForm.performancePoints === null || matrixForm.totalCalls === undefined || matrixForm.totalCalls === null) {
        throw new Error("Performance Points and Total Calls are compulsory.");
      }

      // Validate Top 4 Talktimes
      if (!matrixForm.topTalktimes || matrixForm.topTalktimes.length < 4) {
        throw new Error("Please provide all 4 top talktimes.");
      }

      const phoneRegex = /^\d{10}$/;
      matrixForm.topTalktimes.forEach((t, idx) => {
        const phoneNo = t.phoneNo?.trim() || '';
        if (!t.tripId || !phoneNo || !t.talktime) {
          throw new Error(`Top Talktime #${idx + 1} is incomplete. All fields are compulsory.`);
        }
        if (!phoneRegex.test(phoneNo)) {
          throw new Error(`Top Talktime #${idx + 1} phone number must be exactly 10 digits.`);
        }
      });

      // Enforce 3-day limit for submission and editing
      if (!isAdmin) {
        const reportDate = new Date(matrixForm.date || '');
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const diffDays = Math.floor((today.getTime() - reportDate.getTime()) / (1000 * 60 * 60 * 24));
        if (diffDays > 2) {
          throw new Error("You can only submit or edit reports for the last 3 days.");
        }
      }

      const reportData = {
        ...matrixForm,
        employeeId,
        employeeName,
        updatedBy: profile?.email || 'Unknown',
        updatedByName: profile?.displayName || 'Unknown',
        updatedAt: serverTimestamp(),
      };

      if (matrixEditId) {
        await updateDoc(doc(db, 'matrixReports', matrixEditId), reportData);
        setSuccess('Matrix report updated successfully!');
      } else {
        await addDoc(collection(db, 'matrixReports'), {
          ...reportData,
          createdBy: profile?.email || 'Unknown',
          createdByName: profile?.displayName || 'Unknown',
          createdAt: serverTimestamp(),
        });
        setSuccess('Matrix report submitted successfully!');
      }

      setIsMatrixModalOpen(false);
      setMatrixEditId(null);
      setMatrixForm({
        date: format(new Date(), 'yyyy-MM-dd'),
        performancePoints: 0,
        totalCalls: 0,
        totalTalktime: '',
        topTalktimes: [
          { tripId: '', phoneNo: '', talktime: '' },
          { tripId: '', phoneNo: '', talktime: '' },
          { tripId: '', phoneNo: '', talktime: '' },
          { tripId: '', phoneNo: '', talktime: '' },
        ],
        loginTime: '',
        logoutTime: '',
        breakTime: '',
        employeeId: '',
        employeeName: '',
      });
    } catch (err: any) {
      setError(err.message || "Failed to save matrix report.");
      handleFirestoreError(err, OperationType.WRITE, 'matrixReports');
    } finally {
      setIsSavingMatrix(false);
    }
  };

  const handleCreateSales = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!hasPermission('sales')) return;
    
    setIsSavingSales(true);
    setError(null);
    setSuccess(null);

    try {
      // Check for duplicate Trip ID
      if (salesForm.tripId) {
        const duplicate = sales.find(s => 
          s.tripId === salesForm.tripId && 
          (!isEditingSales || s.id !== salesForm.id)
        );
        if (duplicate) {
          throw new Error(`Trip ID "${salesForm.tripId}" already exists. Please use a unique Trip ID.`);
        }
      }

      if (isEditingSales && salesForm.id) {
        const docRef = doc(db, 'sales', salesForm.id);
        await updateDoc(docRef, {
          ...salesForm,
          updatedAt: serverTimestamp()
        });
        setSuccess('Sales record updated successfully!');
      } else {
        await addDoc(collection(db, 'sales'), {
          ...salesForm,
          createdAt: serverTimestamp()
        });
        setSuccess(`Sales record for ${salesForm.guestName} created successfully!`);
      }
      
      setSalesForm({
        week: '',
        date: new Date().toISOString().split('T')[0],
        guestName: '',
        agent: '',
        agentEmail: '',
        agentPercentage: 100,
        bde: '',
        associateBde: '',
        assocBdePercentage: 0,
        destination: '',
        tripId: '',
        tripDate: '',
        packageValue: undefined,
        lessThan10Percent20k: '',
        ppMargin: undefined,
        noOfPax: undefined,
        totalMargin: undefined,
        flight: undefined,
        source: '',
        converted: 'Yes',
        tasksPending: '',
        pHotel: '',
        pFlight: '',
        land: '',
        hflIssue: '',
        workPending: '',
        remarks: '',
      });
      setIsEditingSales(false);
      setIsSalesModalOpen(false);
    } catch (err: any) {
      setError(err.message || "Failed to save sales record.");
      handleFirestoreError(err, OperationType.WRITE, 'sales');
    } finally {
      setIsSavingSales(false);
    }
  };

  const handleEditSales = (sale: Sales) => {
    // Ensure emails are populated if they are missing but the name matches an employee/BDE
    const updatedSale = { ...sale };
    if (!updatedSale.agentEmail && updatedSale.agent) {
      updatedSale.agentEmail = employees.find(e => e.name === updatedSale.agent)?.email || '';
    }
    if (!updatedSale.bdeEmail && updatedSale.bde) {
      updatedSale.bdeEmail = bdes.find(b => b.name === updatedSale.bde)?.email || '';
    }
    if (!updatedSale.associateBdeEmail && updatedSale.associateBde) {
      updatedSale.associateBdeEmail = employees.find(e => e.name === updatedSale.associateBde)?.email || '';
    }
    
    setSalesForm(updatedSale);
    setIsEditingSales(true);
    setIsSalesModalOpen(true);
    setError(null);
    setSuccess(null);
  };

  const handleDeleteSales = (id: string, name: string) => {
    if (!isAdmin) return;
    setDeleteConfirmation({ id, type: 'sales', name });
  };

  const handleEditBDE = (bde: BDE) => {
    setBDEForm(bde);
    setIsEditingBDE(true);
    setIsBDEModalOpen(true);
    setError(null);
    setSuccess(null);
  };

  const filteredUsers = useMemo(() => {
    if (!searchQuery.trim()) return [];
    return allUsers.filter(u => 
      u.uid !== profile?.uid && 
      (u.displayName.toLowerCase().includes(searchQuery.toLowerCase()) || 
       u.email.toLowerCase().includes(searchQuery.toLowerCase()))
    ).slice(0, 5);
  }, [searchQuery, allUsers, profile]);

  const filteredTransactions = useMemo(() => {
    if (isPrivileged) return transactions;
    return transactions.filter(t => 
      !t.fromUid || !t.toUid || // If no owner, visible to all with permission
      t.fromUid === profile?.uid || 
      t.toUid === profile?.uid
    );
  }, [transactions, profile, isPrivileged]);

  const leaderboard = useMemo(() => {
    return [...leaderboardUsers].sort((a, b) => b.pointsBalance - a.pointsBalance).slice(0, 5);
  }, [leaderboardUsers]);

  const allMonths = useMemo(() => {
    const monthMap = new Map<string, string>(); // monthName -> earliestStartDate
    weeks.forEach(w => {
      if (!monthMap.has(w.month) || w.startDate < monthMap.get(w.month)!) {
        monthMap.set(w.month, w.startDate);
      }
    });
    return Array.from(monthMap.entries())
      .sort((a, b) => a[1].localeCompare(b[1]))
      .map(entry => entry[0]);
  }, [weeks]);

  const allSortedWeeks = useMemo(() => {
    return [...weeks].sort((a, b) => a.startDate.localeCompare(b.startDate));
  }, [weeks]);

  const missingWeeksFromSales = useMemo(() => {
    const configuredWeekNames = new Set(weeks.map(w => w.weekName));
    const missing = new Set<string>();
    sales.forEach(s => {
      if (s.week && !configuredWeekNames.has(s.week)) {
        missing.add(s.week);
      }
    });
    return Array.from(missing);
  }, [weeks, sales]);

  const lastTwoWeeksNames = useMemo(() => {
    return allSortedWeeks.slice(-2).map(w => w.weekName);
  }, [allSortedWeeks]);

  const currentWeekName = lastTwoWeeksNames[1];
  const lastWeekName = lastTwoWeeksNames[0];

  const sortedWeeks = useMemo(() => {
    return allSortedWeeks.slice(-24);
  }, [allSortedWeeks]);

  useEffect(() => {
    if (allSortedWeeks.length > 0 && !hasSetDefaultMonth) {
      const latestWeek = allSortedWeeks[allSortedWeeks.length - 1];
      setSalesMonthFilter(latestWeek.month);
      setSalesWeekFilter(''); // Default to all weeks of that month
      setHasSetDefaultMonth(true);
      setHasSetDefaultWeek(true);
    }
  }, [allSortedWeeks, hasSetDefaultMonth]);

  useEffect(() => {
    if (allMonths.length > 0 && !hasSetDefaultStatsMonth) {
      const lastMonthDate = new Date();
      lastMonthDate.setMonth(lastMonthDate.getMonth() - 1);
      const lastMonthName = lastMonthDate.toLocaleString('default', { month: 'long', year: 'numeric' });
      
      if (allMonths.includes(lastMonthName)) {
        setStatsOverviewMonthFilter(lastMonthName);
      } else {
        // Fallback to the latest month in data if last month is not found
        setStatsOverviewMonthFilter(allMonths[allMonths.length - 1]);
      }
      setHasSetDefaultStatsMonth(true);
    }
  }, [allMonths, hasSetDefaultStatsMonth]);

  const getSaleMonth = useCallback((sale: Sales, weeksList: Week[]) => {
    const saleWeekObj = weeksList.find(w => w.weekName === sale.week);
    if (saleWeekObj?.month) return saleWeekObj.month;
    
    if (!sale.week) return null;
    const lowerWeek = sale.week.toLowerCase();
    const months = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
    const foundMonth = months.find(m => lowerWeek.includes(m) || lowerWeek.includes(m.substring(0, 3)));
    if (foundMonth) {
      return foundMonth.charAt(0).toUpperCase() + foundMonth.slice(1);
    }
    return null;
  }, []);

  const filteredSales = useMemo(() => {
    const level = getPermissionLevel('sales');
    if (level === 'None') return [];

    return sales.filter(sale => {
      const filterEmp = employees.find(e => e.name === salesAgentFilter);
      const filterBde = bdes.find(b => b.name === salesBdeFilter);
      
      const matchesAgentFilter = !salesAgentFilter || 
        sale.agent?.toLowerCase() === salesAgentFilter.toLowerCase() ||
        (filterEmp?.email && sale.agentEmail?.toLowerCase() === filterEmp.email.toLowerCase());

      const matchesBde = !salesBdeFilter || 
        sale.bde === salesBdeFilter ||
        (filterBde?.email && sale.bdeEmail?.toLowerCase() === filterBde.email.toLowerCase());
      
      const matchesWeek = !salesWeekFilter || sale.week === salesWeekFilter;
      const matchesSalesBy = salesByFilter === 'All' || 
        (salesByFilter === 'Inhouse' && (!sale.salesBy || sale.salesBy === 'Inhouse')) ||
        sale.salesBy === salesByFilter;
      
      // Find the month for the sale's week
      const saleMonth = getSaleMonth(sale, weeks);
      const matchesMonth = !salesMonthFilter || (saleMonth === salesMonthFilter);
      
      const matchesSearch = !salesSearchQuery || 
        sale.guestName?.toLowerCase().includes(salesSearchQuery.toLowerCase()) ||
        sale.tripId?.toLowerCase().includes(salesSearchQuery.toLowerCase()) ||
        sale.destination?.toLowerCase().includes(salesSearchQuery.toLowerCase()) ||
        sale.agent?.toLowerCase().includes(salesSearchQuery.toLowerCase()) ||
        sale.bde?.toLowerCase().includes(salesSearchQuery.toLowerCase());

      // Privilege restriction: only see own sales if Limited
      const canSeeSale = level === 'Complete' || 
        (level === 'Limited' && (
          (profile?.email && (
            sale.agentEmail?.toLowerCase() === profile.email.toLowerCase() ||
            sale.bdeEmail?.toLowerCase() === profile.email.toLowerCase() ||
            sale.associateBdeEmail?.toLowerCase() === profile.email.toLowerCase()
          )) ||
          sale.agent?.toLowerCase() === profile?.displayName?.toLowerCase() ||
          (currentEmployeeName && sale.agent?.toLowerCase() === currentEmployeeName.toLowerCase()) ||
          (currentBDEName && sale.bde === currentBDEName) ||
          (currentBDEName && sale.associateBde === currentBDEName)
        ));

      return matchesAgentFilter && matchesBde && matchesWeek && matchesMonth && matchesSalesBy && canSeeSale && matchesSearch;
    }).sort((a, b) => {
      const dateA = a.date ? new Date(a.date).getTime() : 0;
      const dateB = b.date ? new Date(b.date).getTime() : 0;
      return dateA - dateB; // Ascending order
    });
  }, [sales, salesAgentFilter, salesBdeFilter, salesWeekFilter, salesMonthFilter, salesByFilter, salesSearchQuery, weeks, profile, getPermissionLevel, currentBDEName, currentEmployeeName]);

  const issueOverviewCounts = useMemo(() => {
    const level = getPermissionLevel('issueOverview');
    if (level === 'None') return { All: 0, Hotel: 0, Flight: 0, Land: 0, HFL: 0, 'Work Pending': 0, 'Credit Note': 0 };

    const baseFiltered = sales.filter(sale => {
      const hasIssue = 
        sale.pHotel === 'Issue' || 
        sale.pFlight === 'Issue' || 
        sale.land === 'Issue' || 
        sale.hflIssue === 'Issue' || 
        sale.workPending === 'Issue' ||
        sale.advanceCN === 'Credit Note';
      
      if (!hasIssue) return false;

      const canSeeSale = level === 'Complete' || 
        (level === 'Limited' && (
          sale.agentEmail?.toLowerCase() === profile?.email?.toLowerCase() || 
          sale.agent?.toLowerCase() === profile?.displayName?.toLowerCase() ||
          (currentEmployeeName && sale.agent?.toLowerCase() === currentEmployeeName.toLowerCase()) ||
          (currentBDEName && sale.bde === currentBDEName) ||
          (currentBDEName && sale.associateBde === currentBDEName)
        ));

      return canSeeSale;
    });

    return {
      All: baseFiltered.filter(s => s.advanceCN !== 'Credit Note').length,
      Hotel: baseFiltered.filter(s => s.pHotel === 'Issue').length,
      Flight: baseFiltered.filter(s => s.pFlight === 'Issue').length,
      Land: baseFiltered.filter(s => s.land === 'Issue').length,
      HFL: baseFiltered.filter(s => s.hflIssue === 'Issue').length,
      'Work Pending': baseFiltered.filter(s => s.workPending === 'Issue').length,
      'Credit Note': baseFiltered.filter(s => s.advanceCN === 'Credit Note').length
    };
  }, [sales, profile, getPermissionLevel, currentBDEName, currentEmployeeName]);

  const globalIssueCount = useMemo(() => {
    const level = getPermissionLevel('issueOverview');
    if (level === 'None') return 0;

    return sales.filter(sale => {
      const hasIssue = 
        sale.pHotel === 'Issue' || 
        sale.pFlight === 'Issue' || 
        sale.land === 'Issue' || 
        sale.hflIssue === 'Issue' || 
        sale.workPending === 'Issue' ||
        sale.advanceCN === 'Credit Note';
      
      if (!hasIssue) return false;

      const canSeeSale = level === 'Complete' || 
        (level === 'Limited' && (
          sale.agentEmail?.toLowerCase() === profile?.email?.toLowerCase() || 
          sale.agent?.toLowerCase() === profile?.displayName?.toLowerCase() ||
          (currentEmployeeName && sale.agent?.toLowerCase() === currentEmployeeName.toLowerCase()) ||
          (currentBDEName && sale.bde === currentBDEName) ||
          (currentBDEName && sale.associateBde === currentBDEName)
        ));

      return canSeeSale;
    }).length;
  }, [sales, profile, getPermissionLevel, currentBDEName, currentEmployeeName]);

  const filteredIssueSales = useMemo(() => {
    const level = getPermissionLevel('issueOverview');
    if (level === 'None') return [];

    return sales.filter(sale => {
      const hasIssue = 
        sale.pHotel === 'Issue' || 
        sale.pFlight === 'Issue' || 
        sale.land === 'Issue' || 
        sale.hflIssue === 'Issue' || 
        sale.workPending === 'Issue' ||
        sale.advanceCN === 'Credit Note';
      
      if (!hasIssue) return false;

      const matchesTab = 
        (issueOverviewActiveTab === 'All' && sale.advanceCN !== 'Credit Note') ||
        (issueOverviewActiveTab === 'Hotel' && sale.pHotel === 'Issue') ||
        (issueOverviewActiveTab === 'Flight' && sale.pFlight === 'Issue') ||
        (issueOverviewActiveTab === 'Land' && sale.land === 'Issue') ||
        (issueOverviewActiveTab === 'HFL' && sale.hflIssue === 'Issue') ||
        (issueOverviewActiveTab === 'Work Pending' && sale.workPending === 'Issue') ||
        (issueOverviewActiveTab === 'Credit Note' && sale.advanceCN === 'Credit Note');

      const canSeeSale = level === 'Complete' || 
        (level === 'Limited' && (
          sale.agentEmail?.toLowerCase() === profile?.email?.toLowerCase() || 
          sale.agent?.toLowerCase() === profile?.displayName?.toLowerCase() ||
          (currentEmployeeName && sale.agent?.toLowerCase() === currentEmployeeName.toLowerCase()) ||
          (currentBDEName && sale.bde === currentBDEName) ||
          (currentBDEName && sale.associateBde === currentBDEName)
        ));

      return matchesTab && canSeeSale;
    }).sort((a, b) => {
      const dateA = a.date ? new Date(a.date).getTime() : 0;
      const dateB = b.date ? new Date(b.date).getTime() : 0;
      return dateA - dateB; // Chronological order
    });
  }, [sales, issueOverviewActiveTab, profile, getPermissionLevel, currentBDEName, currentEmployeeName]);

  const filteredIssueOverviewAllSales = useMemo(() => {
    const level = getPermissionLevel('issueOverview');
    if (level === 'None') return [];

    return sales.filter(sale => {
      let matchesWeek = false;
      if (issueOverviewWeekFilter === 'All') {
        matchesWeek = lastTwoWeeksNames.includes(sale.week);
      } else if (issueOverviewWeekFilter === 'Current') {
        matchesWeek = sale.week === currentWeekName;
      } else if (issueOverviewWeekFilter === 'Last') {
        matchesWeek = sale.week === lastWeekName;
      }

      const canSeeSale = level === 'Complete' || 
        (level === 'Limited' && (
          sale.agentEmail?.toLowerCase() === profile?.email?.toLowerCase() || 
          sale.agent?.toLowerCase() === profile?.displayName?.toLowerCase() ||
          (currentEmployeeName && sale.agent?.toLowerCase() === currentEmployeeName.toLowerCase()) ||
          (currentBDEName && sale.bde === currentBDEName) ||
          (currentBDEName && sale.associateBde === currentBDEName)
        ));

      return matchesWeek && canSeeSale;
    }).sort((a, b) => {
      const dateA = a.date ? new Date(a.date).getTime() : 0;
      const dateB = b.date ? new Date(b.date).getTime() : 0;
      return dateA - dateB; // Newest first (Ascending as requested)
    });
  }, [sales, issueOverviewWeekFilter, lastTwoWeeksNames, currentWeekName, lastWeekName, profile, getPermissionLevel, currentBDEName, currentEmployeeName]);

  const totalPkgValue = useMemo(() => {
    return filteredSales.reduce((sum, sale) => sum + (sale.packageValue || 0), 0);
  }, [filteredSales]);

  const salesStats = useMemo(() => {
    const stats = {
      inhouse: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
      branch: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
      franchisee: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
    };

    filteredSales.forEach(s => {
      let category: 'inhouse' | 'branch' | 'franchisee' = 'inhouse';
      if (s.salesBy === 'Branch') category = 'branch';
      else if (s.salesBy === 'Franchisee Sales') category = 'franchisee';

      const val = s.packageValue || 0;
      stats[category].total += 1;
      stats[category].value += val;

      if (s.advanceCN === 'Advance') {
        stats[category].advance += 1;
        stats[category].advanceValue += val;
      } else if (s.advanceCN === 'Credit Note') {
        stats[category].cn += 1;
        stats[category].cnValue += val;
      } else if (s.advanceCN === 'Done' || s.advanceCN === 'Confirmed') {
        stats[category].confirmed += 1;
        stats[category].confirmedValue += val;
      } else if (s.advanceCN === 'Cancel') {
        stats[category].cancel += 1;
        stats[category].cancelValue += val;
      }
    });

    return stats;
  }, [filteredSales]);

  if (fatalError) {
    let message = fatalError.message || "An unexpected error occurred. Please try refreshing the page.";
    let isQuotaError = false;
    
    try {
      const errorString = typeof fatalError === 'string' ? fatalError : (fatalError.message || String(fatalError));
      if (errorString.includes('Quota exceeded') || errorString.includes('quota limit exceeded')) {
        isQuotaError = true;
        message = "Firestore daily quota exceeded. The free tier limit has been reached for today. The quota will reset tomorrow. You can check detailed quota info at https://firebase.google.com/pricing#cloud-firestore";
      } else if (message.startsWith('{')) {
        const errObj = JSON.parse(message);
        if (errObj.error) {
          if (errObj.error.includes('Quota exceeded')) {
            isQuotaError = true;
            message = "Firestore daily quota exceeded. The free tier limit has been reached for today. The quota will reset tomorrow.";
          } else {
            message = `Firestore Error: ${errObj.error}`;
          }
        }
      }
    } catch (e) {
      // Not a JSON string, keep original message
    }

    return (
      <div className="min-h-screen bg-zinc-50 flex flex-col items-center justify-center p-4 text-center">
        <div className={cn("p-3 rounded-full mb-4", isQuotaError ? "bg-orange-100" : "bg-red-100")}>
          <AlertCircle className={cn("w-12 h-12", isQuotaError ? "text-orange-600" : "text-red-600")} />
        </div>
        <h1 className="text-2xl font-bold text-zinc-900 mb-2">
          {isQuotaError ? "Quota Limit Reached" : "Application Error"}
        </h1>
        <p className="text-zinc-600 mb-6 max-w-md whitespace-pre-wrap">{message}</p>
        <Button onClick={() => window.location.reload()} className={isQuotaError ? "bg-orange-600 hover:bg-orange-700" : ""}>
          Refresh Page
        </Button>
      </div>
    );
  }

  if (isLoading) {
    return (
      <div className="min-h-screen bg-zinc-50 flex items-center justify-center">
        <Loader2 className="w-8 h-8 text-orange-600 animate-spin" />
      </div>
    );
  }

  const handleUpdateUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    
    setIsSavingUser(true);
    setError(null);
    setSuccess(null);

    try {
      if (isEditingUser && userForm.uid) {
        const userRef = doc(db, 'users', userForm.uid);
        await updateDoc(userRef, {
          role: userForm.role,
          permissions: userForm.permissions,
          updatedAt: serverTimestamp()
        });
        setSuccess('User permissions updated successfully!');
      } else {
        // Pre-provisioning a user profile
        if (!userForm.email) throw new Error("Email is required");
        
        // Check if user already exists
        const q = query(collection(db, 'users'), where('email', '==', userForm.email), limit(1));
        const querySnap = await getDocs(q);
        if (!querySnap.empty) {
          throw new Error("A user with this email already exists.");
        }

        const newUserRef = doc(collection(db, 'users'));
        await setDoc(newUserRef, {
          ...userForm,
          role: userForm.role || 'Agent',
          permissions: userForm.permissions || ['dashboard'],
          uid: newUserRef.id, // Temporary ID, will be migrated on login
          displayName: userForm.displayName || userForm.email.split('@')[0],
          photoURL: userForm.photoURL || `https://ui-avatars.com/api/?name=${userForm.email}&background=random`,
          pointsBalance: 100,
          createdAt: serverTimestamp()
        });
        setSuccess('User pre-provisioned successfully!');
      }
      setIsUserModalOpen(false);
    } catch (err: any) {
      setError(err.message || "Failed to update user.");
      handleFirestoreError(err, OperationType.WRITE, `users`);
    } finally {
      setIsSavingUser(false);
    }
  };

  const handleEditUser = (user: UserProfile) => {
    setUserForm(user);
    setIsEditingUser(true);
    setIsUserModalOpen(true);
    setError(null);
    setSuccess(null);
  };

  const handleCreateUser = () => {
    setUserForm({
      role: 'Agent',
      permissions: ['employees'],
      email: '',
      displayName: '',
    });
    setIsEditingUser(false);
    setIsUserModalOpen(true);
    setError(null);
    setSuccess(null);
  };

  const downloadEmployeeMaster = () => {
    const data = employees.map(({ id, ...rest }) => {
      const formatted: any = {};
      Object.entries(rest).forEach(([key, value]) => {
        // Convert camelCase to Title Case for headers (e.g. employeeCode -> Employee Code)
        const titleKey = key.replace(/([A-Z])/g, ' $1').replace(/^./, str => str.toUpperCase());
        if (Array.isArray(value)) {
          formatted[titleKey] = value.join(', ');
        } else if (typeof value === 'boolean') {
          formatted[titleKey] = value ? 'Yes' : 'No';
        } else {
          formatted[titleKey] = value;
        }
      });
      return formatted;
    });
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Employee Master");
    XLSX.writeFile(workbook, "Employee_Master.xlsx");
  };

  const downloadSalesMaster = () => {
    const sortedSales = [...sales].sort((a, b) => {
      const dateA = a.date ? new Date(a.date).getTime() : 0;
      const dateB = b.date ? new Date(b.date).getTime() : 0;
      return dateB - dateA;
    });
    const data = sortedSales.map(s => ({
      'Guest Name': s.guestName,
      'Week': s.week,
      'Date': s.date,
      'Agent': s.agent,
      '% Agent': s.agentPercentage,
      'BDE': s.bde,
      'Assoc BDE': s.associateBde,
      '% Assoc BDE': s.assocBdePercentage,
      'Destination': s.destination,
      'Trip ID': s.tripId,
      'Trip Date': s.tripDate,
      'Pkg Value': s.packageValue,
      'Less than 10% / 20k': s.lessThan10Percent20k,
      'Advance/CN': s.advanceCN,
      'Sales by': s.salesBy || 'Inhouse',
      'PP Margin': s.ppMargin,
      'Pax': s.noOfPax,
      'Total Margin': s.totalMargin,
      'Flight': s.flight,
      'Source': s.source,
      'Converted': s.converted,
      'Tasks Pending': s.tasksPending,
      'P - Hotel': s.pHotel,
      'P Flight': s.pFlight,
      'Land': s.land,
      'HFL Issue': s.hflIssue,
      'Work Pending': s.workPending,
      'Remarks': s.remarks
    }));
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sales Master");
    XLSX.writeFile(workbook, "Sales_Master.xlsx");
  };

  const downloadBDEComparison = (data: any[], timeframe: string) => {
    const worksheet = XLSX.utils.json_to_sheet(data.map(d => ({
      'BDE Name': d.name,
      'Total Sales': d.totalSales,
      'Approx Margin (12%)': d.approxMargin,
      'Expenses': d.expenses,
      'Net Profit': d.netProfit,
      'Sales/Salary Ratio': d.salesSalaryRatio.toFixed(2)
    })));
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "BDE Comparison");
    XLSX.writeFile(workbook, `BDE_Comparison_${timeframe.replace(/\s+/g, '_')}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const downloadWeeklyOverview = () => {
    const activeEmployees = employees.filter(e => e.status === 'Active');
    
    const getOverallBravoAchievement = (numWeeks: number) => {
      const targetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks);
      let totalSales = 0;
      let totalTarget = 0;

      activeEmployees.forEach(emp => {
        const agentSales = sales.filter(s => 
          s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
          (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase()) ||
          s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
          (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())
        );
        const weeklyBaseTarget = emp.target / 4;

        targetWeeks.forEach(wk => {
          const wkSales = agentSales.filter(s => s.week === wk.weekName).reduce((sum, s) => {
            let credit = 0;
            if ((s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim()) || (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase())) {
              credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
            }
            if ((s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim()) || (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())) {
              credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
            }
            return sum + credit;
          }, 0);
          totalSales += wkSales;

          const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
          const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
          
          let percentageVal = 0;
          if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
            const tenureWeek = wkIndex - joiningWkIndex + 1;
            if (tenureWeek <= 6) {
              percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
            } else {
              percentageVal = 100;
            }
          }
          totalTarget += (weeklyBaseTarget * percentageVal) / 100;
        });
      });

      return totalTarget > 0 ? Math.round((totalSales / totalTarget) * 100) : 0;
    };

    const getTotalLastWeekTarget = () => {
      const lastWeek = allSortedWeeks[allSortedWeeks.length - 1];
      if (!lastWeek) return 0;
      
      let totalTarget = 0;
      activeEmployees.forEach(emp => {
        const weeklyBaseTarget = emp.target / 4;
        const wkIndex = allSortedWeeks.findIndex(w => w.weekName === lastWeek.weekName);
        const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
        
        let percentageVal = 0;
        if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
          const tenureWeek = wkIndex - joiningWkIndex + 1;
          if (tenureWeek <= 6) {
            percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
          } else {
            percentageVal = 100;
          }
        }
        totalTarget += (weeklyBaseTarget * percentageVal) / 100;
      });
      return totalTarget;
    };

    const totalRow: any = {
      'Agent': 'TOTAL',
      'BDE': '-',
      'Weekly Target': getTotalLastWeekTarget(),
      'Joining Date': '-',
      'Bravo 4 %': `${getOverallBravoAchievement(4)}%`,
      'Bravo 8 %': `${getOverallBravoAchievement(8)}%`
    };

    [...sortedWeeks].reverse().forEach((w) => {
      const totalWeekSales = sales.filter(s => s.week === w.weekName).reduce((sum, s) => {
        let credit = 0;
        if (activeEmployees.some(e => (e.name.toLowerCase().trim() === s.agent?.toLowerCase().trim()) || (s.agentEmail && e.email?.toLowerCase() === s.agentEmail.toLowerCase()))) {
          credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
        }
        if (s.associateBde && activeEmployees.some(e => (e.name.toLowerCase().trim() === s.associateBde?.toLowerCase().trim()) || (s.associateBdeEmail && e.email?.toLowerCase() === s.associateBdeEmail.toLowerCase()))) {
          credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
        }
        return sum + credit;
      }, 0);
      totalRow[`${w.weekName} %`] = '-';
      totalRow[`${w.weekName} Sales`] = totalWeekSales;
    });

    const agentData = activeEmployees.map(emp => {
      const agentSales = sales.filter(s => 
        s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
        (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase()) ||
        s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
        (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())
      );
      
      const getBravoAchievement = (numWeeks: number) => {
        const targetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks);
        let totalSales = 0;
        let totalTarget = 0;
        const weeklyBaseTarget = emp.target / 4;

        targetWeeks.forEach(wk => {
          const wkSales = agentSales.filter(s => s.week === wk.weekName).reduce((sum, s) => sum + getAgentCreditForSale(s, emp, 'packageValue'), 0);
          totalSales += wkSales;

          const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
          const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
          
          let percentageVal = 0;
          if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
            const tenureWeek = wkIndex - joiningWkIndex + 1;
            if (tenureWeek <= 6) {
              percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
            } else {
              percentageVal = 100;
            }
          }
          totalTarget += (weeklyBaseTarget * percentageVal) / 100;
        });

        return totalTarget > 0 ? Math.round((totalSales / totalTarget) * 100) : 0;
      };

      const getLastWeekTarget = () => {
        const lastWeek = allSortedWeeks[allSortedWeeks.length - 1];
        if (!lastWeek) return 0;
        
        const weeklyBaseTarget = emp.target / 4;
        const wkIndex = allSortedWeeks.findIndex(w => w.weekName === lastWeek.weekName);
        const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
        
        let percentageVal = 0;
        if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
          const tenureWeek = wkIndex - joiningWkIndex + 1;
          if (tenureWeek <= 6) {
            percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
          } else {
            percentageVal = 100;
          }
        }
        return (weeklyBaseTarget * percentageVal) / 100;
      };

      const row: any = {
        'Agent': emp.name,
        'BDE': emp.bde,
        'Weekly Target': getLastWeekTarget(),
        'Joining Date': emp.joiningDate,
        'Bravo 4 %': `${getBravoAchievement(4)}%`,
        'Bravo 8 %': `${getBravoAchievement(8)}%`
      };

      [...sortedWeeks].reverse().forEach((w) => {
        const weekSales = agentSales.filter(s => s.week === w.weekName).reduce((sum, s) => sum + getAgentCreditForSale(s, emp, 'packageValue'), 0);
        const wkIndex = allSortedWeeks.findIndex(sw => sw.weekName === w.weekName);
        const joiningWeekIndex = allSortedWeeks.findIndex(sw => sw.weekName === emp.joiningWeek);
        
        let percentage = '-';
        if (joiningWeekIndex !== -1 && wkIndex >= joiningWeekIndex) {
          const tenureWeek = wkIndex - joiningWeekIndex + 1;
          if (tenureWeek <= 6) {
            percentage = `${(emp as any)[`week${tenureWeek}Target`]}%`;
          } else {
            percentage = '100%';
          }
        }
        
        row[`${w.weekName} %`] = percentage;
        row[`${w.weekName} Sales`] = weekSales;
      });

      return row;
    });

    const data = [totalRow, ...agentData];
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Weekly Overview");
    XLSX.writeFile(workbook, "Weekly_Overview.xlsx");
  };

  if (!user) {
    return (
      <div className="min-h-screen bg-zinc-50 flex flex-col items-center justify-center p-4">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="max-w-md w-full text-center space-y-8"
        >
          <div className="flex justify-center">
            <div className="w-20 h-20 bg-orange-100 rounded-3xl flex items-center justify-center shadow-inner">
              <Award className="w-10 h-10 text-orange-600" />
            </div>
          </div>
          <div className="space-y-2">
            <h1 className="text-4xl font-black tracking-tight text-zinc-900">Nitsa Bravo</h1>
            <p className="text-zinc-500 font-medium">The peer recognition platform for exceptional contributions.</p>
          </div>
          <Card className="p-8 space-y-6">
            <p className="text-sm text-zinc-600">Sign in with your Google account to start awarding and receiving Bravo points.</p>
            <Button onClick={handleLogin} className="w-full py-4 text-lg" variant="primary">
              <img src="https://www.google.com/favicon.ico" className="w-5 h-5 mr-2" alt="Google" />
              Continue with Google
            </Button>
          </Card>
          <p className="text-xs text-zinc-400">By continuing, you agree to our terms of service.</p>
        </motion.div>
      </div>
    );
  }

  if (!profile) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-zinc-50">
        <Loader2 className="w-10 h-10 text-orange-600 animate-spin" />
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-zinc-50 text-zinc-900 font-sans selection:bg-orange-100">
        {/* --- Header --- */}
      <header className="sticky top-0 z-50 bg-white/80 backdrop-blur-md border-bottom border-zinc-200">
        <div className="max-w-7xl mx-auto px-4 h-16 flex items-center justify-between">
          <div 
            className="flex items-center gap-2 cursor-pointer hover:opacity-80 transition-opacity"
            onClick={() => handleTabChange('dashboard')}
          >
            <div className="w-8 h-8 bg-orange-600 rounded-lg flex items-center justify-center">
              <Award className="w-5 h-5 text-white" />
            </div>
            <span className="font-bold text-lg tracking-tight">Nitsa Bravo</span>
          </div>

          <div className="flex items-center gap-2 sm:gap-4">
            <div className="hidden sm:flex items-center gap-2 bg-zinc-100 px-3 py-1.5 rounded-full">
              <Star className="w-4 h-4 text-orange-500 fill-orange-500" />
              <span className="text-sm font-bold">{profile?.pointsBalance || 0} Points</span>
            </div>
            
            <div className="flex items-center gap-2">
              <img 
                src={profile?.photoURL} 
                className="w-8 h-8 rounded-full border border-zinc-200" 
                alt={profile?.displayName} 
                referrerPolicy="no-referrer"
              />
              
              {profile && (
                <div className="relative" ref={menuRef}>
                  <button 
                    className="p-2 bg-zinc-100 rounded-lg hover:bg-zinc-200 transition-colors" 
                    onClick={() => setIsMenuOpen(!isMenuOpen)}
                  >
                    <Menu className="w-6 h-6 text-zinc-700" />
                  </button>
                  {isMenuOpen && (
                    <div className="absolute right-0 top-full mt-2 w-[calc(100vw-2rem)] sm:w-max max-w-[280px] sm:max-w-none bg-white border border-zinc-200 rounded-xl shadow-xl p-4 sm:p-6 flex flex-col sm:flex-row gap-6 sm:gap-8 z-50 overflow-y-auto max-h-[80vh]">
                      <div className="flex flex-col gap-2">
                        <h4 className="font-bold text-zinc-900 border-b pb-2 mb-2">General</h4>
                        <button onClick={() => { handleTabChange('dashboard'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Dashboard</button>
                        <button onClick={handleLogout} className="text-sm text-red-600 hover:text-red-700 text-left flex items-center gap-2 mt-2">
                          <LogOut className="w-4 h-4" />
                          Logout
                        </button>
                      </div>
                      <div className="flex flex-col gap-2">
                        <h4 className="font-bold text-zinc-900 border-b pb-2 mb-2">Masters</h4>
                        {hasPermission('employees') && (
                          <button onClick={() => { handleTabChange('employees'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Employee master</button>
                        )}
                        {hasPermission('weeks') && (
                          <button onClick={() => { handleTabChange('weeks'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Weeks master</button>
                        )}
                        {hasPermission('bdes') && (
                          <button onClick={() => { handleTabChange('bdes'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">BDE master</button>
                        )}
                        {hasPermission('sales') && (
                          <button onClick={() => { handleTabChange('sales'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Sales Master</button>
                        )}
                        {hasPermission('training') && (
                          <button onClick={() => { handleTabChange('training'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Training Master</button>
                        )}
                        {hasPermission('incentiveMaster') && (
                          <button onClick={() => { handleTabChange('incentiveMaster'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Incentive Master</button>
                        )}
                        {hasPermission('matrixMaster') && (
                          <button onClick={() => { handleTabChange('matrixMaster'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Matrix Master</button>
                        )}
                      </div>
                      <div className="flex flex-col gap-2">
                        <h4 className="font-bold text-zinc-900 border-b pb-2 mb-2">Review</h4>
                        {hasPermission('weeklyOverview') && (
                          <button onClick={() => { handleTabChange('weeklyOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Weekly overview</button>
                        )}
                        {hasPermission('agentOverview') && (
                          <button onClick={() => { handleTabChange('agentOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left flex items-center justify-between group">
                            <span>Agent Overview</span>
                            {globalIssueCount > 0 && (
                              <span className="bg-red-100 text-red-600 text-[10px] font-bold px-1.5 py-0.5 rounded-full group-hover:bg-red-200 transition-colors">
                                {globalIssueCount}
                              </span>
                            )}
                          </button>
                        )}
                        {hasPermission('trainingOverview') && (
                          <button onClick={() => { handleTabChange('trainingOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Training Overview</button>
                        )}
                        {hasPermission('bdeOverview') && (
                          <button onClick={() => { handleTabChange('bdeOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left flex items-center justify-between group">
                            <span>BDE Overview</span>
                            {globalIssueCount > 0 && (
                              <span className="bg-red-100 text-red-600 text-[10px] font-bold px-1.5 py-0.5 rounded-full group-hover:bg-red-200 transition-colors">
                                {globalIssueCount}
                              </span>
                            )}
                          </button>
                        )}
                        {hasPermission('statsOverview') && (
                          <button onClick={() => { handleTabChange('statsOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Stats Overview</button>
                        )}
                        {hasPermission('issueOverview') && (
                          <button onClick={() => { handleTabChange('issueOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left flex items-center justify-between group">
                            <span>Issue Overview</span>
                            {globalIssueCount > 0 && (
                              <span className="bg-red-600 text-white text-[10px] font-bold px-1.5 py-0.5 rounded-full shadow-sm">
                                {globalIssueCount}
                              </span>
                            )}
                          </button>
                        )}
                      </div>
                      <div className="flex flex-col gap-2">
                        <h4 className="font-bold text-zinc-900 border-b pb-2 mb-2">Management</h4>
                        {hasPermission('userManagement') && (
                          <button onClick={() => { handleTabChange('userManagement'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">User management</button>
                        )}
                        {hasPermission('passwordManager') && (
                          <button onClick={() => { handleTabChange('passwordManager'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Password Manager</button>
                        )}
                      </div>
                    </div>
                  )}
                </div>
              )}
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
        {isQuotaExceeded && (
          <div className="mb-6 p-4 bg-orange-50 border border-orange-200 rounded-xl flex items-center gap-3 text-orange-800 shadow-sm animate-in fade-in slide-in-from-top-2">
            <div className="p-2 bg-orange-100 rounded-lg">
              <Database className="w-5 h-5 text-orange-600" />
            </div>
            <div className="flex-1">
              <p className="text-sm font-bold">Viewing Cached Data</p>
              <p className="text-xs opacity-90">The daily Firestore quota has been reached. You are viewing the last-seen data from your device's cache. New updates will appear once the quota resets tomorrow.</p>
            </div>
          </div>
        )}
        <Modal
          isOpen={summaryDetailModal.isOpen}
          onClose={() => setSummaryDetailModal({ ...summaryDetailModal, isOpen: false })}
          title={summaryDetailModal.title}
        >
          <div className="max-h-[70vh] overflow-auto">
            <table className="w-full text-left border-collapse">
              <thead className="sticky top-0 bg-white shadow-sm z-10">
                <tr>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b text-center">Log</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Date</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Guest</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Trip ID</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Dest.</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Agent</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Assoc BDE</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b text-center">% Agent</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b text-center">% Assoc BDE</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Value</th>
                  <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border-b">Status</th>
                </tr>
              </thead>
              <tbody>
                {[...summaryDetailModal.sales]
                  .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
                  .map((sale, idx) => (
                  <tr key={sale.id || idx} className="hover:bg-zinc-50 transition-colors border-b last:border-0">
                    <td className="px-4 py-3 border-b text-center">
                      <button
                        onClick={() => handleOpenRemarks(sale)}
                        className={cn(
                          "inline-flex items-center gap-1 px-2 py-1 rounded-lg border transition-all",
                          isRemarkUnread(sale) 
                            ? "bg-orange-600 border-orange-600 text-white shadow-md shadow-orange-600/20" 
                            : "border-zinc-200 hover:bg-white hover:border-orange-300 hover:text-orange-600 text-zinc-400"
                        )}
                      >
                        <MessageSquare className={cn(
                          "w-3 h-3",
                          isRemarkUnread(sale) ? "text-white" : "text-zinc-400"
                        )} />
                        <span className={cn(
                          "text-[10px] font-black",
                          isRemarkUnread(sale) ? "text-white" : "text-zinc-900"
                        )}>{sale.remarksCount || 0}</span>
                      </button>
                    </td>
                    <td className="px-4 py-3 text-sm text-zinc-600">{sale.date}</td>
                    <td className="px-4 py-3 text-sm font-medium text-zinc-900">{sale.guestName}</td>
                    <td className="px-4 py-3 text-sm text-zinc-600">{sale.tripId}</td>
                    <td className="px-4 py-3 text-sm text-zinc-600">{sale.destination}</td>
                    <td className="px-4 py-3 text-sm text-zinc-600">{sale.agent}</td>
                    <td className="px-4 py-3 text-sm text-zinc-600">{sale.associateBde}</td>
                    <td className="px-4 py-3 text-sm text-center font-medium">{sale.agentPercentage}%</td>
                    <td className="px-4 py-3 text-sm text-center font-medium">{sale.assocBdePercentage}%</td>
                    <td className="px-4 py-3 text-sm font-bold text-zinc-900">₹{(sale.packageValue || 0).toLocaleString()}</td>
                    <td className="px-4 py-3 text-sm">
                      <span className={cn(
                        "px-2 py-0.5 rounded-full text-[10px] font-bold uppercase",
                        sale.advanceCN === 'Done' || sale.advanceCN === 'Confirmed' ? "bg-emerald-100 text-emerald-700" :
                        sale.advanceCN === 'Cancel' ? "bg-red-100 text-red-700" :
                        "bg-orange-100 text-orange-700"
                      )}>
                        {sale.advanceCN}
                      </span>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </Modal>

        <Modal
          isOpen={bravoCalculationModal.isOpen}
          onClose={() => setBravoCalculationModal({ ...bravoCalculationModal, isOpen: false })}
          title={bravoCalculationModal.title}
        >
          <div className="max-h-[70vh] overflow-auto p-4 space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
              <Card className="p-4 bg-blue-50 border-blue-100">
                <p className="text-[10px] font-bold text-blue-600 uppercase tracking-wider mb-1">Total Sales</p>
                <p className="text-xl font-bold text-blue-900">₹{bravoCalculationModal.totalSales.toLocaleString()}</p>
              </Card>
              <Card className="p-4 bg-purple-50 border-purple-100">
                <p className="text-[10px] font-bold text-purple-600 uppercase tracking-wider mb-1">Total Target</p>
                <p className="text-xl font-bold text-purple-900">₹{bravoCalculationModal.totalTarget.toLocaleString()}</p>
              </Card>
              <Card className="p-4 bg-orange-50 border-orange-100">
                <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider mb-1">Achievement</p>
                <p className="text-xl font-bold text-orange-900">{bravoCalculationModal.achievement}%</p>
              </Card>
            </div>

            <div className="space-y-3">
              <h4 className="font-bold text-zinc-900 flex items-center gap-2">
                <Calculator className="w-4 h-4 text-orange-600" />
                Calculation Breakdown
              </h4>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase">
                      <th className="px-4 py-3 border border-zinc-200">Week</th>
                      <th className="px-4 py-3 border border-zinc-200">Tenure</th>
                      <th className="px-4 py-3 border border-zinc-200 text-center">% Target</th>
                      <th className="px-4 py-3 border border-zinc-200 text-right">Weekly Target</th>
                      <th className="px-4 py-3 border border-zinc-200 text-right">Sales Done</th>
                      <th className="px-4 py-3 border border-zinc-200 text-center">Status</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bravoCalculationModal.weeksData.filter(wk => !wk.isReference).map((wk, idx) => (
                      <tr key={idx} className="hover:bg-zinc-50 transition-colors">
                        <td className="px-4 py-3 text-sm border border-zinc-200 font-medium">{wk.weekName}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200">Wk {wk.tenureWeek}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200 text-center">{wk.percentage}%</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200 text-right">₹{wk.target.toLocaleString()}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200 text-right font-bold">₹{wk.sales.toLocaleString()}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200 text-center">
                          {wk.sales >= wk.target ? (
                            <CheckCircle2 className="w-4 h-4 text-green-600 mx-auto" />
                          ) : (
                            <AlertCircle className="w-4 h-4 text-zinc-300 mx-auto" />
                          )}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-zinc-100 font-bold">
                      <td colSpan={3} className="px-4 py-3 text-sm border border-zinc-200">TOTAL</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200 text-right">₹{bravoCalculationModal.totalTarget.toLocaleString()}</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200 text-right">₹{bravoCalculationModal.totalSales.toLocaleString()}</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200 text-center">
                        {bravoCalculationModal.achievement}%
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
              {bravoCalculationModal.weeksData.find(wk => wk.isReference) && (
                <div className="mt-2 p-2 px-4 bg-orange-50 border border-orange-100 rounded-lg text-orange-600 font-bold text-[11px] italic flex justify-between items-center">
                  <span>Reference Week: {bravoCalculationModal.weeksData.find(wk => wk.isReference)?.weekName} (Wk {bravoCalculationModal.weeksData.find(wk => wk.isReference)?.tenureWeek})</span>
                  <span>Target: ₹{bravoCalculationModal.weeksData.find(wk => wk.isReference)?.target.toLocaleString()} | Sales: ₹{bravoCalculationModal.weeksData.find(wk => wk.isReference)?.sales.toLocaleString()} (Not included in calculation)</span>
                </div>
              )}
            </div>

            <div className="p-4 bg-zinc-50 rounded-lg border border-zinc-200">
              <p className="text-xs font-bold text-zinc-500 uppercase mb-2 tracking-wider">Formula:</p>
              <p className="text-sm font-mono text-zinc-700 bg-white p-2 rounded border border-zinc-100">
                (Total Sales in last {bravoCalculationModal.numWeeks} weeks / Total Adjusted Target in last {bravoCalculationModal.numWeeks} weeks) * 100
              </p>
              <p className="text-[10px] text-zinc-400 mt-2 italic">
                * Adjusted target = (Base Weekly Target * Tenure Percentage) / 100
              </p>
            </div>
          </div>
        </Modal>

        <Modal
          isOpen={contributionCalculationModal.isOpen}
          onClose={() => setContributionCalculationModal({ ...contributionCalculationModal, isOpen: false })}
          title={contributionCalculationModal.title}
        >
          <div className="max-h-[70vh] overflow-auto p-4 space-y-6">
            <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
              <Card className="p-4 bg-emerald-50 border-emerald-100">
                <p className="text-[10px] font-bold text-emerald-600 uppercase tracking-wider mb-1">Total Contribution</p>
                <p className={cn("text-xl font-bold", contributionCalculationModal.totalContribution >= 0 ? "text-emerald-900" : "text-red-900")}>
                  ₹{contributionCalculationModal.totalContribution.toLocaleString()}
                </p>
              </Card>
              <Card className="p-4 bg-zinc-50 border-zinc-100">
                <p className="text-[10px] font-bold text-zinc-600 uppercase tracking-wider mb-1">Period</p>
                <p className="text-xl font-bold text-zinc-900">{contributionCalculationModal.numWeeks} Weeks</p>
              </Card>
            </div>

            <div className="space-y-3">
              <h4 className="font-bold text-zinc-900 flex items-center gap-2">
                <Calculator className="w-4 h-4 text-emerald-600" />
                Contribution Breakdown
              </h4>
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase">
                      <th className="px-4 py-3 border border-zinc-200">Week</th>
                      <th className="px-4 py-3 border border-zinc-200 text-right">Sales (Net)</th>
                      <th className="px-4 py-3 border border-zinc-200 text-right">Workstation Cost</th>
                      <th className="px-4 py-3 border border-zinc-200 text-right">Contribution</th>
                    </tr>
                  </thead>
                  <tbody>
                    {contributionCalculationModal.weeksData.map((wk, idx) => (
                      <tr key={idx} className="hover:bg-zinc-50 transition-colors">
                        <td className="px-4 py-3 text-sm border border-zinc-200 font-medium">{wk.weekName}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200 text-right">₹{wk.sales.toLocaleString()}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200 text-right">₹{wk.workstationCost.toLocaleString()}</td>
                        <td className={cn("px-4 py-3 text-sm border border-zinc-200 text-right font-bold", wk.contribution >= 0 ? "text-emerald-600" : "text-red-600")}>
                          ₹{wk.contribution.toLocaleString()}
                        </td>
                      </tr>
                    ))}
                  </tbody>
                  <tfoot>
                    <tr className="bg-zinc-100 font-bold">
                      <td colSpan={3} className="px-4 py-3 text-sm border border-zinc-200">TOTAL CONTRIBUTION</td>
                      <td className={cn("px-4 py-3 text-sm border border-zinc-200 text-right", contributionCalculationModal.totalContribution >= 0 ? "text-emerald-700" : "text-red-700")}>
                        ₹{contributionCalculationModal.totalContribution.toLocaleString()}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>

            <div className="p-4 bg-zinc-50 rounded-lg border border-zinc-200">
              <p className="text-xs font-bold text-zinc-500 uppercase mb-2 tracking-wider">Formula:</p>
              <p className="text-sm font-mono text-zinc-700 bg-white p-2 rounded border border-zinc-100">
                Contribution = (Net Sales * 12%) - Workstation Cost
              </p>
              <p className="text-[10px] text-zinc-400 mt-2 italic">
                * Workstation Cost = (Salary * 3) / 4 (per week)
                <br />
                * Net Sales excludes Cancellations and Credit Notes.
              </p>
            </div>
          </div>
        </Modal>
        {/* Incentive Master Page */}
      {activeTab === 'incentiveMaster' && hasPermission('incentiveMaster') && (
        <div className="space-y-8">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <IndianRupee className="w-6 h-6 text-orange-600" />
              <h2 className="text-3xl font-bold tracking-tight">Incentive Master</h2>
            </div>
            <div className="flex items-center gap-3">
              <Button onClick={() => setIsIncentiveModalOpen(true)} className="bg-orange-600">
                <Plus className="w-4 h-4" />
                Add Incentive Due
              </Button>
              <Button onClick={() => setIsPaymentModalOpen(true)} variant="outline">
                <Plus className="w-4 h-4" />
                Record Payment
              </Button>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
            {/* Summary Cards */}
            <div className="lg:col-span-4 grid grid-cols-1 md:grid-cols-3 gap-6">
              <Card 
                className="p-6 bg-orange-600 text-white cursor-pointer hover:bg-orange-700 transition-colors"
                onClick={() => setIsOutstandingDuesModalOpen(true)}
              >
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-orange-100 text-sm font-bold uppercase tracking-wider">Total Approx Outstanding Dues</p>
                    <h3 className="text-3xl font-black mt-1">
                      ₹{(() => {
                        const totalDue = incentives.reduce((sum, i) => sum + i.amount, 0);
                        const totalPaid = incentivePayments.reduce((sum, p) => sum + p.amount, 0);
                        return (totalDue - totalPaid).toLocaleString();
                      })()}
                    </h3>
                  </div>
                  <div className="w-12 h-12 bg-white/20 rounded-2xl flex items-center justify-center">
                    <TrendingUp className="w-6 h-6 text-white" />
                  </div>
                </div>
                <p className="text-orange-100 text-xs mt-4 flex items-center gap-1">
                  <Info className="w-3 h-3" />
                  Click to view detailed list of all employees
                </p>
              </Card>

              <Card className="p-6 bg-zinc-900 text-white">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-zinc-400 text-sm font-bold uppercase tracking-wider">Total Approx Incentives Due</p>
                    <h3 className="text-3xl font-black mt-1">
                      ₹{incentives.reduce((sum, i) => sum + i.amount, 0).toLocaleString()}
                    </h3>
                  </div>
                  <div className="w-12 h-12 bg-white/10 rounded-2xl flex items-center justify-center text-red-400">
                    <ArrowUpRight className="w-6 h-6" />
                  </div>
                </div>
              </Card>

              <Card className="p-6 bg-zinc-900 text-white">
                <div className="flex items-center justify-between">
                  <div>
                    <p className="text-zinc-400 text-sm font-bold uppercase tracking-wider">Total Approx Paid</p>
                    <h3 className="text-3xl font-black mt-1">
                      ₹{incentivePayments.reduce((sum, p) => sum + p.amount, 0).toLocaleString()}
                    </h3>
                  </div>
                  <div className="w-12 h-12 bg-white/10 rounded-2xl flex items-center justify-center text-green-400">
                    <ArrowDownRight className="w-6 h-6" />
                  </div>
                </div>
              </Card>
            </div>

            {/* Selected Employee Summary Table */}
            {ledgerEmployeeFilter && ledgerEmployeeFilter !== 'all' && (
              <div className="lg:col-span-4">
                <Card className="overflow-hidden border-zinc-200">
                  <div className="bg-zinc-50 px-4 py-2 border-b border-zinc-200 flex justify-between items-center">
                    <h4 className="text-xs font-bold text-zinc-500 uppercase">Employee Summary (Incentive + Quaterly) (Approx)</h4>
                    <div className="flex items-center gap-4 text-[10px] font-bold uppercase tracking-wider">
                      <div className="flex items-center gap-1.5">
                        <div className="w-3 h-3 bg-white border border-zinc-300 rounded-sm"></div>
                        <span className="text-zinc-500">White - Due</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <div className="w-3 h-3 bg-purple-100 border border-purple-200 rounded-sm"></div>
                        <span className="text-purple-600">Purple - Eligible to be paid</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <div className="w-3 h-3 bg-green-100 border border-green-200 rounded-sm"></div>
                        <span className="text-green-600">Green - Paid</span>
                      </div>
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-white text-[10px] font-bold text-zinc-400 uppercase">
                          <th className="px-4 py-2 border-r">Month</th>
                          <th className="px-4 py-2 border-r">Daily</th>
                          <th className="px-4 py-2 border-r">Weekly</th>
                          <th className="px-4 py-2 border-r">Monthly</th>
                          <th className="px-4 py-2 border-r">Quarterly</th>
                          <th className="px-4 py-2 border-r">Annually</th>
                          <th className="px-4 py-2 border-r">Total Approx Due</th>
                          <th className="px-4 py-2 border-r">Total Approx Paid</th>
                          <th className="px-4 py-2 bg-orange-50 text-orange-600">Approx Balance</th>
                        </tr>
                      </thead>
                      <tbody>
                        {monthlyIncentiveData.map((row: any) => (
                          <tr key={row.month} className="text-sm font-medium border-b border-zinc-100 hover:bg-zinc-50 transition-colors text-zinc-600">
                            <td className="px-4 py-2 border-r font-bold text-zinc-900">{row.month}</td>
                            <td className={cn("px-4 py-2 border-r", row.typeStatuses.Daily.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Daily.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Daily.toLocaleString()}</td>
                            <td className={cn("px-4 py-2 border-r", row.typeStatuses.Weekly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Weekly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Weekly.toLocaleString()}</td>
                            <td className={cn("px-4 py-2 border-r", row.typeStatuses.Monthly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Monthly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Monthly.toLocaleString()}</td>
                            <td className={cn("px-4 py-2 border-r", row.typeStatuses.Quarterly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Quarterly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Quarterly.toLocaleString()}</td>
                            <td className={cn("px-4 py-2 border-r", row.typeStatuses.Annually.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Annually.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Annually.toLocaleString()}</td>
                            <td 
                              className="px-4 py-2 border-r text-red-600 cursor-pointer hover:bg-red-50 transition-colors font-bold"
                              onClick={() => setAgentLedgerModal({ isOpen: true, employeeId: ledgerEmployeeFilter })}
                            >
                              ₹{row.totalDue.toLocaleString()}
                            </td>
                            <td className="px-4 py-2 border-r text-green-600 font-bold">₹{row.totalPaid.toLocaleString()}</td>
                            <td className="px-4 py-2 bg-orange-50 text-orange-600 font-bold">₹{(row.totalDue - row.totalPaid).toLocaleString()}</td>
                          </tr>
                        ))}
                        <tr className="text-sm font-black text-zinc-900 bg-zinc-50">
                          <td className="px-4 py-3 border-r uppercase">Grand Total</td>
                          <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === ledgerEmployeeFilter && i.type === 'Daily').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                          <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === ledgerEmployeeFilter && i.type === 'Weekly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                          <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === ledgerEmployeeFilter && i.type === 'Monthly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                          <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === ledgerEmployeeFilter && i.type === 'Quarterly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                          <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === ledgerEmployeeFilter && i.type === 'Annually').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                          <td 
                            className="px-4 py-3 border-r text-red-600 cursor-pointer hover:bg-red-100 transition-colors font-bold"
                            onClick={() => setAgentLedgerModal({ isOpen: true, employeeId: ledgerEmployeeFilter })}
                          >
                            ₹{incentives.filter(i => i.employeeId === ledgerEmployeeFilter).reduce((sum, i) => sum + i.amount, 0).toLocaleString()}
                          </td>
                          <td className="px-4 py-3 border-r text-green-600 font-bold">₹{incentivePayments.filter(p => p.employeeId === ledgerEmployeeFilter).reduce((sum, p) => sum + p.amount, 0).toLocaleString()}</td>
                          <td className="px-4 py-3 bg-orange-100 text-orange-700 text-lg font-bold">
                            ₹{(
                              incentives.filter(i => i.employeeId === ledgerEmployeeFilter).reduce((sum, i) => sum + i.amount, 0) - 
                              incentivePayments.filter(p => p.employeeId === ledgerEmployeeFilter).reduce((sum, p) => sum + p.amount, 0)
                            ).toLocaleString()}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </Card>
              </div>
            )}

            {/* All Employees Summary Table */}
            {ledgerEmployeeFilter === 'all' && (
              <div className="lg:col-span-4">
                <Card className="overflow-hidden border-zinc-200">
                  <div className="bg-zinc-50 px-4 py-2 border-b border-zinc-200 flex justify-between items-center">
                    <h4 className="text-xs font-bold text-zinc-500 uppercase">All Employees Incentive Summary</h4>
                    <div className="flex items-center gap-4 text-[10px] font-bold uppercase tracking-wider">
                      <div className="flex items-center gap-1.5">
                        <div className="w-3 h-3 bg-white border border-zinc-300 rounded-sm"></div>
                        <span className="text-zinc-500">White - Due</span>
                      </div>
                      <div className="flex items-center gap-1.5">
                        <div className="w-3 h-3 bg-purple-100 border border-purple-200 rounded-sm"></div>
                        <span className="text-purple-600">Purple - Eligible</span>
                      </div>
                    </div>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-zinc-100 text-[10px] font-bold text-zinc-500 uppercase">
                          <th rowSpan={2} className="px-4 py-2 border-r border-b">Employees</th>
                          <th colSpan={2} className="px-4 py-1 border-r border-b text-center">Daily</th>
                          <th colSpan={2} className="px-4 py-1 border-r border-b text-center">Weekly</th>
                          <th colSpan={2} className="px-4 py-1 border-r border-b text-center">Monthly</th>
                          <th colSpan={2} className="px-4 py-1 border-r border-b text-center">Quarterly</th>
                          <th colSpan={2} className="px-4 py-1 border-r border-b text-center">Annually</th>
                          <th colSpan={2} className="px-4 py-1 border-r border-b text-center">Total Approx Due</th>
                          <th rowSpan={2} className="px-4 py-2 border-r border-b">Total Approx Paid</th>
                          <th rowSpan={2} className="px-4 py-2 border-b">Approx Balance</th>
                        </tr>
                        <tr className="bg-zinc-50 text-[9px] font-bold text-zinc-400 uppercase">
                          <th className="px-2 py-1 border-r border-b">Due</th>
                          <th className="px-2 py-1 border-r border-b">Eligible</th>
                          <th className="px-2 py-1 border-r border-b">Due</th>
                          <th className="px-2 py-1 border-r border-b">Eligible</th>
                          <th className="px-2 py-1 border-r border-b">Due</th>
                          <th className="px-2 py-1 border-r border-b">Eligible</th>
                          <th className="px-2 py-1 border-r border-b">Due</th>
                          <th className="px-2 py-1 border-r border-b">Eligible</th>
                          <th className="px-2 py-1 border-r border-b">Due</th>
                          <th className="px-2 py-1 border-r border-b">Eligible</th>
                          <th className="px-2 py-1 border-r border-b">Due</th>
                          <th className="px-2 py-1 border-r border-b">Eligible</th>
                        </tr>
                      </thead>
                      <tbody>
                        {allEmployeesIncentiveSummary.map((row: any) => (
                          <tr key={row.employeeId} className="text-xs font-medium border-b border-zinc-100 hover:bg-zinc-50 transition-colors text-zinc-600">
                            <td 
                              className="px-4 py-2 border-r font-bold text-zinc-900 cursor-pointer hover:text-orange-600 transition-colors"
                              onClick={() => setAgentLedgerModal({ isOpen: true, employeeId: row.employeeId })}
                            >
                              {row.employeeName}
                            </td>
                            <td className="px-2 py-2 border-r">₹{row.Daily.due.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r bg-purple-100 text-purple-700">₹{row.Daily.eligible.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r">₹{row.Weekly.due.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r bg-purple-100 text-purple-700">₹{row.Weekly.eligible.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r">₹{row.Monthly.due.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r bg-purple-100 text-purple-700">₹{row.Monthly.eligible.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r">₹{row.Quarterly.due.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r bg-purple-100 text-purple-700">₹{row.Quarterly.eligible.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r">₹{row.Annually.due.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r bg-purple-100 text-purple-700">₹{row.Annually.eligible.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r text-red-600 font-bold">₹{row.totalDue.due.toLocaleString()}</td>
                            <td className="px-2 py-2 border-r bg-purple-100 text-purple-700 font-bold">₹{row.totalDue.eligible.toLocaleString()}</td>
                            <td className="px-4 py-2 border-r text-green-600 font-bold">₹{row.totalPaid.toLocaleString()}</td>
                            <td className="px-4 py-2 bg-orange-50 text-orange-600 font-bold">₹{row.balance.toLocaleString()}</td>
                          </tr>
                        ))}
                        <tr className="text-xs font-black text-zinc-900 bg-zinc-100">
                          <td className="px-4 py-3 border-r uppercase">Grand Total</td>
                          <td className="px-2 py-3 border-r">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Daily.due, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r bg-purple-100 text-purple-800">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Daily.eligible, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Weekly.due, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r bg-purple-100 text-purple-800">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Weekly.eligible, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Monthly.due, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r bg-purple-100 text-purple-800">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Monthly.eligible, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Quarterly.due, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r bg-purple-100 text-purple-800">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Quarterly.eligible, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Annually.due, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r bg-purple-100 text-purple-800">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.Annually.eligible, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r text-red-600">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.totalDue.due, 0).toLocaleString()}</td>
                          <td className="px-2 py-3 border-r bg-purple-100 text-purple-800">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.totalDue.eligible, 0).toLocaleString()}</td>
                          <td className="px-4 py-3 border-r text-green-600">₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.totalPaid, 0).toLocaleString()}</td>
                          <td className="px-4 py-3 bg-orange-100 text-orange-700 text-sm font-bold">
                            ₹{allEmployeesIncentiveSummary.reduce((sum, r) => sum + r.balance, 0).toLocaleString()}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>
                </Card>
              </div>
            )}

            {/* Ledger Section */}
            <div className="lg:col-span-4 space-y-6">
              <Card className="p-6">
                <div className="flex items-center justify-between mb-6">
                  <div className="space-y-1">
                    <h3 className="text-lg font-bold">Employee Incentive Ledger</h3>
                    <div className="flex flex-col gap-2">
                      {ledgerEmployeeFilter && (
                        <div className="flex items-center gap-2">
                          <span className="text-xs font-bold text-zinc-500 uppercase">Final Approx Due:</span>
                          <span className="text-sm font-black text-orange-600">
                            ₹{(() => {
                              const totalDue = incentives.filter(i => i.employeeId === ledgerEmployeeFilter).reduce((sum, i) => sum + i.amount, 0);
                              const totalPaid = incentivePayments.filter(p => p.employeeId === ledgerEmployeeFilter).reduce((sum, p) => sum + p.amount, 0);
                              return (totalDue - totalPaid).toLocaleString();
                            })()}
                          </span>
                        </div>
                      )}
                      <div className="flex items-center gap-4 text-[10px] font-bold uppercase tracking-wider">
                        <div className="flex items-center gap-1.5">
                          <div className="w-3 h-3 rounded border border-zinc-200 bg-white"></div>
                          <span className="text-zinc-500">Due</span>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <div className="w-3 h-3 rounded bg-purple-100 border border-purple-200"></div>
                          <span className="text-purple-600">Eligible to be paid</span>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <div className="w-3 h-3 rounded bg-green-100 border border-green-200"></div>
                          <span className="text-green-600">Paid</span>
                        </div>
                      </div>
                    </div>
                  </div>
                  <div className="w-64">
                    <SearchableSelect
                      options={[
                        { value: '', label: 'Select Employee' },
                        { value: 'all', label: 'All Employees' },
                        ...employees
                          .sort((a, b) => a.name.localeCompare(b.name))
                          .map(emp => ({ value: emp.id, label: emp.name, key: emp.id }))
                      ]}
                      value={ledgerEmployeeFilter}
                      onChange={setLedgerEmployeeFilter}
                      placeholder="Select Employee"
                    />
                  </div>
                </div>

                {ledgerEmployeeFilter && ledgerEmployeeFilter !== 'all' ? (
                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse">
                      <thead>
                        <tr className="bg-zinc-50 text-xs font-bold text-zinc-500 uppercase">
                          <th className="px-4 py-3 border-b">Date</th>
                          <th className="px-4 py-3 border-b">Description</th>
                          <th className="px-4 py-3 border-b text-right">Approx Due (₹)</th>
                          <th className="px-4 py-3 border-b text-right">Approx Paid (₹)</th>
                          <th className="px-4 py-3 border-b text-right">Approx Balance (₹)</th>
                          <th className="px-4 py-3 border-b text-center">Actions</th>
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          const empIncentives = incentives
                            .filter(i => i.employeeId === ledgerEmployeeFilter)
                            .map(i => ({ ...i, entryType: 'due' as const }));
                          const empPayments = incentivePayments
                            .filter(p => p.employeeId === ledgerEmployeeFilter)
                            .map(p => ({ ...p, entryType: 'paid' as const }));
                          
                          const allEntries = [...empIncentives, ...empPayments].sort((a, b) => 
                            new Date(a.date).getTime() - new Date(b.date).getTime()
                          );

                          let runningBalance = 0;
                          return allEntries.map((entry, idx) => {
                            const due = entry.entryType === 'due' ? (entry as Incentive).amount : 0;
                            const paid = entry.entryType === 'paid' ? (entry as IncentivePayment).amount : 0;
                            runningBalance += due - paid;

                            return (
                              <tr key={entry.id || idx} className="hover:bg-zinc-50 transition-colors border-b last:border-0 group">
                                <td className="px-4 py-3 text-sm text-zinc-600">
                                  {entry.date}
                                  {entry.recordedAt && (
                                    <p className="text-[10px] text-zinc-400 mt-0.5">
                                      {entry.recordedAt}
                                    </p>
                                  )}
                                </td>
                                <td className="px-4 py-3 text-sm">
                                  <div className="font-medium text-zinc-900">
                                    {entry.entryType === 'due' ? (entry as Incentive).type : 'Payment'}
                                  </div>
                                  <div className="flex flex-col gap-0.5">
                                    {entry.remarks && <p className="text-xs text-zinc-400 italic">{entry.remarks}</p>}
                                    {entry.recordedBy && (
                                      <p className="text-[10px] text-zinc-500 font-medium">
                                        By: {entry.recordedBy}
                                      </p>
                                    )}
                                  </div>
                                </td>
                                <td className={cn(
                                  "px-4 py-3 text-sm text-right font-medium transition-colors",
                                  due > 0 && (entry as Incentive).isPaid ? "bg-green-100 text-green-700" : 
                                  due > 0 && (entry as Incentive).isEligible ? "bg-purple-100 text-purple-700" : 
                                  due > 0 ? "text-red-600" : "text-zinc-400"
                                )}>
                                  {due > 0 ? (
                                    <div className="flex flex-col items-end gap-1">
                                      <span>₹{due.toLocaleString()}</span>
                                      <div className="flex gap-1">
                                        {(entry as Incentive).isEligible && (
                                          <span className="px-1.5 py-0.5 bg-purple-50 text-purple-600 text-[9px] font-bold rounded border border-purple-100 uppercase tracking-tighter">
                                            Eligible
                                          </span>
                                        )}
                                        {(entry as Incentive).isPaid && (
                                          <span className="px-1.5 py-0.5 bg-green-50 text-green-600 text-[9px] font-bold rounded border border-green-100 uppercase tracking-tighter">
                                            Paid
                                          </span>
                                        )}
                                      </div>
                                    </div>
                                  ) : '-'}
                                </td>
                                <td className="px-4 py-3 text-sm text-right font-medium text-green-600">
                                  {paid > 0 ? `₹${paid.toLocaleString()}` : '-'}
                                </td>
                                <td className="px-4 py-3 text-sm text-right font-bold text-zinc-900">
                                  ₹{runningBalance.toLocaleString()}
                                </td>
                                <td className="px-4 py-3 text-center">
                                  <button
                                    onClick={() => entry.entryType === 'due' ? handleEditIncentive(entry as Incentive) : handleEditPayment(entry as IncentivePayment)}
                                    className="p-1.5 text-zinc-400 hover:text-orange-600 hover:bg-orange-50 rounded-lg transition-all opacity-0 group-hover:opacity-100"
                                  >
                                    <Edit2 className="w-4 h-4" />
                                  </button>
                                </td>
                              </tr>
                            );
                          });
                        })()}
                        {incentives.filter(i => i.employeeId === ledgerEmployeeFilter).length === 0 && 
                         incentivePayments.filter(p => p.employeeId === ledgerEmployeeFilter).length === 0 && (
                          <tr>
                            <td colSpan={6} className="px-4 py-12 text-center text-zinc-400 italic">
                              No transactions found for this employee.
                            </td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="py-20 text-center space-y-4">
                    <div className="w-16 h-16 bg-zinc-100 rounded-full flex items-center justify-center mx-auto">
                      <Users className="w-8 h-8 text-zinc-400" />
                    </div>
                    <p className="text-zinc-500 font-medium">Select an employee to view their incentive ledger.</p>
                  </div>
                )}
              </Card>
            </div>
          </div>

          <Modal
            isOpen={isOutstandingDuesModalOpen}
            onClose={() => setIsOutstandingDuesModalOpen(false)}
            title="All Employees Approx Outstanding Dues"
          >
            <div className="space-y-4 max-h-[70vh] overflow-y-auto pr-2">
              <table className="w-full text-left border-collapse">
                <thead className="sticky top-0 bg-white shadow-sm z-10">
                  <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase">
                    <th className="px-4 py-3 border-b">Employee Name</th>
                    <th className="px-4 py-3 border-b text-right">Daily</th>
                    <th className="px-4 py-3 border-b text-right">Weekly</th>
                    <th className="px-4 py-3 border-b text-right">Monthly</th>
                    <th className="px-4 py-3 border-b text-right">Quarterly</th>
                    <th className="px-4 py-3 border-b text-right">Annually</th>
                    <th className="px-4 py-3 border-b text-right">Total Approx Due</th>
                    <th className="px-4 py-3 border-b text-right">Approx Paid</th>
                    <th className="px-4 py-3 border-b text-right bg-orange-50 text-orange-600 font-black">Approx Balance</th>
                  </tr>
                </thead>
                <tbody>
                  {employees
                    .filter(emp => emp.status === 'Active')
                    .map(emp => {
                      const empIncentives = incentives.filter(i => i.employeeId === emp.id);
                      const daily = empIncentives.filter(i => i.type === 'Daily').reduce((sum, i) => sum + i.amount, 0);
                      const weekly = empIncentives.filter(i => i.type === 'Weekly').reduce((sum, i) => sum + i.amount, 0);
                      const monthly = empIncentives.filter(i => i.type === 'Monthly').reduce((sum, i) => sum + i.amount, 0);
                      const quarterly = empIncentives.filter(i => i.type === 'Quarterly').reduce((sum, i) => sum + i.amount, 0);
                      const annually = empIncentives.filter(i => i.type === 'Annually').reduce((sum, i) => sum + i.amount, 0);
                      
                      const totalDue = empIncentives.reduce((sum, i) => sum + i.amount, 0);
                      const totalPaid = incentivePayments
                        .filter(p => p.employeeId === emp.id)
                        .reduce((sum, p) => sum + p.amount, 0);
                      const balance = totalDue - totalPaid;

                      if (totalDue === 0 && totalPaid === 0) return null;

                      return (
                        <tr 
                          key={emp.id} 
                          className="hover:bg-zinc-50 transition-colors border-b last:border-0 cursor-pointer group"
                          onClick={() => {
                            setLedgerEmployeeFilter(emp.id!);
                            setIsOutstandingDuesModalOpen(false);
                          }}
                        >
                          <td className="px-4 py-3 text-sm">
                            <div className="font-bold text-zinc-900 group-hover:text-orange-600 transition-colors">
                              {emp.name}
                            </div>
                            <p className="text-[10px] text-zinc-500 uppercase tracking-wider">{emp.employeeCode}</p>
                          </td>
                          <td className="px-4 py-3 text-xs text-right text-zinc-600">₹{daily.toLocaleString()}</td>
                          <td className="px-4 py-3 text-xs text-right text-zinc-600">₹{weekly.toLocaleString()}</td>
                          <td className="px-4 py-3 text-xs text-right text-zinc-600">₹{monthly.toLocaleString()}</td>
                          <td className="px-4 py-3 text-xs text-right text-zinc-600">₹{quarterly.toLocaleString()}</td>
                          <td className="px-4 py-3 text-xs text-right text-zinc-600">₹{annually.toLocaleString()}</td>
                          <td className="px-4 py-3 text-sm text-right text-red-600 font-bold">
                            ₹{totalDue.toLocaleString()}
                          </td>
                          <td className="px-4 py-3 text-sm text-right text-green-600 font-bold">
                            ₹{totalPaid.toLocaleString()}
                          </td>
                          <td className="px-4 py-3 text-sm text-right font-bold bg-orange-50 text-orange-600">
                            ₹{balance.toLocaleString()}
                          </td>
                        </tr>
                      );
                    })
                    .filter(Boolean)}
                </tbody>
              </table>
            </div>
          </Modal>

          {/* Add Incentive Modal */}
          <Modal
            isOpen={isIncentiveModalOpen}
            onClose={() => {
              setIsIncentiveModalOpen(false);
              setIsEditingIncentive(false);
              setSelectedIncentive(null);
              setIncentiveForm({
                type: 'Daily',
                date: new Date().toISOString().split('T')[0],
                amount: 0,
                remarks: '',
                isEligible: false,
                isPaid: false
              });
            }}
            title={isEditingIncentive ? "Edit Incentive Due" : "Add Incentive Due"}
          >
            <form onSubmit={handleCreateIncentive} className="space-y-6">
              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">Employee / BDE</label>
                <SearchableSelect
                  required
                  options={[
                    { value: '', label: 'Select Employee' },
                    ...employees
                      .sort((a, b) => a.name.localeCompare(b.name))
                      .map(emp => ({ value: emp.id, label: emp.name, key: emp.id }))
                  ]}
                  value={incentiveForm.employeeId || ''}
                  onChange={(val) => setIncentiveForm({ ...incentiveForm, employeeId: val })}
                  placeholder="Select Employee"
                  disabled={isEditingIncentive}
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Incentive Type</label>
                  <select
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={incentiveForm.type || 'Daily'}
                    onChange={(e) => setIncentiveForm({ ...incentiveForm, type: e.target.value as any })}
                  >
                    <option value="Daily">Daily</option>
                    <option value="Weekly">Weekly</option>
                    <option value="Monthly">Monthly</option>
                    <option value="Quarterly">Quarterly</option>
                    <option value="Annually">Annually</option>
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Date</label>
                  <input
                    type="date"
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={incentiveForm.date || ''}
                    onChange={(e) => setIncentiveForm({ ...incentiveForm, date: e.target.value })}
                  />
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">Approx Amount (₹)</label>
                <input
                  type="number"
                  required
                  className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                  value={incentiveForm.amount || ''}
                  onChange={(e) => setIncentiveForm({ ...incentiveForm, amount: Number(e.target.value) })}
                  onFocus={(e) => e.target.select()}
                />
              </div>

              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">Remarks</label>
                <textarea
                  className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none resize-none"
                  rows={3}
                  value={incentiveForm.remarks || ''}
                  onChange={(e) => setIncentiveForm({ ...incentiveForm, remarks: e.target.value })}
                />
              </div>

              <div className="grid grid-cols-2 gap-4 bg-zinc-50 p-4 rounded-xl border border-zinc-200">
                <label className="flex items-center gap-3 cursor-pointer group">
                  <div className="relative flex items-center">
                    <input
                      type="checkbox"
                      className="w-5 h-5 rounded border-zinc-300 text-orange-600 focus:ring-orange-500 cursor-pointer transition-all"
                      checked={incentiveForm.isEligible || false}
                      onChange={(e) => setIncentiveForm({ ...incentiveForm, isEligible: e.target.checked })}
                    />
                  </div>
                  <span className="text-sm font-bold text-zinc-700 group-hover:text-orange-600 transition-colors">Eligibility</span>
                </label>
                <label className="flex items-center gap-3 cursor-pointer group">
                  <div className="relative flex items-center">
                    <input
                      type="checkbox"
                      className="w-5 h-5 rounded border-zinc-300 text-orange-600 focus:ring-orange-500 cursor-pointer transition-all"
                      checked={incentiveForm.isPaid || false}
                      onChange={(e) => setIncentiveForm({ ...incentiveForm, isPaid: e.target.checked })}
                    />
                  </div>
                  <span className="text-sm font-bold text-zinc-700 group-hover:text-orange-600 transition-colors">Paid</span>
                </label>
              </div>

              <div className="flex gap-3 pt-4">
                <Button type="submit" className="flex-1" isLoading={isSavingIncentive}>
                  {isEditingIncentive ? "Update Incentive" : "Save Incentive Due"}
                </Button>
                <Button type="button" variant="ghost" onClick={() => {
                  setIsIncentiveModalOpen(false);
                  setIsEditingIncentive(false);
                  setSelectedIncentive(null);
                  setIncentiveForm({
                    type: 'Daily',
                    date: new Date().toISOString().split('T')[0],
                    amount: 0,
                    remarks: '',
                    isEligible: false,
                    isPaid: false
                  });
                }}>
                  Cancel
                </Button>
              </div>
            </form>
          </Modal>



          {/* Record Payment Modal */}
          <Modal
            isOpen={isPaymentModalOpen}
            onClose={() => {
              setIsPaymentModalOpen(false);
              setIsEditingPayment(false);
              setSelectedPayment(null);
              setPaymentForm({
                date: new Date().toISOString().split('T')[0],
                amount: 0,
                paymentMethod: 'Cash',
                remarks: ''
              });
            }}
            title={isEditingPayment ? "Edit Approx Incentive Payment" : "Record Approx Incentive Payment"}
          >
            <form onSubmit={handleCreatePayment} className="space-y-6">
              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">Employee / BDE</label>
                <SearchableSelect
                  required
                  options={[
                    { value: '', label: 'Select Employee' },
                    ...employees
                      .sort((a, b) => a.name.localeCompare(b.name))
                      .map(emp => ({ value: emp.id, label: emp.name, key: emp.id }))
                  ]}
                  value={paymentForm.employeeId || ''}
                  onChange={(val) => setPaymentForm({ ...paymentForm, employeeId: val })}
                  placeholder="Select Employee"
                  disabled={isEditingPayment}
                />
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Approx Payment Method</label>
                  <select
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={paymentForm.paymentMethod || 'Cash'}
                    onChange={(e) => setPaymentForm({ ...paymentForm, paymentMethod: e.target.value })}
                  >
                    <option value="Cash">Cash</option>
                    <option value="Bank Transfer">Bank Transfer</option>
                    <option value="UPI">UPI</option>
                    <option value="Cheque">Cheque</option>
                  </select>
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Approx Date</label>
                  <input
                    type="date"
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={paymentForm.date || ''}
                    onChange={(e) => setPaymentForm({ ...paymentForm, date: e.target.value })}
                  />
                </div>
              </div>

              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">Approx Amount (₹)</label>
                <input
                  type="number"
                  required
                  className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                  value={paymentForm.amount || ''}
                  onChange={(e) => setPaymentForm({ ...paymentForm, amount: Number(e.target.value) })}
                  onFocus={(e) => e.target.select()}
                />
              </div>

              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">Approx Remarks</label>
                <textarea
                  className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none resize-none"
                  rows={3}
                  value={paymentForm.remarks || ''}
                  onChange={(e) => setPaymentForm({ ...paymentForm, remarks: e.target.value })}
                />
              </div>

              <div className="flex gap-3 pt-4">
                <Button type="submit" className="flex-1" isLoading={isSavingPayment}>
                  {isEditingPayment ? "Update Approx Payment" : "Record Approx Payment"}
                </Button>
                <Button type="button" variant="ghost" onClick={() => {
                  setIsPaymentModalOpen(false);
                  setIsEditingPayment(false);
                  setSelectedPayment(null);
                  setPaymentForm({
                    date: new Date().toISOString().split('T')[0],
                    amount: 0,
                    paymentMethod: 'Cash',
                    remarks: ''
                  });
                }}>
                  Cancel
                </Button>
              </div>
            </form>
          </Modal>
        </div>
      )}

      {/* Matrix Master Page */}
      {activeTab === 'matrixMaster' && hasPermission('matrixMaster') && (
        <div className="space-y-8">
          <div className="flex items-center justify-between">
            <div className="flex items-center gap-2">
              <LayoutGrid className="w-6 h-6 text-orange-600" />
              <h2 className="text-3xl font-bold tracking-tight">Matrix Master</h2>
            </div>
            <div className="flex items-center gap-3">
              <div className="flex bg-zinc-100 p-1 rounded-lg mr-4">
                <button
                  onClick={() => setMatrixView('reports')}
                  className={cn(
                    "px-4 py-1.5 text-xs font-bold rounded-md transition-all flex items-center gap-2",
                    matrixView === 'reports' 
                      ? "bg-white text-orange-600 shadow-sm" 
                      : "text-zinc-500 hover:text-zinc-700"
                  )}
                >
                  <LayoutGrid className="w-3.5 h-3.5" />
                  Daily Reports
                </button>
                <button
                  onClick={() => setMatrixView('trips')}
                  className={cn(
                    "px-4 py-1.5 text-xs font-bold rounded-md transition-all flex items-center gap-2",
                    matrixView === 'trips' 
                      ? "bg-white text-orange-600 shadow-sm" 
                      : "text-zinc-500 hover:text-zinc-700"
                  )}
                >
                  <PhoneCall className="w-3.5 h-3.5" />
                  Trip ID Summary
                </button>
              </div>
              <Button onClick={() => {
                setMatrixEditId(null);
                setMatrixForm({
                  date: format(new Date(), 'yyyy-MM-dd'),
                  performancePoints: 0,
                  totalCalls: 0,
                  totalTalktime: '',
                  topTalktimes: [
                    { tripId: '', phoneNo: '', talktime: '' },
                    { tripId: '', phoneNo: '', talktime: '' },
                    { tripId: '', phoneNo: '', talktime: '' },
                    { tripId: '', phoneNo: '', talktime: '' },
                  ],
                  loginTime: '',
                  logoutTime: '',
                  breakTime: '',
                  employeeId: !isAdmin && !isBDEUser ? currentEmployee?.id : '',
                  employeeName: !isAdmin && !isBDEUser ? currentEmployee?.name : '',
                });
                setError(null);
                setSuccess(null);
                setIsMatrixModalOpen(true);
              }} className="bg-orange-600">
                <Plus className="w-4 h-4" />
                Submit Daily Report
              </Button>
            </div>
          </div>

          {/* Filters */}
          <div className="space-y-4">
            <div className="flex items-center gap-2">
              <Button 
                variant="outline" 
                size="sm" 
                onClick={() => {
                  const today = format(new Date(), 'yyyy-MM-dd');
                  setMatrixStartDate(today);
                  setMatrixEndDate(today);
                }}
                className="text-xs h-8"
              >
                Today
              </Button>
              <Button 
                variant="outline" 
                size="sm" 
                onClick={() => {
                  const yesterday = format(subDays(new Date(), 1), 'yyyy-MM-dd');
                  setMatrixStartDate(yesterday);
                  setMatrixEndDate(yesterday);
                }}
                className="text-xs h-8"
              >
                Yesterday
              </Button>
            </div>
            <div className="grid grid-cols-1 md:grid-cols-4 gap-4">
              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">Start Date</label>
                <input
                  type="date"
                  className="w-full px-4 py-2 bg-white border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm"
                  value={matrixStartDate}
                  onChange={(e) => setMatrixStartDate(e.target.value)}
                />
              </div>
              <div className="space-y-2">
                <label className="text-xs font-bold text-zinc-500 uppercase">End Date</label>
                <input
                  type="date"
                  className="w-full px-4 py-2 bg-white border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm"
                  value={matrixEndDate}
                  onChange={(e) => setMatrixEndDate(e.target.value)}
                />
              </div>
              {(isAdmin || isBDEUser) && (
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by Agent</label>
                  <select
                    className="w-full px-4 py-2 bg-white border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm"
                    value={matrixAgentFilter}
                    onChange={(e) => setMatrixAgentFilter(e.target.value)}
                  >
                    <option value="all">All Agents</option>
                    {employees
                      .filter(emp => isAdmin || emp.bde === currentBDEName)
                      .map(emp => (
                        <option key={emp.id} value={emp.id}>{emp.name}</option>
                      ))}
                  </select>
                </div>
              )}
              {isAdmin && (
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by BDE</label>
                  <select
                    className="w-full px-4 py-2 bg-white border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm"
                    value={matrixBdeFilter}
                    onChange={(e) => setMatrixBdeFilter(e.target.value)}
                  >
                    <option value="all">All BDEs</option>
                    {bdes.map(bde => (
                      <option key={bde.id} value={bde.id}>{bde.name}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>
          </div>

          {/* Reports Table */}
          {matrixView === 'reports' && (
            <Card className="overflow-hidden border-zinc-200">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead>
                    <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase">
                      <th className="px-4 py-3 border-b border-r">Date</th>
                      <th className="px-4 py-3 border-b border-r">Employee</th>
                      <th className="px-4 py-3 border-b border-r text-center">Points</th>
                      <th className="px-4 py-3 border-b border-r text-center">Calls</th>
                      <th className="px-4 py-3 border-b border-r text-center">Talktime</th>
                      <th className="px-4 py-3 border-b border-r text-center">Floor Time</th>
                      {matrixStartDate === matrixEndDate && (
                        <>
                          <th className="px-4 py-3 border-b border-r">Top 4 Talktimes (Trip ID | Phone | Time)</th>
                          <th className="px-4 py-3 border-b border-r text-center">Approx In/Out</th>
                        </>
                      )}
                      <th className="px-4 py-3 border-b border-r text-center">Approx Break</th>
                      {matrixStartDate === matrixEndDate && (
                        <th className="px-4 py-3 border-b border-r">Log (Created/Updated)</th>
                      )}
                      <th className="px-4 py-3 border-b text-center">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {(() => {
                      const filtered = matrixReports.filter(r => {
                        // Access Control
                        if (!isAdmin) {
                          if (isBDEUser) {
                            // BDE can see their team members
                            const emp = employees.find(e => e.id === r.employeeId);
                            if (emp?.bde !== currentBDEName) return false;
                          } else {
                            // Agent can only see their own profile
                            if (r.employeeName !== currentEmployeeName) return false;
                          }
                        }

                        const matchesAgent = matrixAgentFilter === 'all' || r.employeeId === matrixAgentFilter;
                        const emp = employees.find(e => e.id === r.employeeId);
                        const matchesBde = matrixBdeFilter === 'all' || emp?.bde === matrixBdeFilter;
                        const reportDate = r.date;
                        const matchesDate = reportDate >= matrixStartDate && reportDate <= matrixEndDate;
                        return matchesAgent && matchesBde && matchesDate;
                      });

                      if (filtered.length === 0) {
                        return (
                          <tr>
                            <td colSpan={10} className="px-4 py-12 text-center text-zinc-400 italic">
                              No matrix reports found for the selected criteria.
                            </td>
                          </tr>
                        );
                      }

                      if (matrixStartDate === matrixEndDate) {
                        return filtered
                          .sort((a, b) => new Date(b.date).getTime() - new Date(a.date).getTime())
                          .map((report) => {
                            const reportDate = new Date(report.date);
                            const today = new Date();
                            today.setHours(0, 0, 0, 0);
                            const diffDays = Math.floor((today.getTime() - reportDate.getTime()) / (1000 * 60 * 60 * 24));
                            const canEdit = diffDays <= 3 || isAdmin;

                            return (
                              <tr key={report.id} className="hover:bg-zinc-50 transition-colors border-b last:border-0 group">
                                <td className="px-4 py-3 text-sm font-bold text-zinc-900 border-r">{report.date}</td>
                                <td className="px-4 py-3 text-sm text-zinc-600 border-r">{report.employeeName}</td>
                                <td className="px-4 py-3 text-sm text-center border-r font-black text-orange-600">{report.performancePoints}</td>
                                <td className="px-4 py-3 text-sm text-center border-r font-medium">{report.totalCalls}</td>
                                <td className="px-4 py-3 text-sm text-center border-r font-medium">{report.totalTalktime}</td>
                                <td className="px-4 py-3 text-sm text-center border-r font-bold text-emerald-600">
                                  {(() => {
                                    const calculateFloorTime = (login: string, logout: string, breakStr: string): number => {
                                      if (!login || !logout) return 0;
                                      const [loginH, loginM] = login.split(':').map(Number);
                                      const [logoutH, logoutM] = logout.split(':').map(Number);
                                      
                                      let diff = (logoutH * 60 + logoutM) - (loginH * 60 + loginM);
                                      if (diff < 0) diff += 24 * 60;
                                      
                                      const breakMins = talktimeToMinutes(breakStr);
                                      return Math.max(0, diff - breakMins);
                                    };
                                    return minutesToTalktime(calculateFloorTime(report.loginTime, report.logoutTime, report.breakTime));
                                  })()}
                                </td>
                                <td className="px-4 py-3 text-[10px] border-r">
                                  <div className="space-y-1">
                                    {report.topTalktimes.map((t, idx) => (
                                      <div key={idx} className="flex items-center gap-2 border-b border-zinc-100 last:border-0 pb-0.5">
                                        <span className="text-zinc-400 font-bold w-3">{idx + 1}.</span>
                                        <span className="text-zinc-600 font-medium w-16 truncate">{t.tripId || '-'}</span>
                                        <span className="text-zinc-400">|</span>
                                        <span className="text-zinc-600 w-20 truncate">{t.phoneNo || '-'}</span>
                                        <span className="text-zinc-400">|</span>
                                        <span className="font-bold text-orange-600">{t.talktime || '-'}</span>
                                      </div>
                                    ))}
                                  </div>
                                </td>
                                <td className="px-4 py-3 text-xs text-center border-r">
                                  <div className="flex flex-col">
                                    <span className="text-zinc-400">In: {report.loginTime}</span>
                                    <span className="text-zinc-400">Out: {report.logoutTime}</span>
                                  </div>
                                </td>
                                <td className="px-4 py-3 text-sm text-center border-r font-medium">{report.breakTime}</td>
                                <td className="px-4 py-3 text-[9px] border-r">
                                  <div className="space-y-1">
                                    <div className="flex flex-col">
                                      <span className="text-zinc-400 uppercase font-bold">Created:</span>
                                      <span className="text-zinc-600">{report.createdByName}</span>
                                      <span className="text-zinc-400">{report.createdAt?.toDate ? format(report.createdAt.toDate(), 'dd/MM/yy HH:mm') : '-'}</span>
                                    </div>
                                    {report.updatedAt && (
                                      <div className="flex flex-col border-t border-zinc-100 pt-1">
                                        <span className="text-zinc-400 uppercase font-bold">Updated:</span>
                                        <span className="text-zinc-600">{report.updatedByName}</span>
                                        <span className="text-zinc-400">{report.updatedAt?.toDate ? format(report.updatedAt.toDate(), 'dd/MM/yy HH:mm') : '-'}</span>
                                      </div>
                                    )}
                                  </div>
                                </td>
                                <td className="px-4 py-3 text-center">
                                  {canEdit ? (
                                    <button
                                      onClick={() => {
                                        setMatrixEditId(report.id!);
                                        setMatrixForm({
                                          ...report
                                        });
                                        setIsMatrixModalOpen(true);
                                      }}
                                      className="p-1.5 text-zinc-400 hover:text-orange-600 hover:bg-orange-50 rounded-lg transition-all"
                                    >
                                      <Edit2 className="w-4 h-4" />
                                    </button>
                                  ) : (
                                    <Lock className="w-4 h-4 text-zinc-300 mx-auto" />
                                  )}
                                </td>
                              </tr>
                            );
                          });
                      } else {
                        const aggregated: Record<string, any> = {};
                        filtered.forEach(r => {
                          const calculateFloorTime = (login: string, logout: string, breakStr: string): number => {
                            if (!login || !logout) return 0;
                            const [loginH, loginM] = login.split(':').map(Number);
                            const [logoutH, logoutM] = logout.split(':').map(Number);
                            
                            let diff = (logoutH * 60 + logoutM) - (loginH * 60 + loginM);
                            if (diff < 0) diff += 24 * 60;
                            
                            const breakMins = talktimeToMinutes(breakStr);
                            return Math.max(0, diff - breakMins);
                          };

                          if (!aggregated[r.employeeId]) {
                            aggregated[r.employeeId] = {
                              employeeId: r.employeeId,
                              employeeName: r.employeeName,
                              performancePoints: 0,
                              totalCalls: 0,
                              totalTalktimeMins: 0,
                              totalFloorTimeMins: 0,
                              breakTimeMins: 0,
                              count: 0
                            };
                          }
                          aggregated[r.employeeId].performancePoints += r.performancePoints;
                          aggregated[r.employeeId].totalCalls += r.totalCalls;
                          aggregated[r.employeeId].totalTalktimeMins += talktimeToMinutes(r.totalTalktime);
                          aggregated[r.employeeId].totalFloorTimeMins += calculateFloorTime(r.loginTime, r.logoutTime, r.breakTime);
                          aggregated[r.employeeId].breakTimeMins += talktimeToMinutes(r.breakTime);
                          aggregated[r.employeeId].count += 1;
                        });

                        return Object.values(aggregated)
                          .sort((a, b) => b.performancePoints - a.performancePoints)
                          .map(agg => (
                            <tr key={agg.employeeId} className="hover:bg-zinc-50 transition-colors border-b last:border-0 group">
                              <td className="px-4 py-3 text-sm font-bold text-zinc-900 border-r">{matrixStartDate} to {matrixEndDate}</td>
                              <td className="px-4 py-3 text-sm text-zinc-600 border-r">{agg.employeeName}</td>
                              <td className="px-4 py-3 text-sm text-center border-r font-black text-orange-600">{agg.performancePoints}</td>
                              <td className="px-4 py-3 text-sm text-center border-r font-medium">{agg.totalCalls}</td>
                              <td className="px-4 py-3 text-sm text-center border-r font-medium">{minutesToTalktime(agg.totalTalktimeMins)}</td>
                              <td className="px-4 py-3 text-sm text-center border-r font-bold text-emerald-600">{minutesToTalktime(agg.totalFloorTimeMins)}</td>
                              <td className="px-4 py-3 text-sm text-center border-r font-medium">{minutesToTalktime(agg.breakTimeMins)}</td>
                              <td className="px-4 py-3 text-center">
                                <span className="text-[10px] text-zinc-400 uppercase font-bold">Aggregated ({agg.count} days)</span>
                              </td>
                            </tr>
                          ));
                      }
                    })()}
                  </tbody>
                </table>
              </div>
            </Card>
          )}

          {/* Trip ID Aggregated Table */}
          {matrixView === 'trips' && (
            <div className="space-y-4">
              <div className="flex items-center gap-2">
                <PhoneCall className="w-5 h-5 text-orange-600" />
                <h3 className="text-xl font-bold tracking-tight">Trip ID Performance Summary</h3>
              </div>
              <Card className="overflow-hidden border-zinc-200">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase">
                        <th className="px-4 py-3 border-b border-r">Trip ID</th>
                        <th className="px-4 py-3 border-b border-r">Phone No</th>
                        <th className="px-4 py-3 border-b border-r">Agent</th>
                        <th className="px-4 py-3 border-b border-r">BDE</th>
                        <th className="px-4 py-3 border-b border-r text-center">Total Talktime</th>
                        <th className="px-4 py-3 border-b border-r text-center">No of Calls</th>
                        <th className="px-4 py-3 border-b">Dates</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const filtered = matrixReports.filter(r => {
                          if (!isAdmin) {
                            if (isBDEUser) {
                              const emp = employees.find(e => e.id === r.employeeId);
                              if (emp?.bde !== currentBDEName) return false;
                            } else {
                              if (r.employeeName !== currentEmployeeName) return false;
                            }
                          }
                          const matchesAgent = matrixAgentFilter === 'all' || r.employeeId === matrixAgentFilter;
                          const emp = employees.find(e => e.id === r.employeeId);
                          const matchesBde = matrixBdeFilter === 'all' || emp?.bde === matrixBdeFilter;
                          const reportDate = r.date;
                          const matchesDate = reportDate >= matrixStartDate && reportDate <= matrixEndDate;
                          return matchesAgent && matchesBde && matchesDate;
                        });

                        const tripGroups: Record<string, {
                          tripId: string;
                          phoneNo: string;
                          agentNames: Set<string>;
                          bdeNames: Set<string>;
                          totalTalktimeMins: number;
                          noOfCalls: number;
                          dates: Set<string>;
                        }> = {};

                        filtered.forEach(report => {
                          const emp = employees.find(e => e.id === report.employeeId);
                          report.topTalktimes.forEach(t => {
                            if (!t.tripId) return;
                            const tid = t.tripId.trim();
                            if (!tripGroups[tid]) {
                              tripGroups[tid] = {
                                tripId: tid,
                                phoneNo: t.phoneNo || '-',
                                agentNames: new Set(),
                                bdeNames: new Set(),
                                totalTalktimeMins: 0,
                                noOfCalls: 0,
                                dates: new Set()
                              };
                            }
                            tripGroups[tid].totalTalktimeMins += talktimeToMinutes(t.talktime);
                            tripGroups[tid].noOfCalls += 1;
                            tripGroups[tid].dates.add(report.date);
                            tripGroups[tid].agentNames.add(report.employeeName);
                            if (emp?.bde) tripGroups[tid].bdeNames.add(emp.bde);
                          });
                        });

                        const sortedTrips = Object.values(tripGroups).sort((a, b) => b.totalTalktimeMins - a.totalTalktimeMins);

                        if (sortedTrips.length === 0) {
                          return (
                            <tr>
                              <td colSpan={7} className="px-4 py-12 text-center text-zinc-400 italic">
                                No trip data found for the selected criteria.
                              </td>
                            </tr>
                          );
                        }

                        return sortedTrips.map(trip => (
                          <tr key={trip.tripId} className="hover:bg-zinc-50 transition-colors border-b last:border-0 group">
                            <td className="px-4 py-3 text-sm font-bold text-zinc-900 border-r">{trip.tripId}</td>
                            <td className="px-4 py-3 text-sm text-zinc-600 border-r">{trip.phoneNo}</td>
                            <td className="px-4 py-3 text-sm text-zinc-600 border-r">{Array.from(trip.agentNames).join(', ')}</td>
                            <td className="px-4 py-3 text-sm text-zinc-600 border-r">{Array.from(trip.bdeNames).join(', ')}</td>
                            <td className="px-4 py-3 text-sm text-center border-r font-black text-orange-600">{minutesToTalktime(trip.totalTalktimeMins)}</td>
                            <td className="px-4 py-3 text-sm text-center border-r font-medium">{trip.noOfCalls}</td>
                            <td className="px-4 py-3 text-[10px] text-zinc-500 italic">
                              {Array.from(trip.dates).sort().join(', ')}
                            </td>
                          </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                </div>
              </Card>
            </div>
          )}

          {/* Matrix Report Modal */}
          <Modal
            isOpen={isMatrixModalOpen}
            onClose={() => {
              setIsMatrixModalOpen(false);
              setMatrixEditId(null);
            }}
            title={matrixEditId ? "Edit Matrix Report" : "Submit Daily Matrix Report"}
          >
            <form onSubmit={handleSaveMatrixReport} className="space-y-6">
              {(isAdmin || isBDEUser) && !matrixEditId && (
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Select Agent</label>
                  <select
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.employeeId || ''}
                    onChange={(e) => {
                      const emp = employees.find(emp => emp.id === e.target.value);
                      setMatrixForm({ 
                        ...matrixForm, 
                        employeeId: e.target.value,
                        employeeName: emp?.name || ''
                      });
                    }}
                  >
                    <option value="">Select Agent</option>
                    {employees
                      .filter(e => e.status === 'Active')
                      .filter(e => isAdmin || e.bde === currentBDEName)
                      .map(emp => (
                        <option key={emp.id} value={emp.id}>{emp.name}</option>
                      ))}
                  </select>
                </div>
              )}

              {!isAdmin && !isBDEUser && !currentEmployee && (
                <div className="p-3 bg-amber-50 border border-amber-100 rounded-xl flex items-center gap-2 text-amber-600 text-xs">
                  <AlertCircle className="w-4 h-4" />
                  Warning: Your profile is not linked to an employee record. Please contact Admin.
                </div>
              )}

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Date</label>
                  <input
                    type="date"
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.date || ''}
                    onChange={(e) => setMatrixForm({ ...matrixForm, date: e.target.value })}
                    min={!isAdmin ? format(subDays(new Date(), 2), 'yyyy-MM-dd') : undefined}
                    max={format(new Date(), 'yyyy-MM-dd')}
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Performance Points</label>
                  <input
                    type="number"
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.performancePoints || ''}
                    onChange={(e) => setMatrixForm({ ...matrixForm, performancePoints: Number(e.target.value) })}
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Total Calls</label>
                  <input
                    type="number"
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.totalCalls || ''}
                    onChange={(e) => setMatrixForm({ ...matrixForm, totalCalls: Number(e.target.value) })}
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Total Talktime</label>
                  <select
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.totalTalktime || ''}
                    onChange={(e) => setMatrixForm({ ...matrixForm, totalTalktime: e.target.value })}
                  >
                    <option value="">Select Talktime</option>
                    {generateTalktimeOptions().map(opt => (
                      <option key={opt.value} value={opt.value}>{opt.value}</option>
                    ))}
                  </select>
                </div>
              </div>

              <div className="space-y-3">
                <label className="text-xs font-bold text-zinc-500 uppercase block">Top 4 Talktimes</label>
                <div className="grid grid-cols-1 gap-3">
                  {matrixForm.topTalktimes?.map((t, idx) => (
                    <div key={idx} className="flex gap-2 items-center">
                      <span className="text-xs font-bold text-zinc-400 w-4">{idx + 1}.</span>
                      <input
                        type="text"
                        required
                        placeholder="Trip ID"
                        className="flex-1 px-3 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-xs"
                        value={t.tripId}
                        onChange={(e) => {
                          const newTops = [...(matrixForm.topTalktimes || [])];
                          newTops[idx] = { ...newTops[idx], tripId: e.target.value };
                          setMatrixForm({ ...matrixForm, topTalktimes: newTops });
                        }}
                      />
                      <input
                        type="text"
                        required
                        placeholder="Phone No"
                        className="flex-1 px-3 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-xs"
                        value={t.phoneNo}
                        onChange={(e) => {
                          const newTops = [...(matrixForm.topTalktimes || [])];
                          newTops[idx] = { ...newTops[idx], phoneNo: e.target.value };
                          setMatrixForm({ ...matrixForm, topTalktimes: newTops });
                        }}
                      />
                      <select
                        required
                        className="w-24 px-3 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-xs"
                        value={t.talktime}
                        onChange={(e) => {
                          const newTops = [...(matrixForm.topTalktimes || [])];
                          newTops[idx] = { ...newTops[idx], talktime: e.target.value };
                          setMatrixForm({ ...matrixForm, topTalktimes: newTops });
                        }}
                      >
                        <option value="">Time</option>
                        {generateTalktimeOptions().map(opt => (
                          <option key={opt.value} value={opt.value}>{opt.value}</option>
                        ))}
                      </select>
                    </div>
                  ))}
                </div>
              </div>

              <div className="grid grid-cols-3 gap-4">
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Approx Login Time</label>
                  <input
                    type="time"
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.loginTime || ''}
                    onChange={(e) => setMatrixForm({ ...matrixForm, loginTime: e.target.value })}
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Approx Logout Time</label>
                  <input
                    type="time"
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.logoutTime || ''}
                    onChange={(e) => setMatrixForm({ ...matrixForm, logoutTime: e.target.value })}
                  />
                </div>
                <div className="space-y-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Approx Break Time</label>
                  <select
                    required
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={matrixForm.breakTime || ''}
                    onChange={(e) => setMatrixForm({ ...matrixForm, breakTime: e.target.value })}
                  >
                    <option value="">Select Break</option>
                    {generateTalktimeOptions().map(opt => (
                      <option key={opt.value} value={opt.value}>{opt.value}</option>
                    ))}
                  </select>
                </div>
              </div>

              {error && (
                <div className="p-3 bg-red-50 border border-red-100 rounded-xl flex items-center gap-2 text-red-600 text-xs">
                  <AlertCircle className="w-4 h-4" />
                  {error}
                </div>
              )}

              {success && (
                <div className="p-3 bg-green-50 border border-green-100 rounded-xl flex items-center gap-2 text-green-600 text-xs">
                  <CheckCircle2 className="w-4 h-4" />
                  {success}
                </div>
              )}

              <div className="flex gap-3 pt-4">
                <Button type="submit" className="flex-1" isLoading={isSavingMatrix}>
                  {matrixEditId ? "Update Report" : "Submit Report"}
                </Button>
                <Button type="button" variant="ghost" onClick={() => {
                  setIsMatrixModalOpen(false);
                  setMatrixEditId(null);
                }}>
                  Cancel
                </Button>
              </div>
            </form>
          </Modal>
        </div>
      )}

      {/* Password Manager Page */}
      {activeTab === 'passwordManager' && (
        <div className="space-y-6">
          <div className="flex items-center justify-between">
            <div className="space-y-1">
              <h2 className="text-2xl font-black tracking-tight text-zinc-900">Password Manager</h2>
              <p className="text-zinc-500 font-medium">Manage password protection for application pages.</p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {/* Global Delete Password */}
            {(() => {
              const config = pagePasswords['globalDelete'] || { password: '', isOpen: false };
              return (
                <Card className="p-6 space-y-4 border-red-100 bg-red-50/30">
                  <div className="flex items-center justify-between">
                    <div className="flex items-center gap-2">
                      <div className="p-1.5 bg-red-100 rounded-lg">
                        <Trash2 className="w-4 h-4 text-red-600" />
                      </div>
                      <h3 className="font-bold text-zinc-900">Delete Protection</h3>
                    </div>
                    <div className={cn(
                      "px-2 py-1 rounded-full text-[10px] font-bold uppercase",
                      config.isOpen ? "bg-emerald-100 text-emerald-700" : "bg-red-100 text-red-700"
                    )}>
                      {config.isOpen ? 'No Password' : 'Password Required'}
                    </div>
                  </div>

                  <p className="text-[10px] text-zinc-500 font-medium leading-relaxed">
                    When enabled, a password will be required for all delete operations across the application.
                  </p>

                  <div className="flex items-center gap-4">
                    <button
                      onClick={() => handleUpdatePagePassword('globalDelete', config.password, true)}
                      className={cn(
                        "flex-1 py-2 text-xs font-bold rounded-lg transition-colors",
                        config.isOpen ? "bg-red-600 text-white" : "bg-white border border-zinc-200 text-zinc-600 hover:bg-zinc-50"
                      )}
                    >
                      Disable
                    </button>
                    <button
                      onClick={() => handleUpdatePagePassword('globalDelete', config.password, false)}
                      className={cn(
                        "flex-1 py-2 text-xs font-bold rounded-lg transition-colors",
                        !config.isOpen ? "bg-red-600 text-white" : "bg-white border border-zinc-200 text-zinc-600 hover:bg-zinc-50"
                      )}
                    >
                      Enable
                    </button>
                  </div>

                  {!config.isOpen && (
                    <div className="space-y-2 pt-2 border-t border-red-100">
                      <label className="text-[10px] font-bold text-red-600 uppercase">Delete Password</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          className="flex-1 px-3 py-1.5 bg-white border border-red-200 rounded-lg text-sm focus:ring-2 focus:ring-red-500 outline-none"
                          placeholder="Set delete password"
                          defaultValue={config.password}
                          onBlur={(e) => {
                            if (e.target.value !== config.password) {
                              handleUpdatePagePassword('globalDelete', e.target.value, false);
                            }
                          }}
                        />
                      </div>
                    </div>
                  )}
                </Card>
              );
            })()}

            {allPermissions.map((page) => {
              const config = pagePasswords[page.id] || { password: '', isOpen: true };
              return (
                <Card key={page.id} className="p-6 space-y-4">
                  <div className="flex items-center justify-between">
                    <h3 className="font-bold text-zinc-900">{page.label}</h3>
                    <div className={cn(
                      "px-2 py-1 rounded-full text-[10px] font-bold uppercase",
                      config.isOpen ? "bg-emerald-100 text-emerald-700" : "bg-orange-100 text-orange-700"
                    )}>
                      {config.isOpen ? 'Open' : 'Password Protected'}
                    </div>
                  </div>

                  <div className="flex items-center gap-4">
                    <button
                      onClick={() => handleUpdatePagePassword(page.id, config.password, true)}
                      className={cn(
                        "flex-1 py-2 text-xs font-bold rounded-lg transition-colors",
                        config.isOpen ? "bg-orange-600 text-white" : "bg-zinc-100 text-zinc-600 hover:bg-zinc-200"
                      )}
                    >
                      Open
                    </button>
                    <button
                      onClick={() => handleUpdatePagePassword(page.id, config.password, false)}
                      className={cn(
                        "flex-1 py-2 text-xs font-bold rounded-lg transition-colors",
                        !config.isOpen ? "bg-orange-600 text-white" : "bg-zinc-100 text-zinc-600 hover:bg-zinc-200"
                      )}
                    >
                      Password
                    </button>
                  </div>

                  {!config.isOpen && (
                    <div className="space-y-2 pt-2 border-t border-zinc-100">
                      <label className="text-[10px] font-bold text-zinc-500 uppercase">Page Password</label>
                      <div className="flex gap-2">
                        <input
                          type="text"
                          className="flex-1 px-3 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-sm"
                          placeholder="Enter password"
                          defaultValue={config.password}
                          onBlur={(e) => {
                            if (e.target.value !== config.password) {
                              handleUpdatePagePassword(page.id, e.target.value, false);
                            }
                          }}
                        />
                      </div>
                    </div>
                  )}
                </Card>
              );
            })}
          </div>
        </div>
      )}

      {activeTab === 'dashboard' && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-8">
            {/* --- Left Column: Stats & Award --- */}
            <div className="lg:col-span-8 space-y-8">
              {/* Welcome Section */}
              <section className="space-y-4">
                <h2 className="text-3xl font-bold tracking-tight">Hello, {profile?.displayName?.split(' ')[0]}! 👋</h2>
                <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                  <Card className="p-6 bg-orange-600 text-white border-none">
                    <p className="text-orange-100 text-xs font-bold uppercase tracking-wider">Your Balance</p>
                    <div className="flex items-baseline gap-2 mt-1">
                      <span className="text-4xl font-black">{profile?.pointsBalance || 0}</span>
                      <span className="text-orange-200 font-medium">pts</span>
                    </div>
                    <TrendingUp className="absolute top-4 right-4 w-12 h-12 text-white/10" />
                  </Card>
                  <Card className="p-6">
                    <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Awards Given</p>
                    <div className="flex items-baseline gap-2 mt-1">
                      <span className="text-4xl font-black">
                        {transactions.filter(t => t.fromUid === profile?.uid).length}
                      </span>
                    </div>
                  </Card>
                  <Card className="p-6">
                    <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Awards Received</p>
                    <div className="flex items-baseline gap-2 mt-1">
                      <span className="text-4xl font-black">
                        {transactions.filter(t => t.toUid === profile?.uid).length}
                      </span>
                    </div>
                  </Card>
                </div>
              </section>

              {/* Award Points Form */}
              <section id="award-points" className="space-y-4">
                <div className="flex items-center gap-2">
                  <Send className="w-5 h-5 text-orange-600" />
                  <h3 className="text-xl font-bold">Award Bravo Points</h3>
                </div>
                <Card className="p-6">
                  <form onSubmit={handleAwardPoints} className="space-y-6">
                    <div className="space-y-2 relative">
                      <label className="text-sm font-bold text-zinc-700">Recipient</label>
                      <div className="relative">
                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-400" />
                        <input 
                          type="text" 
                          placeholder="Search by name or email..."
                          className="w-full pl-10 pr-4 py-3 bg-zinc-50 border border-zinc-200 rounded-xl focus:ring-2 focus:ring-orange-500 focus:border-transparent outline-none transition-all"
                          value={searchQuery}
                          onChange={(e) => setSearchQuery(e.target.value)}
                        />
                      </div>
                      
                      <AnimatePresence>
                        {filteredUsers.length > 0 && (
                          <motion.div 
                            initial={{ opacity: 0, y: -10 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, y: -10 }}
                            className="absolute z-10 w-full mt-1 bg-white border border-zinc-200 rounded-xl shadow-xl overflow-hidden"
                          >
                            {filteredUsers.map(u => (
                              <button
                                key={u.uid}
                                type="button"
                                className="w-full flex items-center gap-3 p-3 hover:bg-orange-50 transition-colors text-left"
                                onClick={() => {
                                  setSelectedUser(u);
                                  setSearchQuery('');
                                }}
                              >
                                <img src={u.photoURL} className="w-8 h-8 rounded-full" alt="" referrerPolicy="no-referrer" />
                                <div>
                                  <p className="text-sm font-bold">{u.displayName}</p>
                                  <p className="text-xs text-zinc-500">{u.email}</p>
                                </div>
                              </button>
                            ))}
                          </motion.div>
                        )}
                      </AnimatePresence>

                      {selectedUser && (
                        <div className="flex items-center justify-between p-3 bg-orange-50 border border-orange-100 rounded-xl">
                          <div className="flex items-center gap-3">
                            <img src={selectedUser.photoURL} className="w-10 h-10 rounded-full" alt="" referrerPolicy="no-referrer" />
                            <div>
                              <p className="text-sm font-bold">{selectedUser.displayName}</p>
                              <p className="text-xs text-orange-600">Recipient Selected</p>
                            </div>
                          </div>
                          <Button variant="ghost" className="p-2" onClick={() => setSelectedUser(null)}>Remove</Button>
                        </div>
                      )}
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                      <div className="space-y-2">
                        <label className="text-sm font-bold text-zinc-700">Points to Award</label>
                        <div className="flex gap-2">
                          {[10, 20, 50, 100].map(p => (
                            <button
                              key={p}
                              type="button"
                              className={cn(
                                "flex-1 py-2 rounded-lg font-bold transition-all border",
                                pointsToAward === p 
                                  ? "bg-orange-600 text-white border-orange-600" 
                                  : "bg-white text-zinc-600 border-zinc-200 hover:border-orange-300"
                              )}
                              onClick={() => setPointsToAward(p)}
                            >
                              {p}
                            </button>
                          ))}
                        </div>
                      </div>
                      <div className="space-y-2">
                        <label className="text-sm font-bold text-zinc-700">Custom Amount</label>
                        <input 
                          type="number" 
                          className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-xl focus:ring-2 focus:ring-orange-500 outline-none"
                          value={pointsToAward}
                          onChange={(e) => setPointsToAward(parseInt(e.target.value) || 0)}
                        />
                      </div>
                    </div>

                    <div className="space-y-2">
                      <label className="text-sm font-bold text-zinc-700">Reason for Bravo</label>
                      <textarea 
                        placeholder="Tell them why they're awesome..."
                        className="w-full px-4 py-3 bg-zinc-50 border border-zinc-200 rounded-xl focus:ring-2 focus:ring-orange-500 outline-none min-h-[100px]"
                        value={reason}
                        onChange={(e) => setReason(e.target.value)}
                      />
                    </div>

                    {error && (
                      <div className="p-3 bg-red-50 border border-red-100 rounded-xl flex items-center gap-2 text-red-600 text-sm">
                        <AlertCircle className="w-4 h-4" />
                        {error}
                      </div>
                    )}

                    {success && (
                      <div className="p-3 bg-green-50 border border-green-100 rounded-xl flex items-center gap-2 text-green-600 text-sm">
                        <CheckCircle2 className="w-4 h-4" />
                        {success}
                      </div>
                    )}

                    <Button 
                      type="submit" 
                      className="w-full py-4 text-lg" 
                      isLoading={isAwarding}
                      disabled={!selectedUser || !reason.trim() || pointsToAward <= 0}
                    >
                      Give Bravo Points
                    </Button>
                  </form>
                </Card>
              </section>

              {/* Recent Activity */}
              <section className="space-y-4">
                <div className="flex items-center gap-2">
                  <History className="w-5 h-5 text-orange-600" />
                  <h3 className="text-xl font-bold">Recent Activity</h3>
                </div>
                <div className="space-y-3">
                  {filteredTransactions.length === 0 ? (
                    <Card className="p-12 text-center text-zinc-400">
                      <p>No transactions yet. Be the first to give a Bravo!</p>
                    </Card>
                  ) : (
                    filteredTransactions.map((t, idx) => (
                      <motion.div
                        initial={{ opacity: 0, x: -20 }}
                        animate={{ opacity: 1, x: 0 }}
                        transition={{ delay: idx * 0.05 }}
                        key={t.id}
                      >
                        <Card className="p-4 flex items-center gap-4 hover:shadow-md transition-shadow">
                          <div className="w-12 h-12 bg-zinc-100 rounded-full flex items-center justify-center shrink-0">
                            {t.fromUid === profile?.uid ? (
                              <Send className="w-6 h-6 text-zinc-400" />
                            ) : (
                              <Award className="w-6 h-6 text-orange-600" />
                            )}
                          </div>
                          <div className="flex-1 min-w-0">
                            <p className="text-sm font-bold truncate">
                              <span className="text-orange-600">{t.fromName}</span> awarded <span className="text-orange-600">{t.points} pts</span> to <span className="text-orange-600">{t.toName}</span>
                            </p>
                            <p className="text-xs text-zinc-500 italic truncate">"{t.reason}"</p>
                          </div>
                          <div className="text-right shrink-0">
                            <p className="text-[10px] text-zinc-400 font-bold uppercase tracking-tighter">
                              {t.createdAt?.toDate ? formatDistanceToNow(t.createdAt.toDate(), { addSuffix: true }) : 'Just now'}
                            </p>
                          </div>
                        </Card>
                      </motion.div>
                    ))
                  )}
                </div>
              </section>
            </div>

            {/* --- Right Column: Leaderboard & Profile --- */}
            <div className="lg:col-span-4 space-y-8">
              {/* Leaderboard */}
              <section className="space-y-4">
                <div className="flex items-center gap-2">
                  <Trophy className="w-5 h-5 text-orange-600" />
                  <h3 className="text-xl font-bold">Top Contributors</h3>
                </div>
                <Card className="divide-y divide-zinc-100">
                  {leaderboard.map((u, idx) => (
                    <div key={u.uid} className="p-4 flex items-center gap-3">
                      <div className="w-6 text-center font-black text-zinc-300">
                        {idx + 1}
                      </div>
                      <img src={u.photoURL} className="w-10 h-10 rounded-full border border-zinc-100" alt="" referrerPolicy="no-referrer" />
                      <div className="flex-1 min-w-0">
                        <p className="text-sm font-bold truncate">{u.displayName}</p>
                        <p className="text-xs text-zinc-500">{u.pointsBalance} pts</p>
                      </div>
                      {idx === 0 && <Star className="w-5 h-5 text-yellow-400 fill-yellow-400" />}
                    </div>
                  ))}
                </Card>
              </section>

              {/* Quick Tips */}
              <Card className="p-6 bg-zinc-900 text-white space-y-4">
                <h4 className="font-bold">How it works</h4>
                <ul className="space-y-3 text-sm text-zinc-400">
                  <li className="flex gap-2">
                    <ChevronRight className="w-4 h-4 text-orange-500 shrink-0" />
                    <span>Everyone starts with 100 points.</span>
                  </li>
                  <li className="flex gap-2">
                    <ChevronRight className="w-4 h-4 text-orange-500 shrink-0" />
                    <span>Award points to colleagues for great work.</span>
                  </li>
                  <li className="flex gap-2">
                    <ChevronRight className="w-4 h-4 text-orange-500 shrink-0" />
                    <span>Points you give are deducted from your balance.</span>
                  </li>
                  <li className="flex gap-2">
                    <ChevronRight className="w-4 h-4 text-orange-500 shrink-0" />
                    <span>Points you receive are added to your balance.</span>
                  </li>
                </ul>
              </Card>

              {/* User Profile Summary */}
              <Card className="p-6 space-y-4">
                <div className="flex items-center gap-4">
                  <img src={profile?.photoURL} className="w-16 h-16 rounded-2xl" alt="" referrerPolicy="no-referrer" />
                  <div>
                    <h4 className="font-bold text-lg leading-tight">{profile?.displayName}</h4>
                    <p className="text-sm text-zinc-500">{profile?.email}</p>
                  </div>
                </div>
                <div className="pt-4 border-t border-zinc-100 flex justify-between items-center">
                  <span className="text-sm text-zinc-500">Member since</span>
                  <span className="text-sm font-bold">March 2026</span>
                </div>
              </Card>
            </div>
          </div>
        )}

        {activeTab === 'employees' && hasPermission('employees') && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Users className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Employee Master</h2>
              </div>
              <div className="flex items-center gap-2">
                {getPermissionLevel('employees') === 'Complete' && (
                  <Button 
                    onClick={downloadEmployeeMaster}
                    variant="outline"
                    className="border-zinc-200"
                  >
                    <Download className="w-4 h-4" />
                    Download Excel
                  </Button>
                )}
                {getPermissionLevel('employees') === 'Complete' && (
                  <Button 
                    onClick={() => {
                      setIsEditingEmployee(false);
                      setEmployeeForm({
                        status: 'Active',
                        position: 'Sales Agent',
                        joiningWeek: 'Week 1',
                        bde: 'None',
                        email: '',
                        isAssocBDE: false,
                        week1Target: 25,
                        week2Target: 25,
                        week3Target: 50,
                        week4Target: 50,
                        week5Target: 50,
                        week6Target: 50,
                        salary: 0,
                        completedTrainings: [],
                      });
                      setIsEmployeeModalOpen(true);
                    }} 
                    className="bg-orange-600"
                  >
                    <Plus className="w-4 h-4" />
                    Add Employee
                  </Button>
                )}
              </div>
            </div>

            <Modal 
              isOpen={isEmployeeModalOpen} 
              onClose={() => setIsEmployeeModalOpen(false)}
              title={isEditingEmployee ? 'Edit Employee' : 'Create New Employee'}
            >
              <form onSubmit={handleCreateEmployee} className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Employee Code</label>
                    <input 
                      type="text" 
                      disabled 
                      value={employeeForm.employeeCode || "Auto Generated"} 
                      className="w-full px-4 py-2 bg-zinc-100 border border-zinc-200 rounded-lg text-zinc-400 italic"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Name</label>
                    <input 
                      type="text" 
                      required
                      placeholder="e.g. Vaishali"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.name || ''}
                      onChange={(e) => setEmployeeForm({...employeeForm, name: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Email</label>
                    <input 
                      type="email" 
                      required
                      placeholder="e.g. vaishali@example.com"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.email || ''}
                      onChange={(e) => setEmployeeForm({...employeeForm, email: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Target</label>
                    <input 
                      type="number" 
                      required
                      placeholder="e.g. 500000"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.target || ''}
                      onChange={(e) => setEmployeeForm({...employeeForm, target: parseFloat(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Position</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.position}
                      onChange={(e) => setEmployeeForm({...employeeForm, position: e.target.value})}
                    >
                      <option>Sales Agent</option>
                      <option>HR</option>
                      <option>Associate BDE</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Status</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.status}
                      onChange={(e) => setEmployeeForm({...employeeForm, status: e.target.value as any})}
                    >
                      <option>Active</option>
                      <option>Deactive</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Salary</label>
                    <input 
                      type="number" 
                      required
                      placeholder="e.g. 25000"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.salary || ''}
                      onChange={(e) => setEmployeeForm({...employeeForm, salary: parseFloat(e.target.value)})}
                    />
                  </div>
                  <div className="flex items-center gap-2 pt-6">
                    <input 
                      type="checkbox" 
                      id="isAssocBDE"
                      className="w-4 h-4 text-orange-600 border-zinc-300 rounded focus:ring-orange-500"
                      checked={employeeForm.isAssocBDE || false}
                      onChange={(e) => setEmployeeForm({...employeeForm, isAssocBDE: e.target.checked})}
                    />
                    <label htmlFor="isAssocBDE" className="text-xs font-bold text-zinc-500 uppercase cursor-pointer">Act as Assoc BDE</label>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Joining Date</label>
                    <input 
                      type="date" 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.joiningDate || ''}
                      onChange={(e) => setEmployeeForm({...employeeForm, joiningDate: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Joining Week</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.joiningWeek}
                      onChange={(e) => setEmployeeForm({...employeeForm, joiningWeek: e.target.value})}
                    >
                      <option value="">Select Week</option>
                      {allSortedWeeks.map(wk => (
                        <option key={wk.id} value={wk.weekName}>{wk.weekName}</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">BDE</label>
                    <SearchableSelect
                      options={[
                        { value: 'None', label: 'None' },
                        ...bdes.sort((a, b) => a.name.localeCompare(b.name)).map(bde => ({ value: bde.name, label: bde.name, key: bde.id }))
                      ]}
                      value={employeeForm.bde}
                      onChange={(val) => setEmployeeForm({...employeeForm, bde: val})}
                      placeholder="None"
                    />
                  </div>
                </div>

                <div className="pt-4 border-t border-zinc-100 space-y-4">
                  <p className="text-xs font-bold text-zinc-400 uppercase">Weekly Target Percentages</p>
                  <div className="grid grid-cols-3 md:grid-cols-6 gap-3">
                    {[1, 2, 3, 4, 5, 6].map(num => (
                      <div key={num} className="space-y-1">
                        <label className="text-[10px] font-bold text-zinc-500 uppercase">W{num} %</label>
                        <input 
                          type="number" 
                          className="w-full px-2 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-sm"
                          value={(employeeForm as any)[`week${num}Target`] || 0}
                          onChange={(e) => setEmployeeForm({...employeeForm, [`week${num}Target`]: parseInt(e.target.value)})}
                        />
                      </div>
                    ))}
                  </div>
                </div>

                <div className="pt-4 border-t border-zinc-100 space-y-4">
                  <p className="text-xs font-bold text-zinc-400 uppercase">Modules</p>
                  <div className="grid grid-cols-1 md:grid-cols-3 gap-3 p-2 border border-zinc-200 rounded-lg bg-zinc-50">
                    {['Sales', 'Operations', 'Lead'].map(mod => (
                      <label key={mod} className="flex items-center gap-2 cursor-pointer">
                        <input 
                          type="checkbox"
                          className="w-4 h-4 text-orange-600 border-zinc-300 rounded focus:ring-orange-500"
                          checked={(employeeForm.completedTrainings || []).includes(mod)}
                          onChange={(e) => {
                            const current = employeeForm.completedTrainings || [];
                            if (e.target.checked) {
                              setEmployeeForm({...employeeForm, completedTrainings: [...current, mod]});
                            } else {
                              setEmployeeForm({...employeeForm, completedTrainings: current.filter(m => m !== mod)});
                            }
                          }}
                        />
                        <span className="text-sm text-zinc-700">{mod}</span>
                      </label>
                    ))}
                  </div>
                </div>

                {error && <p className="text-xs text-red-500">{error}</p>}
                {success && <p className="text-xs text-green-500">{success}</p>}

                <div className="flex gap-3 pt-4">
                  <Button type="submit" className="flex-1" isLoading={isSavingEmployee}>
                    {isEditingEmployee ? 'Update Record' : 'Save Employee Record'}
                  </Button>
                  <Button 
                    type="button" 
                    variant="ghost" 
                    onClick={() => setIsEmployeeModalOpen(false)}
                  >
                    Cancel
                  </Button>
                </div>
              </form>
            </Modal>

            <div className="space-y-4">
              <div className="flex items-center justify-between">
                <h3 className="text-lg font-bold">Employee Directory</h3>
                <div className="flex items-center gap-2">
                  <select 
                    className="px-4 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none"
                    value={employeeStatusFilter}
                    onChange={(e) => setEmployeeStatusFilter(e.target.value as any)}
                  >
                    <option value="Active">Active</option>
                    <option value="Deactive">Deactive</option>
                    <option value="All">All Status</option>
                  </select>
                  <select 
                    className="px-4 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none"
                    value={employeeAssocBdeFilter}
                    onChange={(e) => setEmployeeAssocBdeFilter(e.target.value as any)}
                  >
                    <option value="All">All Assoc BDE</option>
                    <option value="Yes">Assoc BDE Only</option>
                    <option value="No">Non-Assoc BDE</option>
                  </select>
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-400" />
                    <input 
                      type="text" 
                      placeholder="Filter employees..."
                      className="pl-9 pr-4 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none"
                      value={employeeSearchQuery}
                      onChange={(e) => setEmployeeSearchQuery(e.target.value)}
                    />
                  </div>
                </div>
              </div>

              <Card className="overflow-x-auto">
                <table className="w-full text-left border-collapse border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50">
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Code</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Name</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Email</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Position</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Target</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Joined</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Status</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredEmployees.length === 0 ? (
                      <tr>
                        <td colSpan={7} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No employee records found.
                        </td>
                      </tr>
                    ) : (
                      filteredEmployees.map(emp => (
                        <tr key={emp.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-mono text-orange-600 border border-zinc-200">{emp.employeeCode}</td>
                          <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{emp.name}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{emp.email}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{emp.position}</td>
                          <td className="px-4 py-3 text-sm font-bold border border-zinc-200">₹{emp.target.toLocaleString()}</td>
                          <td className="px-4 py-3 text-sm text-zinc-500 border border-zinc-200">{emp.joiningDate}</td>
                          <td className="px-4 py-3 border border-zinc-200">
                            <Badge className={emp.status === 'Active' ? 'bg-green-100 text-green-700' : 'bg-zinc-100 text-zinc-500'}>
                              {emp.status}
                            </Badge>
                          </td>
                          <td className="px-4 py-3 border border-zinc-200">
                            <div className="flex items-center gap-2">
                              {getPermissionLevel('employees') === 'Complete' && (
                                <button 
                                  onClick={() => handleEditEmployee(emp)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {getPermissionLevel('employees') === 'Complete' && (
                                <button 
                                  onClick={() => emp.id && handleDeleteEmployee(emp.id, emp.name)}
                                  className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </Card>
            </div>
          </div>
        )}

        {activeTab === 'weeks' && hasPermission('weeks') && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <CalendarIcon className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Weeks Master</h2>
              </div>
              {getPermissionLevel('weeks') === 'Complete' && (
                <Button 
                  onClick={() => {
                    setIsEditingWeek(false);
                    setWeekForm({
                      weekName: '',
                      month: 'January',
                      startDate: '',
                      endDate: ''
                    });
                    setIsWeekModalOpen(true);
                  }} 
                  className="bg-orange-600"
                >
                  <Plus className="w-4 h-4" />
                  Add Week
                </Button>
              )}
            </div>

            <Modal 
              isOpen={isWeekModalOpen} 
              onClose={() => setIsWeekModalOpen(false)}
              title={isEditingWeek ? 'Edit Week' : 'Create New Week'}
            >
              <form onSubmit={handleCreateWeek} className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Week Name</label>
                    <input 
                      type="text" 
                      required
                      placeholder="e.g. Week 1"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={weekForm.weekName || ''}
                      onChange={(e) => setWeekForm({...weekForm, weekName: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Month</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={weekForm.month}
                      onChange={(e) => setWeekForm({...weekForm, month: e.target.value})}
                    >
                      {['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'].map(m => (
                        <option key={m} value={m}>{m}</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Start Date</label>
                    <input 
                      type="date" 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={weekForm.startDate || ''}
                      onChange={(e) => setWeekForm({...weekForm, startDate: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">End Date</label>
                    <input 
                      type="date" 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={weekForm.endDate || ''}
                      onChange={(e) => setWeekForm({...weekForm, endDate: e.target.value})}
                    />
                  </div>
                </div>

                {error && <p className="text-xs text-red-500">{error}</p>}
                {success && <p className="text-xs text-green-500">{success}</p>}

                <div className="flex gap-3 pt-4">
                  <Button type="submit" className="flex-1" isLoading={isSavingWeek}>
                    {isEditingWeek ? 'Update Week' : 'Save Week Record'}
                  </Button>
                  <Button 
                    type="button" 
                    variant="ghost" 
                    onClick={() => setIsWeekModalOpen(false)}
                  >
                    Cancel
                  </Button>
                </div>
              </form>
            </Modal>

            <div className="space-y-4">
              <h3 className="text-lg font-bold">Configured Weeks</h3>
              <Card className="overflow-x-auto">
                <table className="w-full text-left border-collapse border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50">
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Week</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Month</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Start Date</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">End Date</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {allSortedWeeks.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No week records found.
                        </td>
                      </tr>
                    ) : (
                      allSortedWeeks.map(wk => (
                        <tr key={wk.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-bold text-orange-600 border border-zinc-200">{wk.weekName}</td>
                          <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{wk.month}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{wk.startDate}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{wk.endDate}</td>
                          <td className="px-4 py-3 border border-zinc-200">
                            <div className="flex items-center gap-2">
                              {getPermissionLevel('weeks') === 'Complete' && (
                                <button 
                                  onClick={() => handleEditWeek(wk)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {getPermissionLevel('weeks') === 'Complete' && (
                                <button 
                                  onClick={() => wk.id && handleDeleteWeek(wk.id, wk.weekName)}
                                  className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </Card>

              {missingWeeksFromSales.length > 0 && (
                <div className="mt-8 space-y-4">
                  <div className="flex items-center gap-2 text-red-600">
                    <AlertCircle className="w-5 h-5" />
                    <h3 className="text-lg font-bold">Data Integrity Alert: Missing Week Configurations</h3>
                  </div>
                  <p className="text-sm text-zinc-600">
                    The following weeks are referenced in Sales records but are not configured in the Week Master. 
                    This causes sales for these weeks to be hidden in filtered views.
                  </p>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    {missingWeeksFromSales.map(weekName => (
                      <Card key={weekName} className="p-4 border-red-100 bg-red-50/30 flex items-center justify-between">
                        <div>
                          <p className="text-sm font-bold text-zinc-900">{weekName}</p>
                          <p className="text-xs text-zinc-500">Used in {sales.filter(s => s.week === weekName).length} sales</p>
                        </div>
                        <Button 
                          size="sm"
                          variant="outline"
                          className="text-xs border-red-200 hover:bg-red-50 text-red-600"
                          onClick={() => {
                            setIsEditingWeek(false);
                            setWeekForm({
                              weekName: weekName,
                              month: 'February', // Default to Feb as per user request context
                              startDate: '',
                              endDate: ''
                            });
                            setIsWeekModalOpen(true);
                          }}
                        >
                          Configure Now
                        </Button>
                      </Card>
                    ))}
                  </div>
                </div>
              )}
            </div>
          </div>
        )}

        {activeTab === 'bdes' && hasPermission('bdes') && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Users className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">BDE Master</h2>
              </div>
              {getPermissionLevel('bdes') === 'Complete' && (
                <Button 
                  onClick={() => {
                    setIsEditingBDE(false);
                    setBDEForm({ name: '', phone: '', email: '' });
                    setIsBDEModalOpen(true);
                  }} 
                  className="bg-orange-600"
                >
                  <Plus className="w-4 h-4" />
                  Add BDE
                </Button>
              )}
            </div>

            <Modal 
              isOpen={isBDEModalOpen} 
              onClose={() => setIsBDEModalOpen(false)}
              title={isEditingBDE ? 'Edit BDE' : 'Create New BDE'}
            >
              <form onSubmit={handleCreateBDE} className="space-y-6">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">BDE Name</label>
                    <input 
                      type="text" 
                      required
                      placeholder="Enter BDE Name"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={bdeForm.name || ''}
                      onChange={(e) => setBDEForm({...bdeForm, name: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Phone No</label>
                    <input 
                      type="tel" 
                      placeholder="Enter Phone Number"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={bdeForm.phone || ''}
                      onChange={(e) => setBDEForm({...bdeForm, phone: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1 md:col-span-2">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Email ID</label>
                    <input 
                      type="email" 
                      placeholder="Enter Email ID"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={bdeForm.email || ''}
                      onChange={(e) => setBDEForm({...bdeForm, email: e.target.value})}
                    />
                  </div>
                </div>

                {error && <p className="text-xs text-red-500">{error}</p>}
                {success && <p className="text-xs text-green-500">{success}</p>}

                <div className="flex gap-3 pt-4">
                  <Button type="submit" className="flex-1" isLoading={isSavingBDE}>
                    {isEditingBDE ? 'Update BDE' : 'Save BDE'}
                  </Button>
                  <Button 
                    type="button" 
                    variant="ghost" 
                    onClick={() => setIsBDEModalOpen(false)}
                  >
                    Cancel
                  </Button>
                </div>
              </form>
            </Modal>

            <div className="space-y-4">
              <h3 className="text-lg font-bold">Configured BDEs</h3>
              <Card className="overflow-x-auto">
                <table className="w-full text-left border-collapse border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50">
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Name</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Phone</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Email</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bdes.length === 0 ? (
                      <tr>
                        <td colSpan={4} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No BDE records found.
                        </td>
                      </tr>
                    ) : (
                      [...bdes].sort((a, b) => a.name.localeCompare(b.name)).map(bde => (
                        <tr key={bde.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-bold text-orange-600 border border-zinc-200">{bde.name}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{bde.phone || '-'}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{bde.email || '-'}</td>
                          <td className="px-4 py-3 border border-zinc-200">
                            <div className="flex items-center gap-2">
                              {getPermissionLevel('bdes') === 'Complete' && (
                                <button 
                                  onClick={() => handleEditBDE(bde)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {getPermissionLevel('bdes') === 'Complete' && (
                                <button 
                                  onClick={() => bde.id && handleDeleteBDE(bde.id, bde.name)}
                                  className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              )}
                            </div>
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </Card>
            </div>
          </div>
        )}
        {activeTab === 'weeklyOverview' && hasPermission('weeklyOverview') && (
          <div className="space-y-8">
            {(() => {
              const activeEmployees = employees.filter(e => {
                const matchesStatus = e.status === 'Active';
                const level = getPermissionLevel('weeklyOverview');
                const canSee = level === 'Complete' || e.email?.toLowerCase() === profile?.email?.toLowerCase();
                const searchTerm = (weeklyOverviewSearch || '').toLowerCase().trim();
                const matchesSearch = !searchTerm || 
                  (e.name || '').toLowerCase().includes(searchTerm) ||
                  (e.employeeCode || '').toLowerCase().includes(searchTerm) ||
                  (e.bde || '').toLowerCase().includes(searchTerm);
                return matchesStatus && canSee && matchesSearch;
              });

              return (
                <>
                  <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
                    <div className="flex items-center gap-2">
                      <TrendingUp className="w-6 h-6 text-orange-600" />
                      <h2 className="text-3xl font-bold tracking-tight">Weekly Overview</h2>
                    </div>
                    <div className="flex flex-wrap items-center gap-3">
                      <div className="relative w-full md:w-64">
                        <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-400" />
                        <input 
                          type="text"
                          placeholder="Search agents..."
                          className="w-full pl-10 pr-4 py-2 bg-white border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm"
                          value={weeklyOverviewSearch}
                          onChange={(e) => setWeeklyOverviewSearch(e.target.value)}
                        />
                      </div>
                      <Button 
                        onClick={downloadWeeklyOverview}
                        variant="outline"
                        className="flex items-center gap-2"
                      >
                        <Download className="w-4 h-4" />
                        Download Excel
                      </Button>
                    </div>
                  </div>
                  <Card className="overflow-x-auto p-6">
                    <table className="w-full text-left border-collapse border border-zinc-200">
                      <thead>
                        <tr className="bg-zinc-100">
                          <th colSpan={6} className="px-4 py-3 text-xs font-bold text-zinc-600 uppercase border border-zinc-200 text-center">Total</th>
                          <th colSpan={sortedWeeks.length * 2} className="px-4 py-3 text-xs font-bold text-zinc-600 uppercase border border-zinc-200"></th>
                        </tr>
                        <tr className="bg-zinc-50">
                          <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Agent</th>
                          <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">BDE</th>
                          <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Weekly Target</th>
                          <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Joining Date</th>
                          <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200 bg-blue-50/50">Bravo 4</th>
                          <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200 bg-purple-50/50">Bravo 8</th>
                          {[...sortedWeeks].reverse().map((w) => (
                            <React.Fragment key={w.id}>
                              <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200 text-center bg-orange-50/50">
                                %
                              </th>
                              <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">
                                {w.weekName}
                              </th>
                            </React.Fragment>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {(() => {
                          if (activeEmployees.length === 0) return null;

                          const getOverallBravoAchievement = (numWeeks: number) => {
                            const targetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks);
                            let totalSales = 0;
                            let totalTarget = 0;

                            activeEmployees.forEach(emp => {
                              const agentSales = sales.filter(s => 
                                s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
                                (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase()) ||
                                s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
                                (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())
                              );
                              const weeklyBaseTarget = emp.target / 4;

                              targetWeeks.forEach(wk => {
                                const wkSales = agentSales.filter(s => s.week === wk.weekName).reduce((sum, s) => sum + getAgentCreditForSale(s, emp, 'packageValue'), 0);
                                totalSales += wkSales;

                                const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
                                const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
                                
                                let percentageVal = 0;
                                if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                                  const tenureWeek = wkIndex - joiningWkIndex + 1;
                                  if (tenureWeek <= 6) {
                                    percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
                                  } else {
                                    percentageVal = 100;
                                  }
                                }
                                totalTarget += (weeklyBaseTarget * percentageVal) / 100;
                              });
                            });

                            return totalTarget > 0 ? Math.round((totalSales / totalTarget) * 100) : 0;
                          };

                          const overallBravo4 = getOverallBravoAchievement(4);
                          const overallBravo8 = getOverallBravoAchievement(8);

                          const getTotalLastWeekTarget = () => {
                            const lastWeek = allSortedWeeks[allSortedWeeks.length - 1];
                            if (!lastWeek) return 0;
                            
                            let totalTarget = 0;
                            activeEmployees.forEach(emp => {
                              const weeklyBaseTarget = emp.target / 4;
                              const wkIndex = allSortedWeeks.findIndex(w => w.weekName === lastWeek.weekName);
                              const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
                              
                              let percentageVal = 0;
                              if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                                const tenureWeek = wkIndex - joiningWkIndex + 1;
                                if (tenureWeek <= 6) {
                                  percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
                                } else {
                                  percentageVal = 100;
                                }
                              }
                              totalTarget += (weeklyBaseTarget * percentageVal) / 100;
                            });
                            return totalTarget;
                          };

                          const totalLastWeekTarget = getTotalLastWeekTarget();

                          return (
                            <tr className="bg-zinc-100 font-bold border-b border-zinc-200">
                              <td className="px-4 py-3 text-sm border border-zinc-200">TOTAL</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">-</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">₹{totalLastWeekTarget.toLocaleString()}</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">-</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200 text-center text-blue-600">{overallBravo4}%</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200 text-center text-purple-600">{overallBravo8}%</td>
                              {[...sortedWeeks].reverse().map((w) => {
                                const totalWeekSales = activeEmployees.reduce((sum, emp) => {
                                  const weekSales = sales.filter(s => s.week === w.weekName);
                                  return sum + weekSales.reduce((sSum, s) => sSum + getAgentCreditForSale(s, emp, 'packageValue'), 0);
                                }, 0);
                                
                                const totalWeekSalesData = sales.filter(s => s.week === w.weekName && activeEmployees.some(e => 
                                  (e.name.toLowerCase().trim() === s.agent?.toLowerCase().trim()) || (s.agentEmail && e.email?.toLowerCase() === s.agentEmail.toLowerCase()) ||
                                  (e.name.toLowerCase().trim() === s.associateBde?.toLowerCase().trim()) || (s.associateBdeEmail && e.email?.toLowerCase() === s.associateBdeEmail.toLowerCase())
                                ));
                                return (
                                  <React.Fragment key={`total-${w.id}`}>
                                    <td className="px-4 py-3 text-sm border border-zinc-200 text-center bg-orange-50/20">-</td>
                                    <td 
                                      className="px-4 py-3 text-sm border border-zinc-200 text-right cursor-pointer hover:bg-orange-50 transition-colors"
                                      onClick={() => {
                                        if (totalWeekSales > 0) {
                                          setSummaryDetailModal({
                                            isOpen: true,
                                            title: `Total Sales - ${w.weekName}`,
                                            sales: totalWeekSalesData
                                          });
                                        }
                                      }}
                                    >
                                      {totalWeekSales > 0 ? totalWeekSales.toLocaleString() : '-'}
                                    </td>
                                  </React.Fragment>
                                );
                              })}
                            </tr>
                          );
                        })()}
                        {activeEmployees.map(emp => {
                          const agentSales = sales.filter(s => 
                            s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
                            (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase()) ||
                            s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim() ||
                            (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())
                          );
                          
                          const getBravoAchievement = (numWeeks: number) => {
                            const targetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks);
                            let totalSales = 0;
                            let totalTarget = 0;
                            const weeklyBaseTarget = emp.target / 4;

                            targetWeeks.forEach(wk => {
                              const wkSales = agentSales.filter(s => s.week === wk.weekName).reduce((sum, s) => sum + getAgentCreditForSale(s, emp, 'packageValue'), 0);
                              totalSales += wkSales;

                        const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
                        const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
                        
                        let percentageVal = 0;
                        if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                          const tenureWeek = wkIndex - joiningWkIndex + 1;
                          if (tenureWeek <= 6) {
                            percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
                          } else {
                            percentageVal = 100;
                          }
                        }
                        totalTarget += (weeklyBaseTarget * percentageVal) / 100;
                      });

                      return totalTarget > 0 ? Math.round((totalSales / totalTarget) * 100) : 0;
                    };

                    const bravo4 = getBravoAchievement(4);
                    const bravo8 = getBravoAchievement(8);

                    const getLastWeekTarget = () => {
                      const lastWeek = allSortedWeeks[allSortedWeeks.length - 1];
                      if (!lastWeek) return 0;
                      
                      const weeklyBaseTarget = emp.target / 4;
                      const wkIndex = allSortedWeeks.findIndex(w => w.weekName === lastWeek.weekName);
                      const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
                      
                      let percentageVal = 0;
                      if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                        const tenureWeek = wkIndex - joiningWkIndex + 1;
                        if (tenureWeek <= 6) {
                          percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
                        } else {
                          percentageVal = 100;
                        }
                      }
                      return (weeklyBaseTarget * percentageVal) / 100;
                    };

                    const lastWeekTarget = getLastWeekTarget();

                    return (
                      <tr key={emp.id} className="hover:bg-zinc-50 transition-colors">
                        <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{emp.name}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200">{emp.bde}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200">₹{lastWeekTarget.toLocaleString()}</td>
                        <td className="px-4 py-3 text-sm border border-zinc-200">{emp.joiningDate}</td>
                        <td 
                          className="px-4 py-3 text-sm font-bold text-blue-600 border border-zinc-200 bg-blue-50/10 text-center cursor-pointer hover:bg-blue-50 transition-colors"
                          onClick={() => handleOpenBravoCalculation(emp, 4)}
                        >
                          {bravo4}%
                        </td>
                        <td 
                          className="px-4 py-3 text-sm font-bold text-purple-600 border border-zinc-200 bg-purple-50/10 text-center cursor-pointer hover:bg-purple-50 transition-colors"
                          onClick={() => handleOpenBravoCalculation(emp, 8)}
                        >
                          {bravo8}%
                        </td>
                        {[...sortedWeeks].reverse().map((w) => {
                          const weekSales = agentSales.filter(s => s.week === w.weekName).reduce((sum, s) => {
                            let credit = 0;
                            if ((s.agent?.toLowerCase().trim() === emp.name.toLowerCase().trim()) || (emp.email && s.agentEmail?.toLowerCase() === emp.email.toLowerCase())) {
                              credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                            }
                            if ((s.associateBde?.toLowerCase().trim() === emp.name.toLowerCase().trim()) || (emp.email && s.associateBdeEmail?.toLowerCase() === emp.email.toLowerCase())) {
                              credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                            }
                            return sum + credit;
                          }, 0);
                          const weekSalesData = agentSales.filter(s => s.week === w.weekName);
                          const wkIndex = allSortedWeeks.findIndex(sw => sw.weekName === w.weekName);
                          const joiningWeekIndex = allSortedWeeks.findIndex(sw => sw.weekName === emp.joiningWeek);
                          
                          let percentage = '-';
                          if (joiningWeekIndex !== -1 && wkIndex >= joiningWeekIndex) {
                            const tenureWeek = wkIndex - joiningWeekIndex + 1;
                            if (tenureWeek <= 6) {
                              percentage = `${(emp as any)[`week${tenureWeek}Target`]}%`;
                            } else {
                              percentage = '100%';
                            }
                          }

                          return (
                            <React.Fragment key={w.id}>
                              <td className="px-4 py-3 text-sm border border-zinc-200 text-center font-bold text-orange-600 bg-orange-50/20">
                                {percentage}
                              </td>
                              <td 
                                className="px-4 py-3 text-sm border border-zinc-200 text-right cursor-pointer hover:bg-orange-50 transition-colors"
                                onClick={() => {
                                  if (weekSales > 0) {
                                    setSummaryDetailModal({
                                      isOpen: true,
                                      title: `${emp.name} - ${w.weekName} Sales`,
                                      sales: weekSalesData
                                    });
                                  }
                                }}
                              >
                                {weekSales > 0 ? weekSales.toLocaleString() : '-'}
                              </td>
                            </React.Fragment>
                          );
                        })}
                      </tr>
                        );
                      })}
                      {activeEmployees.length === 0 && (
                        <tr>
                          <td colSpan={6 + sortedWeeks.length * 2} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                            No active agents found.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </Card>
              </>
            );
          })()}
        </div>
      )}
        {activeTab === 'sales' && hasPermission('sales') && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Sales Master</h2>
              </div>
              <div className="flex items-center gap-3">
                <div className="relative w-64">
                  <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-400" />
                  <input 
                    type="text"
                    placeholder="Search Trip ID..."
                    className="w-full pl-9 pr-4 py-2 bg-white border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm"
                    value={tripIdSearch}
                    onChange={(e) => setTripIdSearch(e.target.value)}
                    onKeyDown={(e) => e.key === 'Enter' && handleTripIdSearch()}
                  />
                </div>
                <Button 
                  onClick={handleTripIdSearch}
                  isLoading={isSearchingTripId}
                  variant="outline"
                  className="bg-white"
                >
                  Search
                </Button>
                {!isSalesLoaded && (
                  <Button 
                    onClick={() => setIsSalesLoaded(true)}
                    className="bg-zinc-900 hover:bg-zinc-800 flex items-center gap-2"
                  >
                    <Database className="w-4 h-4" />
                    Load All Sales
                  </Button>
                )}
                {getPermissionLevel('sales') === 'Complete' && (
                  <Button 
                    onClick={downloadSalesMaster}
                    variant="outline"
                    className="flex items-center gap-2"
                  >
                    <Download className="w-4 h-4" />
                    Download Excel
                  </Button>
                )}
                {getPermissionLevel('sales') === 'Complete' && (
                  <Button 
                    onClick={() => {
                      setIsEditingSales(false);
                      setSalesForm({
                        week: '',
                        date: new Date().toISOString().split('T')[0],
                        guestName: '',
                        agent: isAgent ? profile?.displayName || '' : '',
                        agentEmail: isAgent ? profile?.email || '' : '',
                        agentPercentage: 100,
                        bde: profile?.role === 'BDE' ? profile?.displayName || '' : '',
                        bdeEmail: profile?.role === 'BDE' ? profile?.email || '' : '',
                        associateBde: '',
                        associateBdeEmail: '',
                        assocBdePercentage: 0,
                        destination: '',
                        tripId: '',
                        tripDate: '',
                        packageValue: undefined,
                        lessThan10Percent20k: '',
                        advanceCN: '',
                        salesBy: 'Inhouse',
                        ppMargin: undefined,
                        noOfPax: undefined,
                        totalMargin: undefined,
                        flight: undefined,
                        source: '',
                        converted: 'Yes',
                        tasksPending: '',
                        pHotel: '',
                        pFlight: '',
                        land: '',
                        hflIssue: '',
                        workPending: '',
                        remarks: '',
                      });
                      setIsSalesModalOpen(true);
                    }} 
                    className="bg-orange-600"
                  >
                    <Plus className="w-4 h-4" />
                    Add Sale
                  </Button>
                )}
            </div>
          </div>

            {!isSalesLoaded && (
              <div className="bg-orange-50 border border-orange-100 p-4 rounded-xl flex items-center gap-3 text-orange-800">
                <Info className="w-5 h-5 text-orange-600" />
                <p className="text-sm">
                  Showing only recent sales (last 5 days) to save database reads. 
                  Click <strong>"Load All Sales"</strong> to see the full history (last {isPrivileged ? '60' : '30'} days).
                </p>
              </div>
            )}

            <Modal 
              isOpen={isSalesModalOpen} 
              onClose={() => setIsSalesModalOpen(false)}
              title={isEditingSales ? 'Edit Sale' : 'Create New Sale'}
            >
              <form onSubmit={handleCreateSales} className="space-y-6 max-h-[70vh] overflow-y-auto px-1">
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Week</label>
                    <select 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.week || ''}
                      onChange={(e) => setSalesForm({...salesForm, week: e.target.value})}
                    >
                      <option value="">Select Week</option>
                      {allSortedWeeks.map(wk => (
                        <option key={wk.id} value={wk.weekName}>{wk.weekName} ({wk.month})</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Date</label>
                    <input 
                      type="date" 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.date || ''}
                      onChange={(e) => setSalesForm({...salesForm, date: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Guest Name</label>
                    <input 
                      type="text" 
                      required
                      placeholder="Enter Guest Name"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.guestName || ''}
                      onChange={(e) => setSalesForm({...salesForm, guestName: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Agent</label>
                    <SearchableSelect
                      required
                      disabled={isAgent && !isAdmin}
                      options={[
                        { value: '', label: 'Select Agent' },
                        ...employees
                          .sort((a, b) => a.name.localeCompare(b.name))
                          .map(emp => ({ value: emp.name, label: emp.name, key: emp.id }))
                      ]}
                      value={salesForm.agent || ''}
                      onChange={(val) => {
                        const emp = employees.find(e => e.name === val);
                        setSalesForm({...salesForm, agent: val, agentEmail: emp?.email || ''});
                      }}
                      placeholder="Select Agent"
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">% Agent</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.agentPercentage ?? 100}
                      onChange={(e) => setSalesForm({...salesForm, agentPercentage: Number(e.target.value)})}
                    >
                      <option value="0">0</option>
                      <option value="30">30</option>
                      <option value="50">50</option>
                      <option value="70">70</option>
                      <option value="100">100</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">BDE</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.bde || ''}
                      onChange={(e) => {
                        const val = e.target.value;
                        const bde = bdes.find(b => b.name === val);
                        setSalesForm({...salesForm, bde: val, bdeEmail: bde?.email || ''});
                      }}
                    >
                      <option value="">Select BDE</option>
                      {bdes.map(bde => (
                        <option key={bde.id} value={bde.name}>{bde.name}</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Associate BDE</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.associateBde || ''}
                      onChange={(e) => {
                        const val = e.target.value;
                        const emp = employees.find(e => e.name === val);
                        setSalesForm({...salesForm, associateBde: val, associateBdeEmail: emp?.email || ''});
                      }}
                    >
                      <option value="">Select Associate BDE</option>
                      {employees.filter(e => e.isAssocBDE).map(emp => (
                        <option key={emp.id} value={emp.name}>{emp.name}</option>
                      ))}
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">% Assoc BDE</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.assocBdePercentage ?? 0}
                      onChange={(e) => setSalesForm({...salesForm, assocBdePercentage: Number(e.target.value)})}
                    >
                      <option value="0">0</option>
                      <option value="30">30</option>
                      <option value="50">50</option>
                      <option value="70">70</option>
                      <option value="100">100</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Destination</label>
                    <input 
                      type="text" 
                      required
                      placeholder="Enter Destination"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.destination || ''}
                      onChange={(e) => setSalesForm({...salesForm, destination: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Trip ID</label>
                    <input 
                      type="text" 
                      required
                      placeholder="Enter Trip ID"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.tripId || ''}
                      onChange={(e) => setSalesForm({...salesForm, tripId: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Trip Date</label>
                    <input 
                      type="date" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.tripDate || ''}
                      onChange={(e) => setSalesForm({...salesForm, tripDate: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Package Value</label>
                    <input 
                      type="number" 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.packageValue ?? ''}
                      onChange={(e) => setSalesForm({...salesForm, packageValue: e.target.value === '' ? undefined : Number(e.target.value)})}
                      onFocus={(e) => e.target.select()}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Less Than 10% / 20k</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.lessThan10Percent20k || ''}
                      onChange={(e) => setSalesForm({...salesForm, lessThan10Percent20k: e.target.value})}
                    >
                      <option value="">Select Option</option>
                      <option value="Issue">Issue</option>
                      <option value="Done">Done</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Advance/CN</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.advanceCN || ''}
                      onChange={(e) => setSalesForm({...salesForm, advanceCN: e.target.value})}
                    >
                      <option value="">Select Option</option>
                      <option value="Advance">Advance</option>
                      <option value="Credit Note">Credit Note</option>
                      <option value="Confirmed">Confirmed</option>
                      <option value="Cancel">Cancel</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Sales By</label>
                    <select 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.salesBy || 'Inhouse'}
                      onChange={(e) => setSalesForm({...salesForm, salesBy: e.target.value as 'Inhouse' | 'Branch' | 'Franchisee Sales'})}
                    >
                      <option value="Inhouse">Inhouse</option>
                      <option value="Branch">Branch</option>
                      <option value="Franchisee Sales">Franchisee Sales</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">PP Margin</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.ppMargin ?? ''}
                      onChange={(e) => setSalesForm({...salesForm, ppMargin: e.target.value === '' ? undefined : Number(e.target.value)})}
                      onFocus={(e) => e.target.select()}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">No. of Pax</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.noOfPax ?? ''}
                      onChange={(e) => setSalesForm({...salesForm, noOfPax: e.target.value === '' ? undefined : Number(e.target.value)})}
                      onFocus={(e) => e.target.select()}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Total Margin</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.totalMargin ?? ''}
                      onChange={(e) => setSalesForm({...salesForm, totalMargin: e.target.value === '' ? undefined : Number(e.target.value)})}
                      onFocus={(e) => e.target.select()}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Flight</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.flight ?? ''}
                      onChange={(e) => setSalesForm({...salesForm, flight: e.target.value === '' ? undefined : Number(e.target.value)})}
                      onFocus={(e) => e.target.select()}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Source</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.source || ''}
                      onChange={(e) => setSalesForm({...salesForm, source: e.target.value})}
                    >
                      <option value="">Select Source</option>
                      <option value="Social">Social</option>
                      <option value="Refral">Refral</option>
                      <option value="Direct">Direct</option>
                      <option value="SC">SC</option>
                      <option value="Done">Done</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Converted</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.converted || 'Yes'}
                      onChange={(e) => setSalesForm({...salesForm, converted: e.target.value})}
                    >
                      <option value="Yes">Yes</option>
                      <option value="No">No</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Tasks Pending</label>
                    <select 
                      className={cn(
                        "w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-orange-500 outline-none transition-colors",
                        salesForm.tasksPending === 'Yes' ? "bg-red-500 text-white border-red-600" : "bg-zinc-50 border-zinc-200"
                      )}
                      value={salesForm.tasksPending || ''}
                      onChange={(e) => setSalesForm({...salesForm, tasksPending: e.target.value})}
                    >
                      <option value="" className="text-zinc-900 bg-white">Select Option</option>
                      <option value="Yes" className="text-zinc-900 bg-white">Yes</option>
                      <option value="No" className="text-zinc-900 bg-white">No</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">P - Hotel</label>
                    <select 
                      className={cn(
                        "w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-orange-500 outline-none transition-colors",
                        salesForm.pHotel === 'Issue' ? "bg-red-500 text-white border-red-600" : "bg-zinc-50 border-zinc-200"
                      )}
                      value={salesForm.pHotel || ''}
                      onChange={(e) => setSalesForm({...salesForm, pHotel: e.target.value})}
                    >
                      <option value="" className="text-zinc-900 bg-white">Select Option</option>
                      <option value="Issue" className="text-zinc-900 bg-white">Issue</option>
                      <option value="Done" className="text-zinc-900 bg-white">Done</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">P Flight</label>
                    <select 
                      className={cn(
                        "w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-orange-500 outline-none transition-colors",
                        salesForm.pFlight === 'Issue' ? "bg-red-500 text-white border-red-600" : "bg-zinc-50 border-zinc-200"
                      )}
                      value={salesForm.pFlight || ''}
                      onChange={(e) => setSalesForm({...salesForm, pFlight: e.target.value})}
                    >
                      <option value="" className="text-zinc-900 bg-white">Select Option</option>
                      <option value="Issue" className="text-zinc-900 bg-white">Issue</option>
                      <option value="Done" className="text-zinc-900 bg-white">Done</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Land</label>
                    <select 
                      className={cn(
                        "w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-orange-500 outline-none transition-colors",
                        salesForm.land === 'Issue' ? "bg-red-500 text-white border-red-600" : "bg-zinc-50 border-zinc-200"
                      )}
                      value={salesForm.land || ''}
                      onChange={(e) => setSalesForm({...salesForm, land: e.target.value})}
                    >
                      <option value="" className="text-zinc-900 bg-white">Select Option</option>
                      <option value="Issue" className="text-zinc-900 bg-white">Issue</option>
                      <option value="Done" className="text-zinc-900 bg-white">Done</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">H+F+L issue</label>
                    <select 
                      className={cn(
                        "w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-orange-500 outline-none transition-colors",
                        salesForm.hflIssue === 'Issue' ? "bg-red-500 text-white border-red-600" : "bg-zinc-50 border-zinc-200"
                      )}
                      value={salesForm.hflIssue || ''}
                      onChange={(e) => setSalesForm({...salesForm, hflIssue: e.target.value})}
                    >
                      <option value="" className="text-zinc-900 bg-white">Select Option</option>
                      <option value="Issue" className="text-zinc-900 bg-white">Issue</option>
                      <option value="Done" className="text-zinc-900 bg-white">Done</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Work Pending</label>
                    <select 
                      className={cn(
                        "w-full px-4 py-2 border rounded-lg focus:ring-2 focus:ring-orange-500 outline-none transition-colors",
                        salesForm.workPending === 'Issue' ? "bg-red-500 text-white border-red-600" : "bg-zinc-50 border-zinc-200"
                      )}
                      value={salesForm.workPending || ''}
                      onChange={(e) => setSalesForm({...salesForm, workPending: e.target.value})}
                    >
                      <option value="" className="text-zinc-900 bg-white">Select Option</option>
                      <option value="Issue" className="text-zinc-900 bg-white">Issue</option>
                      <option value="Done" className="text-zinc-900 bg-white">Done</option>
                    </select>
                  </div>
                </div>

                <div className="space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Remarks</label>
                  <textarea 
                    rows={3}
                    placeholder="Enter Remarks"
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none resize-none"
                    value={salesForm.remarks || ''}
                    onChange={(e) => setSalesForm({...salesForm, remarks: e.target.value})}
                  />
                </div>

                {error && <p className="text-xs text-red-500">{error}</p>}
                {success && <p className="text-xs text-green-500">{success}</p>}

                <div className="flex gap-3 pt-4">
                  <Button type="submit" className="flex-1" isLoading={isSavingSales}>
                    {isEditingSales ? 'Update Sale' : 'Save Sale Record'}
                  </Button>
                  <Button 
                    type="button" 
                    variant="ghost" 
                    onClick={() => setIsSalesModalOpen(false)}
                  >
                    Cancel
                  </Button>
                </div>
              </form>
            </Modal>

            <div className="space-y-4">
              <div className="flex flex-col md:flex-row gap-4 items-end bg-white p-4 rounded-xl border border-zinc-200 shadow-sm">
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by Month</label>
                  <select 
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={salesMonthFilter}
                    onChange={(e) => {
                      setSalesMonthFilter(e.target.value);
                      setSalesWeekFilter(''); // Clear week filter when month changes
                    }}
                  >
                    <option value="">All Months</option>
                    {allMonths.map(month => (
                      <option key={month} value={month}>{month}</option>
                    ))}
                  </select>
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by Week</label>
                  <select 
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={salesWeekFilter}
                    onChange={(e) => setSalesWeekFilter(e.target.value)}
                  >
                    <option value="">All Weeks</option>
                    {allSortedWeeks
                      .filter(wk => !salesMonthFilter || wk.month === salesMonthFilter)
                      .map(wk => (
                        <option key={wk.id} value={wk.weekName}>{wk.weekName} ({wk.month})</option>
                      ))}
                  </select>
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by Agent</label>
                  <SearchableSelect
                    options={[
                      ...(getPermissionLevel('sales') === 'Complete' ? [{ value: '', label: 'All Agents' }] : []),
                      ...employees
                        .filter(e => getPermissionLevel('sales') === 'Complete' || (e.email && e.email.toLowerCase() === profile?.email?.toLowerCase()))
                        .sort((a, b) => a.name.localeCompare(b.name))
                        .map(emp => ({ value: emp.name, label: emp.name, key: emp.id }))
                    ]}
                    value={salesAgentFilter}
                    onChange={setSalesAgentFilter}
                    placeholder="All Agents"
                  />
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by BDE</label>
                  <SearchableSelect
                    options={[
                      { value: '', label: 'All BDEs' },
                      ...bdes.sort((a, b) => a.name.localeCompare(b.name)).map(bde => ({ value: bde.name, label: bde.name, key: bde.id }))
                    ]}
                    value={salesBdeFilter}
                    onChange={setSalesBdeFilter}
                    placeholder="All BDEs"
                  />
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by Sales Type</label>
                  <select 
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={salesByFilter}
                    onChange={(e) => setSalesByFilter(e.target.value as any)}
                  >
                    <option value="All">All Types</option>
                    <option value="Inhouse">Inhouse</option>
                    <option value="Branch">Branch</option>
                    <option value="Franchisee Sales">Franchisee Sales</option>
                  </select>
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Search</label>
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-400" />
                    <input 
                      type="text"
                      placeholder="Search entries..."
                      className="w-full pl-9 pr-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesSearchQuery}
                      onChange={(e) => setSalesSearchQuery(e.target.value)}
                    />
                  </div>
                </div>
                <Button 
                  variant="ghost" 
                  onClick={() => { 
                    setSalesAgentFilter(''); 
                    setSalesBdeFilter(''); 
                    setSalesWeekFilter(''); 
                    setSalesByFilter('All');
                    setSalesSearchQuery('');
                  }}
                  className="text-zinc-500"
                >
                  <Filter className="w-4 h-4 mr-2" />
                  Clear Filters
                </Button>
              </div>

              {/* Summary Cards */}
              <div className="relative">
                {isLoadingAggregates && (
                  <div className="absolute inset-0 bg-white/50 backdrop-blur-[1px] z-10 flex items-center justify-center rounded-xl">
                    <Loader2 className="w-8 h-8 animate-spin text-orange-600" />
                  </div>
                )}
                <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                  {(() => {
                    const stats = dashboardAggregates || salesStats;
                    return (
                      <>
                        <Card className="p-4 bg-orange-50 border-orange-100">
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-orange-100 rounded-lg">
                              <TrendingUp className="w-5 h-5 text-orange-600" />
                            </div>
                            <div className="flex-1">
                              <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Inhouse Sales</p>
                              <p className="text-2xl font-bold text-zinc-900">{stats.inhouse.total}</p>
                              <div className="mt-2 pt-2 border-t border-orange-200 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                                <div className="text-orange-700">Adv: {stats.inhouse.advance}</div>
                                <div className="text-orange-700">CN: {stats.inhouse.cn}</div>
                                <div className="text-orange-700">Conf: {stats.inhouse.confirmed}</div>
                                <div className="text-orange-700">Can: {stats.inhouse.cancel}</div>
                              </div>
                            </div>
                          </div>
                        </Card>
                        <Card className="p-4 bg-orange-50 border-orange-100">
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-orange-100 rounded-lg">
                              <TrendingUp className="w-5 h-5 text-orange-600" />
                            </div>
                            <div className="flex-1">
                              <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Branch Sales</p>
                              <p className="text-2xl font-bold text-zinc-900">{stats.branch.total}</p>
                              <div className="mt-2 pt-2 border-t border-orange-200 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                                <div className="text-orange-700">Adv: {stats.branch.advance}</div>
                                <div className="text-orange-700">CN: {stats.branch.cn}</div>
                                <div className="text-orange-700">Conf: {stats.branch.confirmed}</div>
                                <div className="text-orange-700">Can: {stats.branch.cancel}</div>
                              </div>
                            </div>
                          </div>
                        </Card>
                        <Card className="p-4 bg-orange-50 border-orange-100">
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-orange-100 rounded-lg">
                              <TrendingUp className="w-5 h-5 text-orange-600" />
                            </div>
                            <div className="flex-1">
                              <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Franchisee Sales</p>
                              <p className="text-2xl font-bold text-zinc-900">{stats.franchisee.total}</p>
                              <div className="mt-2 pt-2 border-t border-orange-200 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                                <div className="text-orange-700">Adv: {stats.franchisee.advance}</div>
                                <div className="text-orange-700">CN: {stats.franchisee.cn}</div>
                                <div className="text-orange-700">Conf: {stats.franchisee.confirmed}</div>
                                <div className="text-orange-700">Can: {stats.franchisee.cancel}</div>
                              </div>
                            </div>
                          </div>
                        </Card>
                        <Card className="p-4 bg-zinc-900 text-white">
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-zinc-800 rounded-lg">
                              <IndianRupee className="w-5 h-5 text-orange-400" />
                            </div>
                            <div className="flex-1">
                              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Inhouse Value</p>
                              <p className="text-2xl font-bold text-white">₹{stats.inhouse.value.toLocaleString()}</p>
                              <div className="mt-2 pt-2 border-t border-zinc-800 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                                <div className="text-zinc-400">Adv: ₹{stats.inhouse.advanceValue.toLocaleString()}</div>
                                <div className="text-zinc-400">CN: ₹{stats.inhouse.cnValue.toLocaleString()}</div>
                                <div className="text-zinc-400">Conf: ₹{stats.inhouse.confirmedValue.toLocaleString()}</div>
                                <div className="text-zinc-400">Can: ₹{stats.inhouse.cancelValue.toLocaleString()}</div>
                              </div>
                            </div>
                          </div>
                        </Card>
                        <Card className="p-4 bg-zinc-900 text-white">
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-zinc-800 rounded-lg">
                              <IndianRupee className="w-5 h-5 text-orange-400" />
                            </div>
                            <div className="flex-1">
                              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Branch Value</p>
                              <p className="text-2xl font-bold text-white">₹{stats.branch.value.toLocaleString()}</p>
                              <div className="mt-2 pt-2 border-t border-zinc-800 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                                <div className="text-zinc-400">Adv: ₹{stats.branch.advanceValue.toLocaleString()}</div>
                                <div className="text-zinc-400">CN: ₹{stats.branch.cnValue.toLocaleString()}</div>
                                <div className="text-zinc-400">Conf: ₹{stats.branch.confirmedValue.toLocaleString()}</div>
                                <div className="text-zinc-400">Can: ₹{stats.branch.cancelValue.toLocaleString()}</div>
                              </div>
                            </div>
                          </div>
                        </Card>
                        <Card className="p-4 bg-zinc-900 text-white">
                          <div className="flex items-center gap-3">
                            <div className="p-2 bg-zinc-800 rounded-lg">
                              <IndianRupee className="w-5 h-5 text-orange-400" />
                            </div>
                            <div className="flex-1">
                              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Franchisee Value</p>
                              <p className="text-2xl font-bold text-white">₹{stats.franchisee.value.toLocaleString()}</p>
                              <div className="mt-2 pt-2 border-t border-zinc-800 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                                <div className="text-zinc-400">Adv: ₹{stats.franchisee.advanceValue.toLocaleString()}</div>
                                <div className="text-zinc-400">CN: ₹{stats.franchisee.cnValue.toLocaleString()}</div>
                                <div className="text-zinc-400">Conf: ₹{stats.franchisee.confirmedValue.toLocaleString()}</div>
                                <div className="text-zinc-400">Can: ₹{stats.franchisee.cancelValue.toLocaleString()}</div>
                              </div>
                            </div>
                          </div>
                        </Card>
                      </>
                    );
                  })()}
                </div>
              </div>

              <h3 className="text-lg font-bold">Sales Records</h3>
              <Card className="overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[2000px] border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50">
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase sticky left-0 bg-zinc-50 z-20 border border-zinc-200 min-w-[100px]">Actions</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200 text-center">Log</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase sticky left-[100px] bg-zinc-50 z-20 border border-zinc-200">Guest Name</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Week</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Date</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Agent</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">% Agent</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">BDE</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Assoc BDE</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">% Assoc BDE</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Destination</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Trip ID</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Trip Date</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Pkg Value</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Less than 10% / 20k</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Advance/CN</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Sales By</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">PP Margin</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Pax</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Total Margin</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Flight</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Source</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Converted</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Tasks Pending</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">P - Hotel</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">P Flight</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Land</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">HFL Issue</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Work Pending</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Remarks</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSales.length === 0 ? (
                      <tr>
                        <td colSpan={28} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No sales records found.
                        </td>
                      </tr>
                    ) : (
                      filteredSales.map(sale => (
                        <tr key={sale.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 border border-zinc-200 sticky left-0 bg-white z-10 min-w-[100px]">
                            <div className="flex items-center gap-2">
                              {getPermissionLevel('sales') === 'Complete' && (
                                <button 
                                  onClick={() => handleEditSales(sale)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {getPermissionLevel('sales') === 'Complete' && (
                                <button 
                                  onClick={() => sale.id && handleDeleteSales(sale.id, sale.guestName)}
                                  className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              )}
                            </div>
                          </td>
                          <td className="px-4 py-3 border border-zinc-200 text-center">
                            <button 
                              onClick={() => handleOpenRemarks(sale)}
                              className={cn(
                                "p-1.5 rounded-lg transition-all flex items-center gap-1 mx-auto",
                                isRemarkUnread(sale)
                                  ? "bg-orange-600 text-white shadow-lg shadow-orange-600/20 animate-pulse" 
                                  : "bg-zinc-100 text-zinc-400 hover:bg-zinc-200 hover:text-zinc-600"
                              )}
                            >
                              <MessageSquare className="w-3.5 h-3.5" />
                              <span className="text-[10px] font-black">{sale.remarksCount || 0}</span>
                            </button>
                          </td>
                          <td className="px-4 py-3 text-sm font-bold text-orange-600 sticky left-[100px] bg-white z-10 border border-zinc-200">{sale.guestName}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.week}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.date}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.agent}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.agentPercentage}%</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.bde}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.associateBde}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.assocBdePercentage}%</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.destination}</td>
                          <td className="px-4 py-3 text-sm font-mono border border-zinc-200">{sale.tripId}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.tripDate}</td>
                          <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{sale.packageValue !== undefined ? `₹${sale.packageValue.toLocaleString()}` : ''}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.lessThan10Percent20k}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.advanceCN}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">
                            <span className={cn(
                              "px-2 py-0.5 rounded text-[10px] font-bold uppercase",
                              sale.salesBy === 'Branch' ? "bg-purple-100 text-purple-700" : 
                              sale.salesBy === 'Franchisee Sales' ? "bg-green-100 text-green-700" : "bg-blue-100 text-blue-700"
                            )}>
                              {sale.salesBy || 'Inhouse'}
                            </span>
                          </td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.ppMargin !== undefined ? `₹${sale.ppMargin.toLocaleString()}` : ''}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.noOfPax ?? ''}</td>
                          <td className="px-4 py-3 text-sm font-bold text-green-600 border border-zinc-200">{sale.totalMargin !== undefined ? `₹${sale.totalMargin.toLocaleString()}` : ''}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.flight !== undefined ? `₹${sale.flight.toLocaleString()}` : ''}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.source}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">
                            <span className={cn(
                              "px-2 py-1 rounded-full text-[10px] font-bold uppercase",
                              sale.converted === 'Yes' ? "bg-green-100 text-green-700" : 
                              sale.converted === 'No' ? "bg-red-100 text-red-700" : "bg-yellow-100 text-yellow-700"
                            )}>
                              {sale.converted}
                            </span>
                          </td>
                          <td className={cn(
                            "px-4 py-3 text-sm border border-zinc-200",
                            sale.tasksPending === 'Yes' && "bg-red-500 text-white font-bold"
                          )}>{sale.tasksPending}</td>
                          <td className={cn(
                            "px-4 py-3 text-sm border border-zinc-200",
                            sale.pHotel === 'Issue' && "bg-red-500 text-white font-bold"
                          )}>{sale.pHotel}</td>
                          <td className={cn(
                            "px-4 py-3 text-sm border border-zinc-200",
                            sale.pFlight === 'Issue' && "bg-red-500 text-white font-bold"
                          )}>{sale.pFlight}</td>
                          <td className={cn(
                            "px-4 py-3 text-sm border border-zinc-200",
                            sale.land === 'Issue' && "bg-red-500 text-white font-bold"
                          )}>{sale.land}</td>
                          <td className={cn(
                            "px-4 py-3 text-sm border border-zinc-200",
                            sale.hflIssue === 'Issue' && "bg-red-500 text-white font-bold"
                          )}>{sale.hflIssue}</td>
                          <td className={cn(
                            "px-4 py-3 text-sm border border-zinc-200",
                            sale.workPending === 'Issue' && "bg-red-500 text-white font-bold"
                          )}>{sale.workPending}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.remarks}</td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </Card>

              <div className="mt-8">
                <Last7DaysSalesOverview sales={sales} employees={employees} />
              </div>
            </div>
          </div>
        )}
        {activeTab === 'agentOverview' && hasPermission('agentOverview') && (
          <div className="space-y-8">
            <div className="flex items-center gap-2">
              <UserIcon className="w-6 h-6 text-orange-600" />
              <h2 className="text-3xl font-bold tracking-tight">Agent Overview</h2>
            </div>
            
            <div className="bg-white p-6 rounded-xl border border-zinc-200 shadow-sm space-y-4">
              <div className="flex flex-wrap items-center gap-4">
                <div className="flex-1 min-w-[200px]">
                  <label className="text-sm font-bold text-zinc-700 block mb-2">Select Agent</label>
                  <SearchableSelect
                    options={[
                      { value: '', label: '-- Select an Agent --' },
                      ...employees
                        .filter(e => e.status === 'Active')
                        .filter(e => 
                          getPermissionLevel('agentOverview') === 'Complete' || 
                          (e.email && e.email.toLowerCase() === profile?.email?.toLowerCase()) ||
                          (isBDEUser && e.bde === currentBDEName)
                        )
                        .sort((a, b) => a.name.localeCompare(b.name))
                        .map(emp => ({ value: emp.id!, label: `${emp.name} (${emp.employeeCode})`, key: emp.id }))
                    ]}
                    value={selectedAgentOverview}
                    onChange={setSelectedAgentOverview}
                    placeholder="-- Select an Agent --"
                  />
                </div>

                <div className="flex flex-wrap items-center gap-4">
                  <div className="flex bg-zinc-100 p-1 rounded-lg">
                    {(['Daily', 'Weekly', '4-Weekly', '8-Weekly', 'Monthly'] as const).map((tf) => (
                      <button
                        key={tf}
                        onClick={() => setAgentOverviewTimeframe(tf)}
                        className={cn(
                          "px-4 py-1.5 text-xs font-bold rounded-md transition-all",
                          agentOverviewTimeframe === tf 
                            ? "bg-white text-orange-600 shadow-sm" 
                            : "text-zinc-500 hover:text-zinc-700"
                        )}
                      >
                        {tf}
                      </button>
                    ))}
                  </div>

                  {agentOverviewTimeframe === 'Monthly' && (
                    <select 
                      className="px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm font-bold"
                      value={selectedMonthForAgentOverview}
                      onChange={(e) => setSelectedMonthForAgentOverview(e.target.value)}
                    >
                      <option value="">-- Select Month --</option>
                      {Array.from(new Set(allSortedWeeks.map(w => w.month))).map(month => (
                        <option key={month} value={month}>{month}</option>
                      ))}
                    </select>
                  )}
                </div>
              </div>
            </div>

            {selectedAgentOverview && employees.find(e => e.id === selectedAgentOverview) && (() => {
              const agent = employees.find(e => e.id === selectedAgentOverview)!;
              const agentSales = sales.filter(s => 
                s.agent?.toLowerCase().trim() === agent.name.toLowerCase().trim() ||
                (agent.email && s.agentEmail?.toLowerCase() === agent.email.toLowerCase()) ||
                s.associateBde?.toLowerCase().trim() === agent.name.toLowerCase().trim() ||
                (agent.email && s.associateBdeEmail?.toLowerCase() === agent.email.toLowerCase())
              );
              
              const currentWeek = allSortedWeeks[allSortedWeeks.length - 1];
              const last4WeeksForSummary = allSortedWeeks.slice(-4);
              const last8WeeksForSummary = allSortedWeeks.slice(-8);
              const currentMonth = agentOverviewTimeframe === 'Monthly' && selectedMonthForAgentOverview 
                ? selectedMonthForAgentOverview 
                : (currentWeek?.month || 'Current Month');
              const monthWeeks = allSortedWeeks.filter(w => w.month === currentMonth);
              const today = new Date().toISOString().split('T')[0];

              let filteredSales = [];
              if (agentOverviewTimeframe === 'Daily') {
                filteredSales = agentSales.filter(s => s.date === today);
              } else if (agentOverviewTimeframe === 'Weekly') {
                filteredSales = agentSales.filter(s => s.week === currentWeek?.weekName);
              } else if (agentOverviewTimeframe === '4-Weekly') {
                const last4WeekNames = last4WeeksForSummary.map(w => w.weekName);
                filteredSales = agentSales.filter(s => last4WeekNames.includes(s.week));
              } else if (agentOverviewTimeframe === '8-Weekly') {
                const last8WeekNames = last8WeeksForSummary.map(w => w.weekName);
                filteredSales = agentSales.filter(s => last8WeekNames.includes(s.week));
              } else if (agentOverviewTimeframe === 'Monthly') {
                const monthWeekNames = monthWeeks.map(w => w.weekName);
                filteredSales = agentSales.filter(s => monthWeekNames.includes(s.week));
              }

              const totalSalesValue = filteredSales.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
              const numSales = filteredSales.length;
              const totalMarginValue = filteredSales.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'totalMargin'), 0);
              const avgPackageValue = numSales > 0 ? totalSalesValue / numSales : 0;

              const advanceSales = filteredSales.filter(s => s.advanceCN === 'Advance');
              const advanceCount = advanceSales.length;
              const advanceValue = advanceSales.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
              const advancePerc = totalSalesValue > 0 ? (advanceValue / totalSalesValue) * 100 : 0;

              const cnSales = filteredSales.filter(s => s.advanceCN === 'Credit Note');
              const cnCount = cnSales.length;
              const cnValue = cnSales.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
              const cnPerc = totalSalesValue > 0 ? (cnValue / totalSalesValue) * 100 : 0;

              const cancelSales = filteredSales.filter(s => s.advanceCN === 'Cancel');
              const cancelCount = cancelSales.length;
              const cancelValue = cancelSales.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
              const cancelPerc = totalSalesValue > 0 ? (cancelValue / totalSalesValue) * 100 : 0;

              const workPendingCount = filteredSales.filter(s => s.workPending === 'Issue').length;

              // Helper to get net sales (excluding cancellations, credit notes, and low margin)
              const getNetSales = (salesList: Sales[]) => {
                return salesList.filter(s => {
                  const isCancelled = s.advanceCN === 'Cancel';
                  const isCreditNote = s.advanceCN === 'Credit Note';
                  // Removed low margin check as it was causing incorrect Bravo Point Achievement measurements
                  return !isCancelled && !isCreditNote;
                });
              };

              // Helper to get sales in the last N weeks from master
              const reversedWeeks = [...allSortedWeeks].reverse();
              const getSalesByWeeks = (numWeeks: number) => {
                const targetWeeks = reversedWeeks.slice(0, numWeeks).map(w => w.weekName);
                const rawFilteredSales = agentSales.filter(s => targetWeeks.includes(s.week || ''));
                const netFilteredSales = getNetSales(rawFilteredSales);
                return {
                  amount: netFilteredSales.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0),
                  count: netFilteredSales.length
                };
              };

              // Helper to get final target for a specific week
              const getFinalTargetForWeek = (weekName: string) => {
                const wkIndex = allSortedWeeks.findIndex(w => w.weekName === weekName);
                const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
                const weeklyBaseTarget = agent.target / 4;
                
                let percentageVal = 0;
                if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                  const tenureWeek = wkIndex - joiningWkIndex + 1;
                  if (tenureWeek <= 6) {
                    percentageVal = (agent as any)[`week${tenureWeek}Target`] || 0;
                  } else {
                    percentageVal = 100;
                  }
                }
                return (weeklyBaseTarget * percentageVal) / 100;
              };

              // Helper to get cumulative final target for last N weeks
              const getCumulativeTarget = (numWeeks: number) => {
                const targetWeeks = reversedWeeks.slice(0, numWeeks);
                return targetWeeks.reduce((sum, wk) => sum + getFinalTargetForWeek(wk.weekName), 0);
              };

              // Weekly (last 1 week)
              const weeklySalesData = getSalesByWeeks(1);
              const weeklyTarget = getCumulativeTarget(1);

              // 4 Weeks (last 4 weeks)
              const fourWeekSalesData = getSalesByWeeks(4);
              const fourWeekTarget = getCumulativeTarget(4);

              // 8 Weeks (last 8 weeks)
              const eightWeekSalesData = getSalesByWeeks(8);
              const eightWeekTarget = getCumulativeTarget(8);

              // 12 Weeks (last 12 weeks)
              const twelveWeekSalesData = getSalesByWeeks(12);
              const twelveWeekTarget = getCumulativeTarget(12);

              // 24 Weeks (last 24 weeks)
              const twentyFourWeekSalesData = getSalesByWeeks(24);
              const twentyFourWeekTarget = getCumulativeTarget(24);

              // Calculate target percentage for current week (logic from Weekly Overview)
              const joiningWeekIndex = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
              const currentWeekIndex = allSortedWeeks.length - 1;
              let targetPercentageThisWeek = '-';
              if (joiningWeekIndex !== -1 && currentWeekIndex >= joiningWeekIndex) {
                const tenureWeek = currentWeekIndex - joiningWeekIndex + 1;
                if (tenureWeek <= 6) {
                  targetPercentageThisWeek = `${(agent as any)[`week${tenureWeek}Target`]}%`;
                } else {
                  targetPercentageThisWeek = '100%';
                }
              }

              // Calculate last 12 weeks targets for the table
              const last12Weeks = reversedWeeks.slice(0, 12).reverse();
              // Calculate last 8 weeks targets for the highlight card
              const last8Weeks = reversedWeeks.slice(0, 8).reverse();
              const weeklyBaseTarget = agent.target / 4;

              // Calculate totals for the 12-week table
              const tableTotals = last12Weeks.reduce((acc, wk) => {
                const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
                const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
                const isJoined = joiningWkIndex !== -1 && wkIndex >= joiningWkIndex;
                
                let percentageVal = 0;
                if (isJoined) {
                  const tenureWeek = wkIndex - joiningWkIndex + 1;
                  if (tenureWeek <= 6) {
                    percentageVal = (agent as any)[`week${tenureWeek}Target`] || 0;
                  } else {
                    percentageVal = 100;
                  }
                }
                
                const currentWeeklyBaseTarget = isJoined ? weeklyBaseTarget : 0;
                const finalTarget = (weeklyBaseTarget * percentageVal) / 100;
                const weekSalesRaw = agentSales.filter(s => s.week === wk.weekName);
                const weekSales = getNetSales(weekSalesRaw).reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
                const shortfall = Math.max(0, finalTarget - weekSales);
                
                const workstationCost = isJoined ? (agent.salary * 3) / 4 : 0;
                const contribution = isJoined ? (weekSales * 0.12) - workstationCost : 0;

                return {
                  weeklyTarget: acc.weeklyTarget + currentWeeklyBaseTarget,
                  finalTarget: acc.finalTarget + finalTarget,
                  sales: acc.sales + weekSales,
                  shortfall: acc.shortfall + shortfall,
                  workstationCost: acc.workstationCost + workstationCost,
                  contribution: acc.contribution + contribution
                };
              }, { weeklyTarget: 0, finalTarget: 0, sales: 0, shortfall: 0, workstationCost: 0, contribution: 0 });

              // Training completion
              // Calculate based on active training materials and agent's training points
              const activeTrainingMaterials = trainingMaterials.filter(tm => tm.isActive);
              const totalTrainingPoints = activeTrainingMaterials.reduce((sum, tm) => sum + (tm.points || 0), 0);
              const totalTasksCount = activeTrainingMaterials.length;
              
              const pointsMap = agent.trainingPoints || {};
              const earnedTrainingPoints = activeTrainingMaterials.reduce((sum, tm) => sum + (pointsMap[tm.id!] || 0), 0);
              const completedTasksCount = activeTrainingMaterials.filter(tm => (pointsMap[tm.id!] || 0) >= tm.points).length;
                
              const trainingCompletion = totalTrainingPoints > 0 
                ? Math.round((earnedTrainingPoints / totalTrainingPoints) * 100) 
                : 0;

              // Calculate totals for the 8-week highlight
              const eightWeekTotals = last8Weeks.reduce((acc, wk) => {
                const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
                const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
                const isJoined = joiningWkIndex !== -1 && wkIndex >= joiningWkIndex;
                
                let percentageVal = 0;
                if (isJoined) {
                  const tenureWeek = wkIndex - joiningWkIndex + 1;
                  if (tenureWeek <= 6) {
                    percentageVal = (agent as any)[`week${tenureWeek}Target`] || 0;
                  } else {
                    percentageVal = 100;
                  }
                }
                
                const weekSalesRaw = agentSales.filter(s => s.week === wk.weekName);
                const weekSales = getNetSales(weekSalesRaw).reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
                
                const workstationCost = isJoined ? (agent.salary * 3) / 4 : 0;
                const contribution = isJoined ? (weekSales * 0.12) - workstationCost : 0;

                return {
                  contribution: acc.contribution + contribution
                };
              }, { contribution: 0 });

              const eightWeekAchievement = eightWeekTarget > 0 ? Math.round((eightWeekSalesData.amount / eightWeekTarget) * 100) : 0;
              const fourWeekAchievement = fourWeekTarget > 0 ? Math.round((fourWeekSalesData.amount / fourWeekTarget) * 100) : 0;
              
              const joiningWeekIndexForTenure = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
              const tenureWeeks = joiningWeekIndexForTenure !== -1 ? currentWeekIndex - joiningWeekIndexForTenure + 1 : 0;

              const getStatusStyles = (val: number) => {
                if (val < 70) return { color: 'text-red-600', bg: 'bg-red-50', border: 'border-red-200', label: 'Bad' };
                if (val <= 90) return { color: 'text-orange-600', bg: 'bg-orange-50', border: 'border-orange-200', label: 'Average' };
                if (val <= 120) return { color: 'text-green-600', bg: 'bg-green-50', border: 'border-green-200', label: 'Good' };
                return { color: 'text-blue-600', bg: 'bg-blue-50', border: 'border-blue-200', label: 'Exceptional' };
              };

              const status = getStatusStyles(eightWeekAchievement);

              // Incentive Calculations
              const isWeeklyEligible = eightWeekAchievement > 70;
              const currentWeeklyIncentive = isWeeklyEligible ? (weeklySalesData.amount * 0.001) : 0;

              const isMonthlyEligible = fourWeekSalesData.amount >= fourWeekTarget;
              const currentMonthlyIncentive = isMonthlyEligible 
                ? ((fourWeekSalesData.amount - fourWeekTarget) * 0.03 + 2000) 
                : 0;

              // 3 Month Consecutive Achievement Incentive Logic
              const last12WeeksList = reversedWeeks.slice(0, 12);
              const monthsMap: { [key: string]: { target: number, sales: number, joinedWeeks: number } } = {};
              
              last12WeeksList.forEach(wk => {
                const month = wk.month;
                if (!monthsMap[month]) {
                  monthsMap[month] = { target: 0, sales: 0, joinedWeeks: 0 };
                }
                const target = getFinalTargetForWeek(wk.weekName);
                monthsMap[month].target += target;
                const weekSalesRaw = agentSales.filter(s => s.week === wk.weekName);
                monthsMap[month].sales += getNetSales(weekSalesRaw).reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
                
                const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
                const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
                if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                  monthsMap[month].joinedWeeks++;
                }
              });

              const uniqueMonths = Array.from(new Set(last12WeeksList.map(wk => wk.month)));
              
              let isThreeMonthConsecutiveEligible = uniqueMonths.length >= 3;
              if (isThreeMonthConsecutiveEligible) {
                // Check the last 3 months
                for (let i = 0; i < 3; i++) {
                  const m = uniqueMonths[i];
                  // Must have joined in that month and met target
                  if (monthsMap[m].joinedWeeks === 0 || monthsMap[m].sales < monthsMap[m].target) {
                    isThreeMonthConsecutiveEligible = false;
                    break;
                  }
                }
              }
              const threeMonthIncentive = isThreeMonthConsecutiveEligible ? 15000 : 0;

              return (
                <div className="space-y-6">
                  {/* Summary Boxes */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-3">
                    <Card 
                      className="p-3 bg-orange-600 text-white border-none shadow-lg shadow-orange-100 cursor-pointer hover:scale-[1.02] transition-transform"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Total Sales Details', sales: filteredSales })}
                    >
                      <p className="text-orange-100 text-[10px] font-bold uppercase tracking-wider">Sales Value</p>
                      <p className="text-lg font-black mt-1">₹{totalSalesValue.toLocaleString()}</p>
                      <p className="text-[9px] text-orange-200 mt-1">{numSales} Trans.</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Margin Details', sales: filteredSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Margin</p>
                      <p className={cn("text-lg font-black mt-1", totalMarginValue >= 0 ? "text-emerald-600" : "text-rose-600")}>
                        ₹{totalMarginValue.toLocaleString()}
                      </p>
                      <p className="text-[9px] text-zinc-500 mt-1">{totalSalesValue > 0 ? ((totalMarginValue / totalSalesValue) * 100).toFixed(1) : 0}%</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Avg Package Details', sales: filteredSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Avg Pkg</p>
                      <p className="text-lg font-black mt-1 text-blue-600">₹{Math.round(avgPackageValue).toLocaleString()}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">{numSales} Sales</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Advance Details', sales: advanceSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Advance</p>
                      <p className="text-lg font-black mt-1 text-orange-600">{advanceCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">₹{advanceValue.toLocaleString()} ({advancePerc.toFixed(1)}%)</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Credit Note Details', sales: cnSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">CN</p>
                      <p className="text-lg font-black mt-1 text-purple-600">{cnCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">₹{cnValue.toLocaleString()} ({cnPerc.toFixed(1)}%)</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Cancellation Details', sales: cancelSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Cancel</p>
                      <p className="text-lg font-black mt-1 text-rose-600">{cancelCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">₹{cancelValue.toLocaleString()} ({cancelPerc.toFixed(1)}%)</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Work Pending Details', sales: filteredSales.filter(s => s.workPending === 'Issue') })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Pending</p>
                      <p className="text-lg font-black mt-1 text-amber-600">{workPendingCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">Issues</p>
                    </Card>
                  </div>

                  {/* 8 Weeks Performance Highlight */}
                  <Card className={cn("p-8 border-2 shadow-lg transition-all hover:scale-[1.01]", status.bg, status.border)}>
                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-8 items-center">
                      <div className="space-y-2 text-center lg:text-left lg:border-r lg:border-zinc-200 pr-0 lg:pr-8">
                        <p className={cn("text-sm font-bold uppercase tracking-widest", status.color)}>8 Weeks Bravo Point Achievement</p>
                        <div className="flex items-center justify-center lg:justify-start gap-4">
                          <h1 
                            className={cn("text-7xl font-black tracking-tighter cursor-pointer hover:opacity-80 transition-opacity", status.color)}
                            onClick={() => handleOpenBravoCalculation(agent, 8)}
                            title="Click to see 8-week bravo calculation"
                          >
                            {eightWeekAchievement}%
                          </h1>
                          <div className={cn("px-3 py-1 rounded-full text-[10px] font-bold uppercase border-2", status.color, status.border, "bg-white/50")}>
                            {status.label}
                          </div>
                        </div>
                        <p className="text-zinc-500 text-[10px] font-medium">Based on performance over the last 8 weeks</p>
                      </div>

                      <div className="space-y-2 text-center lg:text-left lg:border-r lg:border-zinc-200 pr-0 lg:pr-8">
                        <p className="text-sm font-bold uppercase tracking-widest text-zinc-500">Total Contribution (8 Weeks)</p>
                        <h2 
                          className={cn("text-5xl font-black tracking-tighter cursor-pointer hover:opacity-80 transition-opacity", eightWeekTotals.contribution >= 0 ? "text-green-600" : "text-red-600")}
                          onClick={() => handleOpenContributionCalculation(agent, 8)}
                          title="Click to see contribution calculation"
                        >
                          ₹{eightWeekTotals.contribution.toLocaleString()}
                        </h2>
                        <p className="text-zinc-500 text-[10px] font-medium">Net contribution to company after workstation costs</p>
                        
                        {/* Approx Eligible Bonus Card */}
                        {(() => {
                          const agentIncentives = incentives.filter(i => i.employeeId === selectedAgentOverview);
                          const agentPayments = incentivePayments.filter(p => p.employeeId === selectedAgentOverview);
                          const totalDue = agentIncentives.reduce((sum, i) => sum + i.amount, 0);
                          const totalPaid = agentPayments.reduce((sum, p) => sum + p.amount, 0);
                          const approxBalance = totalDue - totalPaid;
                          
                          return (
                            <div 
                              className="pt-4 mt-4 border-t-2 border-zinc-200 cursor-pointer hover:bg-zinc-50 p-2 rounded-lg transition-colors group"
                              onClick={() => setAgentLedgerModal({ isOpen: true, employeeId: selectedAgentOverview })}
                            >
                              <p className="text-xs font-bold uppercase tracking-widest text-zinc-500 group-hover:text-blue-600 transition-colors">Approx Eligible Bonus to Agent</p>
                              <h2 className={cn("text-5xl font-black tracking-tighter", approxBalance >= 0 ? "text-blue-600" : "text-red-600")}>
                                ₹{approxBalance.toLocaleString()}
                              </h2>
                              <p className="text-zinc-500 text-[9px] font-medium">Current outstanding (incentive + Quaterly) balance</p>
                            </div>
                          );
                        })()}
                      </div>

                      <div className="space-y-4 text-center lg:text-left">
                        <div>
                          <p className="text-xs font-bold uppercase tracking-widest text-zinc-500 mb-1">Current Floor Timing</p>
                          <div className="flex items-center justify-center lg:justify-start gap-2">
                            <Clock className="w-4 h-4 text-zinc-400" />
                            <span className="text-2xl font-black text-zinc-800">
                              {tenureWeeks > 4 ? (
                                eightWeekAchievement > 90 ? "8:30 Minutes" : eightWeekAchievement >= 70 ? "9:30 Minutes" : "10:30 Minutes"
                              ) : "9:30 AM - 8:00 PM"}
                            </span>
                          </div>
                        </div>
                        <div>
                          <p className="text-xs font-bold uppercase tracking-widest text-zinc-500 mb-1">Sunday Requirement</p>
                          <div className="flex items-center justify-center lg:justify-start gap-2">
                            <CalendarIcon className="w-4 h-4 text-zinc-400" />
                            <span className={cn("text-sm font-bold", tenureWeeks > 6 && fourWeekAchievement < 70 && eightWeekAchievement < 70 ? "text-red-600" : "text-green-600")}>
                              {tenureWeeks > 6 && fourWeekAchievement < 70 && eightWeekAchievement < 70 ? "2nd/4th Sunday Working Required" : "Standard Offs (Sundays Off)"}
                            </span>
                          </div>
                        </div>
                        <div>
                          <p className="text-xs font-bold uppercase tracking-widest text-zinc-500 mb-1">HR Review Status</p>
                          <div className="flex items-center justify-center lg:justify-start gap-2">
                            <Shield className="w-4 h-4 text-zinc-400" />
                            <span className={cn("text-sm font-bold", tenureWeeks > 4 && eightWeekAchievement < 30 ? "text-red-600" : "text-green-600")}>
                              {tenureWeeks > 4 && eightWeekAchievement < 30 ? (
                                <span className="flex items-center gap-1 animate-pulse">
                                  <AlertCircle className="w-3 h-3" /> Talk to HR
                                </span>
                              ) : "Clear"}
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </Card>

                  {/* Daily Sales Analysis Chart */}
                  {(() => {
                    const dailySalesDataRaw = filteredSales.reduce((acc: any, sale) => {
                      const date = sale.date || 'Unknown';
                      if (!acc[date]) acc[date] = { date, count: 0, value: 0 };
                      acc[date].count += 1;
                      acc[date].value += (sale.packageValue || 0);
                      return acc;
                    }, {});

                    const dailySalesData = Object.values(dailySalesDataRaw)
                      .sort((a: any, b: any) => new Date(a.date as string).getTime() - new Date(b.date as string).getTime());

                    return (
                      <Card className="p-6">
                        <h3 className="text-sm font-bold text-zinc-900 uppercase mb-6 flex items-center gap-2">
                          <CalendarIcon className="w-4 h-4 text-orange-600" />
                          Daily Sales Analysis (Value & Count)
                        </h3>
                        <div className="h-[500px] w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dailySalesData} margin={{ top: 130, right: 30, left: 20, bottom: 60 }}>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                              <XAxis 
                                dataKey="date" 
                                fontSize={10} 
                                fontWeight="bold"
                                angle={-45}
                                textAnchor="end"
                                interval={0}
                                tickFormatter={(val) => val === 'Unknown' ? val : new Date(val).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })}
                              />
                              <YAxis fontSize={10} tickFormatter={(val) => `₹${(val/100000).toFixed(2)}L`} />
                              <Tooltip 
                                contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                                formatter={(value: number, name: string) => [name === 'value' ? `₹${value.toLocaleString()}` : value, name === 'value' ? 'Amount' : 'Sales Count']}
                              />
                              <Bar dataKey="value" name="Amount" fill="#ea580c" radius={[4, 4, 0, 0]}>
                                <LabelList 
                                  dataKey="value" 
                                  position="top" 
                                  content={(props: any) => {
                                    const { x, y, width, value, index } = props;
                                    const dataPoint = dailySalesData[index] as any;
                                    const count = dataPoint?.count || 0;
                                    return (
                                      <text 
                                        x={x + width / 2} 
                                        y={y - 45} 
                                        fill="#ea580c" 
                                        textAnchor="middle" 
                                        fontSize={14} 
                                        fontWeight="black"
                                        transform={`rotate(-90, ${x + width / 2}, ${y - 45})`}
                                      >
                                        ₹{(value / 100000).toFixed(2)}L ({count})
                                      </text>
                                    );
                                  }}
                                />
                              </Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </Card>
                    );
                  })()}

                  {/* Personal Details */}
                  <Card className="p-6">
                    <h3 className="text-lg font-bold mb-4 border-b pb-2">Personal Details</h3>
                    <div className="grid grid-cols-1 md:grid-cols-3 lg:grid-cols-4 gap-6">
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Name</p>
                        <p className="text-lg font-medium">{agent.name}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Employee Code</p>
                        <p className="text-lg font-medium">{agent.employeeCode}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Email</p>
                        <p className="text-lg font-medium">{agent.email}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Position</p>
                        <p className="text-lg font-medium">{agent.position}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Joining Date</p>
                        <p className="text-lg font-medium">{agent.joiningDate}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Joining Week</p>
                        <p className="text-lg font-medium">{agent.joiningWeek}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">BDE</p>
                        <p className="text-lg font-medium">{agent.bde}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Status</p>
                        <p className="text-lg font-medium">
                          <Badge className={agent.status === 'Active' ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}>
                            {agent.status}
                          </Badge>
                        </p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Weekly target</p>
                        <p className="text-lg font-medium">₹{(agent.target / 4).toLocaleString()}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Target % (This Week)</p>
                        <p className="text-lg font-medium">{targetPercentageThisWeek}</p>
                      </div>
                      <div>
                        <p className="text-xs text-zinc-500 uppercase font-bold">Assoc BDE</p>
                        <p className="text-lg font-medium">{agent.isAssocBDE ? 'Yes' : 'No'}</p>
                      </div>
                    </div>
                  </Card>

                  {/* Performance Overview */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
                    {/* Weekly */}
                    <Card className="p-6 bg-gradient-to-br from-orange-50 to-white border-orange-100">
                      <h4 className="text-sm font-bold text-orange-800 uppercase mb-4">Weekly Performance</h4>
                      <div className="space-y-2">
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Sales Value</p>
                          <p className="text-xl font-bold text-orange-600">₹{weeklySalesData.amount.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">No. of Sales</p>
                          <p className="text-sm font-medium text-zinc-700">{weeklySalesData.count}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Target</p>
                          <p className="text-sm font-medium text-zinc-700">₹{weeklyTarget.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Shortfall for target</p>
                          <p className="text-sm font-bold text-red-600">₹{Math.max(0, weeklyTarget - weeklySalesData.amount).toLocaleString()}</p>
                        </div>
                        <div className="pt-2 border-t border-orange-100">
                          <div className="flex justify-between text-xs mb-1">
                            <span className="font-medium text-orange-800">Achievement (Bravo Points)</span>
                            <span 
                              className="text-lg font-bold text-orange-600 cursor-pointer hover:underline"
                              onClick={() => handleOpenBravoCalculation(agent, 1)}
                              title="Click to see weekly bravo calculation"
                            >
                              {weeklyTarget > 0 ? Math.round((weeklySalesData.amount / weeklyTarget) * 100) : 0}%
                            </span>
                          </div>
                          <div className="w-full bg-orange-100 rounded-full h-2">
                            <div className="bg-orange-500 h-2 rounded-full" style={{ width: `${Math.min(100, weeklyTarget > 0 ? (weeklySalesData.amount / weeklyTarget) * 100 : 0)}%` }}></div>
                          </div>
                        </div>
                      </div>
                    </Card>

                    {/* 4 Weeks */}
                    <Card className="p-6 bg-gradient-to-br from-blue-50 to-white border-blue-100">
                      <h4 className="text-sm font-bold text-blue-800 uppercase mb-4">4-Week Performance</h4>
                      <div className="space-y-2">
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Sales Value</p>
                          <p className="text-xl font-bold text-blue-600">₹{fourWeekSalesData.amount.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">No. of Sales</p>
                          <p className="text-sm font-medium text-zinc-700">{fourWeekSalesData.count}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Target</p>
                          <p className="text-sm font-medium text-zinc-700">₹{fourWeekTarget.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Shortfall for target</p>
                          <p className="text-sm font-bold text-red-600">₹{Math.max(0, fourWeekTarget - fourWeekSalesData.amount).toLocaleString()}</p>
                        </div>
                        <div className="pt-2 border-t border-blue-100">
                          <div className="flex justify-between text-xs mb-1">
                            <span className="font-medium text-blue-800">Achievement (Bravo Points)</span>
                            <span 
                              className="text-lg font-bold text-blue-600 cursor-pointer hover:underline"
                              onClick={() => handleOpenBravoCalculation(agent, 4)}
                              title="Click to see 4-week bravo calculation"
                            >
                              {fourWeekTarget > 0 ? Math.round((fourWeekSalesData.amount / fourWeekTarget) * 100) : 0}%
                            </span>
                          </div>
                          <div className="w-full bg-blue-100 rounded-full h-2">
                            <div className="bg-blue-500 h-2 rounded-full" style={{ width: `${Math.min(100, fourWeekTarget > 0 ? (fourWeekSalesData.amount / fourWeekTarget) * 100 : 0)}%` }}></div>
                          </div>
                        </div>
                      </div>
                    </Card>

                    {/* 8 Weeks */}
                    <Card className="p-6 bg-gradient-to-br from-purple-50 to-white border-purple-100">
                      <h4 className="text-sm font-bold text-purple-800 uppercase mb-4">8-Week Performance</h4>
                      <div className="space-y-2">
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Sales Value</p>
                          <p className="text-xl font-bold text-purple-600">₹{eightWeekSalesData.amount.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">No. of Sales</p>
                          <p className="text-sm font-medium text-zinc-700">{eightWeekSalesData.count}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Target</p>
                          <p className="text-sm font-medium text-zinc-700">₹{eightWeekTarget.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Shortfall for target</p>
                          <p className="text-sm font-bold text-red-600">₹{Math.max(0, eightWeekTarget - eightWeekSalesData.amount).toLocaleString()}</p>
                        </div>
                        <div className="pt-2 border-t border-purple-100">
                          <div className="flex justify-between text-xs mb-1">
                            <span className="font-medium text-purple-800">Achievement (Bravo Points)</span>
                            <span 
                              className="text-lg font-bold text-purple-600 cursor-pointer hover:underline"
                              onClick={() => handleOpenBravoCalculation(agent, 8)}
                              title="Click to see 8-week bravo calculation"
                            >
                              {eightWeekTarget > 0 ? Math.round((eightWeekSalesData.amount / eightWeekTarget) * 100) : 0}%
                            </span>
                          </div>
                          <div className="w-full bg-purple-100 rounded-full h-2">
                            <div className="bg-purple-500 h-2 rounded-full" style={{ width: `${Math.min(100, eightWeekTarget > 0 ? (eightWeekSalesData.amount / eightWeekTarget) * 100 : 0)}%` }}></div>
                          </div>
                        </div>
                      </div>
                    </Card>

                    {/* 12 Weeks */}
                    <Card className="p-6 bg-gradient-to-br from-indigo-50 to-white border-indigo-100">
                      <h4 className="text-sm font-bold text-indigo-800 uppercase mb-4">12-Week Performance</h4>
                      <div className="space-y-2">
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Sales Value</p>
                          <p className="text-xl font-bold text-indigo-600">₹{twelveWeekSalesData.amount.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">No. of Sales</p>
                          <p className="text-sm font-medium text-zinc-700">{twelveWeekSalesData.count}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Target</p>
                          <p className="text-sm font-medium text-zinc-700">₹{twelveWeekTarget.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Shortfall for target</p>
                          <p className="text-sm font-bold text-red-600">₹{Math.max(0, twelveWeekTarget - twelveWeekSalesData.amount).toLocaleString()}</p>
                        </div>
                        <div className="pt-2 border-t border-indigo-100">
                          <div className="flex justify-between text-xs mb-1">
                            <span className="font-medium text-indigo-800">Achievement (Bravo Points)</span>
                            <span 
                              className="text-lg font-bold text-indigo-600 cursor-pointer hover:underline"
                              onClick={() => handleOpenBravoCalculation(agent, 12)}
                              title="Click to see 12-week bravo calculation"
                            >
                              {twelveWeekTarget > 0 ? Math.round((twelveWeekSalesData.amount / twelveWeekTarget) * 100) : 0}%
                            </span>
                          </div>
                          <div className="w-full bg-indigo-100 rounded-full h-2">
                            <div className="bg-indigo-500 h-2 rounded-full" style={{ width: `${Math.min(100, twelveWeekTarget > 0 ? (twelveWeekSalesData.amount / twelveWeekTarget) * 100 : 0)}%` }}></div>
                          </div>
                        </div>
                      </div>
                    </Card>

                    {/* 24 Weeks */}
                    <Card className="p-6 bg-gradient-to-br from-emerald-50 to-white border-emerald-100">
                      <h4 className="text-sm font-bold text-emerald-800 uppercase mb-4">24-Week Performance</h4>
                      <div className="space-y-2">
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Sales Value</p>
                          <p className="text-xl font-bold text-emerald-600">₹{twentyFourWeekSalesData.amount.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">No. of Sales</p>
                          <p className="text-sm font-medium text-zinc-700">{twentyFourWeekSalesData.count}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Target</p>
                          <p className="text-sm font-medium text-zinc-700">₹{twentyFourWeekTarget.toLocaleString()}</p>
                        </div>
                        <div className="flex justify-between items-end">
                          <p className="text-xs text-zinc-500">Shortfall for target</p>
                          <p className="text-sm font-bold text-red-600">₹{Math.max(0, twentyFourWeekTarget - twentyFourWeekSalesData.amount).toLocaleString()}</p>
                        </div>
                        <div className="pt-2 border-t border-emerald-100">
                          <div className="flex justify-between text-xs mb-1">
                            <span className="font-medium text-emerald-800">Achievement (Bravo Points)</span>
                            <span 
                              className="text-lg font-bold text-emerald-600 cursor-pointer hover:underline"
                              onClick={() => handleOpenBravoCalculation(agent, 24)}
                              title="Click to see 24-week bravo calculation"
                            >
                              {twentyFourWeekTarget > 0 ? Math.round((twentyFourWeekSalesData.amount / twentyFourWeekTarget) * 100) : 0}%
                            </span>
                          </div>
                          <div className="w-full bg-emerald-100 rounded-full h-2">
                            <div className="bg-emerald-500 h-2 rounded-full" style={{ width: `${Math.min(100, twentyFourWeekTarget > 0 ? (twentyFourWeekSalesData.amount / twentyFourWeekTarget) * 100 : 0)}%` }}></div>
                          </div>
                        </div>
                      </div>
                    </Card>
                  </div>

                  {/* Weekly Target Table */}
                  <Card className="p-6">
                    <h3 className="text-lg font-bold mb-4 border-b pb-2">Weekly Target Details (Last 12 Weeks)</h3>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse border border-zinc-200">
                        <thead>
                          <tr className="bg-zinc-50 text-xs font-bold text-zinc-500 uppercase">
                            <th className="px-4 py-3 border border-zinc-200">Week Name</th>
                            <th className="px-4 py-3 border border-zinc-200">Month</th>
                            <th className="px-4 py-3 border border-zinc-200">Weekly target</th>
                            <th className="px-4 py-3 border border-zinc-200">Percentage</th>
                            <th className="px-4 py-3 border border-zinc-200">Final target</th>
                            <th className="px-4 py-3 border border-zinc-200">Sales</th>
                            <th className="px-4 py-3 border border-zinc-200">Shortfall for target</th>
                            <th className="px-4 py-3 border border-zinc-200">Workstation cost</th>
                            <th className="px-4 py-3 border border-zinc-200">Contribution to company</th>
                          </tr>
                        </thead>
                        <tbody>
                          {last12Weeks.map((wk) => {
                            const wkIndex = allSortedWeeks.findIndex(w => w.weekName === wk.weekName);
                            const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
                            const isJoined = joiningWkIndex !== -1 && wkIndex >= joiningWkIndex;
                            
                            let percentageVal = 0;
                            let percentageStr = '-';
                            
                            if (isJoined) {
                              const tenureWeek = wkIndex - joiningWkIndex + 1;
                              if (tenureWeek <= 6) {
                                percentageVal = (agent as any)[`week${tenureWeek}Target`] || 0;
                                percentageStr = `${percentageVal}%`;
                              } else {
                                percentageVal = 100;
                                percentageStr = '100%';
                              }
                            }
                            
                            const currentWeeklyBaseTarget = isJoined ? weeklyBaseTarget : 0;
                            const finalTarget = (weeklyBaseTarget * percentageVal) / 100;
                            const weekSalesRaw = agentSales.filter(s => s.week === wk.weekName);
                            const weekSales = getNetSales(weekSalesRaw).reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
                            const shortfall = Math.max(0, finalTarget - weekSales);
                            
                            const workstationCost = isJoined ? (agent.salary * 3) / 4 : 0;
                            const contribution = isJoined ? (weekSales * 0.12) - workstationCost : 0;

                            return (
                              <tr key={wk.id} className="hover:bg-zinc-50 transition-colors">
                                <td className="px-4 py-3 text-sm font-medium border border-zinc-200">{wk.weekName}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200">{wk.month}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200">₹{currentWeeklyBaseTarget.toLocaleString()}</td>
                                <td className="px-4 py-3 text-lg font-bold text-orange-600 border border-zinc-200 bg-orange-50 cursor-pointer hover:underline" onClick={() => handleOpenBravoCalculation(agent, 1)} title="Click to see weekly bravo calculation">{percentageStr}</td>
                                <td className="px-4 py-3 text-sm font-bold border border-zinc-200">₹{finalTarget.toLocaleString()}</td>
                                <td 
                                  className="px-4 py-3 text-sm border border-zinc-200 text-green-600 font-medium cursor-pointer hover:underline"
                                  onClick={() => setSummaryDetailModal({ isOpen: true, title: `Sales Details - ${wk.weekName}`, sales: weekSalesRaw })}
                                  title="Click to see sales details"
                                >
                                  ₹{weekSales.toLocaleString()}
                                </td>
                                <td className="px-4 py-3 text-sm border border-zinc-200 text-red-600 font-medium">₹{shortfall.toLocaleString()}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200">₹{workstationCost.toLocaleString()}</td>
                                <td 
                                  className="px-4 py-3 text-sm border border-zinc-200 font-medium cursor-pointer hover:underline" 
                                  style={{ color: contribution >= 0 ? '#16a34a' : '#dc2626' }}
                                  onClick={() => handleOpenContributionCalculation(agent, 1)}
                                  title="Click to see contribution calculation"
                                >
                                  ₹{contribution.toLocaleString()}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="bg-zinc-100 font-bold">
                            <td className="px-4 py-3 border border-zinc-200" colSpan={2}>Total</td>
                            <td className="px-4 py-3 border border-zinc-200">₹{tableTotals.weeklyTarget.toLocaleString()}</td>
                            <td className="px-4 py-3 border border-zinc-200">-</td>
                            <td className="px-4 py-3 border border-zinc-200">₹{tableTotals.finalTarget.toLocaleString()}</td>
                            <td className="px-4 py-3 border border-zinc-200 text-green-700">₹{tableTotals.sales.toLocaleString()}</td>
                            <td className="px-4 py-3 border border-zinc-200 text-red-700">₹{tableTotals.shortfall.toLocaleString()}</td>
                            <td className="px-4 py-3 border border-zinc-200">₹{tableTotals.workstationCost.toLocaleString()}</td>
                            <td className="px-4 py-3 border border-zinc-200" style={{ color: tableTotals.contribution >= 0 ? '#15803d' : '#b91c1c' }}>₹{tableTotals.contribution.toLocaleString()}</td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>
                  </Card>

                  {/* Incentive Table & Calculator */}
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
                    {/* Current Incentive Status */}
                    <Card className="p-6">
                      <h3 className="text-lg font-bold mb-4 border-b pb-2 flex items-center gap-2">
                        <IndianRupee className="w-5 h-5 text-green-600" />
                        Current Incentive Status
                      </h3>
                      <div className="space-y-4">
                        <div className={cn("p-4 rounded-lg border flex justify-between items-center", isWeeklyEligible ? "bg-green-50 border-green-200" : "bg-zinc-50 border-zinc-200")}>
                          <div>
                            <p className="font-bold text-sm">Weekly Incentive</p>
                            <p className="text-xs text-zinc-500">Eligibility: 8-week Bravo &gt; 70%</p>
                          </div>
                          <div className="text-right">
                            <p className={cn("text-lg font-bold", isWeeklyEligible ? "text-green-600" : "text-zinc-400")}>
                              ₹{currentWeeklyIncentive.toLocaleString()} approx
                            </p>
                            <Badge className={isWeeklyEligible ? "bg-green-100 text-green-700" : "bg-zinc-200 text-zinc-600"}>
                              {isWeeklyEligible ? "Eligible" : "Not Eligible"}
                            </Badge>
                          </div>
                        </div>
                        {isWeeklyEligible && (
                          <div className="mt-2 pt-2 border-t border-green-200 text-[10px] text-green-700 space-y-1">
                            <p className="font-bold uppercase tracking-wider">Weekly Breakdown:</p>
                            <div className="flex justify-between">
                              <span>Weekly Net Sales:</span>
                              <span>₹{weeklySalesData.amount.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between">
                              <span>Incentive Rate:</span>
                              <span>0.1%</span>
                            </div>
                            <div className="flex justify-between font-bold pt-1 border-t border-green-200">
                              <span>Total Weekly:</span>
                              <span>₹{currentWeeklyIncentive.toLocaleString()}</span>
                            </div>
                          </div>
                        )}

                        <div className={cn("p-4 rounded-lg border flex flex-col gap-2", isMonthlyEligible ? "bg-green-50 border-green-200" : "bg-zinc-50 border-zinc-200")}>
                          <div className="flex justify-between items-center">
                            <div>
                              <p className="font-bold text-sm">Monthly Incentive</p>
                              <p className="text-xs text-zinc-500">Eligibility: Monthly target achieved</p>
                            </div>
                            <div className="text-right">
                              <p className={cn("text-lg font-bold", isMonthlyEligible ? "text-green-600" : "text-zinc-400")}>
                                ₹{currentMonthlyIncentive.toLocaleString()} approx
                              </p>
                              <Badge className={isMonthlyEligible ? "bg-green-100 text-green-700" : "bg-zinc-200 text-zinc-600"}>
                                {isMonthlyEligible ? "Eligible" : "Not Eligible"}
                              </Badge>
                            </div>
                          </div>
                          {isMonthlyEligible && (
                            <div className="mt-2 pt-2 border-t border-green-200 text-[10px] text-green-700 space-y-1">
                              <p className="font-bold uppercase tracking-wider">Monthly Breakdown:</p>
                              <div className="flex justify-between">
                                <span>Sales above target (₹{(fourWeekSalesData.amount - fourWeekTarget).toLocaleString()}):</span>
                                <span>₹{((fourWeekSalesData.amount - fourWeekTarget) * 0.03).toLocaleString()} (3%)</span>
                              </div>
                              <div className="flex justify-between">
                                <span>Base Bonus:</span>
                                <span>₹2,000</span>
                              </div>
                              <div className="flex justify-between font-bold pt-1 border-t border-green-200">
                                <span>Total Monthly:</span>
                                <span>₹{currentMonthlyIncentive.toLocaleString()}</span>
                              </div>
                            </div>
                          )}
                        </div>

                        <div className={cn("p-4 rounded-lg border flex flex-col gap-2", isThreeMonthConsecutiveEligible ? "bg-green-50 border-green-200" : "bg-zinc-50 border-zinc-200")}>
                          <div className="flex justify-between items-center">
                            <div>
                              <p className="font-bold text-sm">3 Month Consecutive Achievement Incentive</p>
                              <p className="text-xs text-zinc-500">Eligibility: Target completed consecutively in all last 3 months</p>
                            </div>
                            <div className="text-right">
                              <p className={cn("text-lg font-bold", isThreeMonthConsecutiveEligible ? "text-green-600" : "text-zinc-400")}>
                                ₹{threeMonthIncentive.toLocaleString()} approx
                              </p>
                              <Badge className={isThreeMonthConsecutiveEligible ? "bg-green-100 text-green-700" : "bg-zinc-200 text-zinc-600"}>
                                {isThreeMonthConsecutiveEligible ? "Eligible" : "Not Eligible"}
                              </Badge>
                            </div>
                          </div>
                          
                          <div className="mt-2 pt-2 border-t border-zinc-200 text-[10px] space-y-1">
                            <p className="font-bold uppercase tracking-wider text-zinc-500">Last 3 Months Status:</p>
                            {uniqueMonths.slice(0, 3).map(m => {
                              const data = monthsMap[m];
                              const met = data && data.joinedWeeks > 0 && data.sales >= data.target;
                              return (
                                <div key={m} className="flex justify-between items-center">
                                  <span className="text-zinc-600 font-medium">{m}:</span>
                                  <div className="flex items-center gap-2">
                                    <span className={cn("font-bold", met ? "text-green-600" : "text-zinc-400")}>
                                      ₹{data?.sales.toLocaleString() || 0} / ₹{data?.target.toLocaleString() || 0}
                                    </span>
                                    {met ? (
                                      <CheckCircle2 className="w-3 h-3 text-green-600" />
                                    ) : (
                                      <AlertCircle className="w-3 h-3 text-zinc-300" />
                                    )}
                                  </div>
                                </div>
                              );
                            })}
                          </div>
                        </div>
                        
                        <div className="mt-4 pt-4 border-t border-zinc-100">
                          <p className="text-[10px] font-bold text-zinc-500 uppercase mb-2 tracking-wider">Terms & Conditions:</p>
                          <ul className="text-[10px] text-zinc-400 italic space-y-1 list-disc pl-4">
                            <li>Subject to verification.</li>
                            <li>If any major flight cost then if will be excluded from package cost.</li>
                            <li>If average margin in packages is less than 10% then package is not considered in calculation of incentive.</li>
                            <li>If any package is cancelled or credit note is provided then it is not considered for incentive.</li>
                            <li>Any other terms and condition may apply.</li>
                            <li>Incentives are for long term success, disbursal is subject to management.</li>
                          </ul>
                        </div>
                      </div>
                    </Card>

                    {/* Incentive Calculator */}
                    <Card className="p-6 bg-zinc-900 text-white">
                      <h3 className="text-lg font-bold mb-4 border-b border-zinc-700 pb-2 flex items-center gap-2">
                        <Calculator className="w-5 h-5 text-orange-500" />
                        Incentive Calculator
                      </h3>
                      <div className="space-y-6">
                        <div className="p-3 bg-zinc-800/50 rounded-lg border border-zinc-700/50 mb-4">
                          <p className="text-[10px] text-zinc-400 uppercase font-bold mb-1 tracking-wider">Current Targets</p>
                          <div className="flex justify-between text-xs font-medium">
                            <span>Weekly Base: ₹{(agent.target / 4).toLocaleString()}</span>
                            <span>Monthly (4-Week): ₹{fourWeekTarget.toLocaleString()}</span>
                          </div>
                        </div>

                        <div className="space-y-2">
                          <label className="text-xs font-bold text-zinc-400 uppercase">Hypothetical Weekly Net Sales (₹)</label>
                          <input 
                            type="number"
                            className="w-full px-4 py-2 bg-zinc-800 border border-zinc-700 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-white"
                            placeholder="Enter weekly sales..."
                            value={incentiveCalcSales || ''}
                            onChange={(e) => setIncentiveCalcSales(parseFloat(e.target.value) || 0)}
                          />
                          <div className="flex items-center gap-2 mt-1">
                            <input 
                              type="checkbox" 
                              id="assumeEligible"
                              checked={assumeWeeklyEligible}
                              onChange={(e) => setAssumeWeeklyEligible(e.target.checked)}
                              className="w-3 h-3 rounded border-zinc-700 bg-zinc-800 text-orange-500 focus:ring-orange-500"
                            />
                            <label htmlFor="assumeEligible" className="text-[10px] text-zinc-500">Assume 8-week Bravo &gt; 70% eligibility</label>
                          </div>
                        </div>
                        <div className="space-y-2">
                          <label className="text-xs font-bold text-zinc-400 uppercase">Hypothetical Monthly Net Sales (₹)</label>
                          <input 
                            type="number"
                            className="w-full px-4 py-2 bg-zinc-800 border border-zinc-700 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-white"
                            placeholder="Enter monthly sales..."
                            value={incentiveCalcMonthlySales || ''}
                            onChange={(e) => setIncentiveCalcMonthlySales(parseFloat(e.target.value) || 0)}
                          />
                        </div>

                        <div className="pt-4 border-t border-zinc-700 space-y-3">
                          <div className="flex justify-between items-center">
                            <span className="text-sm text-zinc-400">Potential Weekly Incentive:</span>
                            <div className="text-right">
                              <p className="text-lg font-bold text-orange-500">
                                ₹{((isWeeklyEligible || assumeWeeklyEligible) ? incentiveCalcSales * 0.001 : 0).toLocaleString()} approx
                              </p>
                              {!(isWeeklyEligible || assumeWeeklyEligible) && (
                                <p className="text-[9px] text-red-400 font-medium">Requires 70%+ Bravo Point Achievement</p>
                              )}
                            </div>
                          </div>
                          <div className="flex justify-between items-center">
                            <span className="text-sm text-zinc-400">Potential Monthly Incentive:</span>
                            <div className="text-right">
                              <p className="text-lg font-bold text-orange-500">
                                ₹{(incentiveCalcMonthlySales >= fourWeekTarget ? ((incentiveCalcMonthlySales - fourWeekTarget) * 0.03 + 2000) : 0).toLocaleString()} approx
                              </p>
                              {incentiveCalcMonthlySales >= fourWeekTarget ? (
                                <p className="text-[10px] text-zinc-500">
                                  Breakdown: ₹{((incentiveCalcMonthlySales - fourWeekTarget) * 0.03).toLocaleString()} (3%) + ₹2,000
                                </p>
                              ) : (
                                <p className="text-[9px] text-red-400 font-medium">Requires meeting monthly target (₹{fourWeekTarget.toLocaleString()})</p>
                              )}
                            </div>
                          </div>
                          <div className="flex justify-between items-center pt-2 border-t border-zinc-700">
                            <span className="font-bold">Total Potential Earnings:</span>
                            <span className="text-2xl font-black text-green-400">
                              ₹{(
                                ((isWeeklyEligible || assumeWeeklyEligible) ? incentiveCalcSales * 0.001 : 0) + 
                                (incentiveCalcMonthlySales >= fourWeekTarget ? ((incentiveCalcMonthlySales - fourWeekTarget) * 0.03 + 2000) : 0)
                              ).toLocaleString()} approx
                            </span>
                          </div>
                        </div>
                      </div>
                    </Card>
                  </div>

                  {/* Working Guidelines */}
                  <Card className="p-6 border-zinc-200 shadow-sm mt-6">
                    <div className="flex items-center gap-2 mb-4 border-b pb-2">
                      <Clock className="w-5 h-5 text-zinc-600" />
                      <h3 className="text-lg font-bold">Working Guidelines</h3>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse border border-zinc-200">
                        <thead>
                          <tr className="bg-zinc-50 text-xs font-bold text-zinc-500 uppercase">
                            <th className="px-4 py-3 border border-zinc-200">Guideline</th>
                            <th className="px-4 py-3 border border-zinc-200">Criteria</th>
                            <th className="px-4 py-3 border border-zinc-200">Current Status</th>
                            <th className="px-4 py-3 border border-zinc-200">Requirement</th>
                          </tr>
                        </thead>
                        <tbody>
                          <tr className="hover:bg-zinc-50 transition-colors">
                            <td className="px-4 py-3 text-sm font-medium border border-zinc-200">Floor Timings</td>
                            <td className="px-4 py-3 text-xs border border-zinc-200">
                              <div className="space-y-1">
                                <p>&gt; 90% Bravo: 8:30 Mins</p>
                                <p>70-90% Bravo: 9:30 Mins</p>
                                <p>&lt; 70% Bravo: 10:30 Mins</p>
                                <p className="text-zinc-400 italic">(Implemented after 4 weeks)</p>
                              </div>
                            </td>
                            <td className="px-4 py-3 text-sm border border-zinc-200">
                              <div className="space-y-1">
                                <p>Tenure: {tenureWeeks} weeks</p>
                                <p>8-Wk Bravo: {eightWeekAchievement}%</p>
                              </div>
                            </td>
                            <td className="px-4 py-3 text-sm border border-zinc-200 font-bold">
                              {tenureWeeks > 4 ? (
                                eightWeekAchievement > 90 ? (
                                  <span className="text-green-600">8:30 Minutes</span>
                                ) : eightWeekAchievement >= 70 ? (
                                  <span className="text-orange-600">9:30 Minutes</span>
                                ) : (
                                  <span className="text-red-600">10:30 Minutes</span>
                                )
                              ) : (
                                <span className="text-blue-600 font-bold">9:30 AM to 8:00 PM flat</span>
                              )}
                            </td>
                          </tr>
                          <tr className="hover:bg-zinc-50 transition-colors">
                            <td className="px-4 py-3 text-sm font-medium border border-zinc-200">HR Review</td>
                            <td className="px-4 py-3 text-xs border border-zinc-200">
                              <p>Bravo &lt; 30% after 4 weeks</p>
                            </td>
                            <td className="px-4 py-3 text-sm border border-zinc-200">
                              <p>8-Wk Bravo: {eightWeekAchievement}%</p>
                            </td>
                            <td className="px-4 py-3 text-sm border border-zinc-200">
                              {tenureWeeks > 4 && eightWeekAchievement < 30 ? (
                                <Badge className="bg-red-100 text-red-700 font-bold animate-pulse">Talk to HR</Badge>
                              ) : (
                                <span className="text-green-600 font-medium">Clear</span>
                              )}
                            </td>
                          </tr>
                          <tr className="hover:bg-zinc-50 transition-colors">
                            <td className="px-4 py-3 text-sm font-medium border border-zinc-200">2nd/4th Sunday working</td>
                            <td className="px-4 py-3 text-xs border border-zinc-200">
                              <p>After 6 weeks: Both 4-Wk & 8-Wk Bravo &lt; 70%</p>
                            </td>
                            <td className="px-4 py-3 text-sm border border-zinc-200">
                              <div className="space-y-1">
                                <p>4-Wk Bravo: {fourWeekAchievement}%</p>
                                <p>8-Wk Bravo: {eightWeekAchievement}%</p>
                              </div>
                            </td>
                            <td className="px-4 py-3 text-sm border border-zinc-200 font-bold">
                              {tenureWeeks > 6 && fourWeekAchievement < 70 && eightWeekAchievement < 70 ? (
                                <span className="text-red-600">Required</span>
                              ) : (
                                <span className="text-green-600 font-medium">Standard Offs</span>
                              )}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                    <div className="mt-4 pt-4 border-t border-zinc-100">
                      <p className="text-xs font-bold text-zinc-500 uppercase mb-2 tracking-wider">Terms and Conditions for all:</p>
                      <p className="text-xs text-zinc-600 italic">
                        This is not a shift based job, the timing mentioned is applicable but the work must be completed perfectly.
                      </p>
                    </div>
                  </Card>

                  {/* Training Completion */}
                  <Card className="p-6">
                    <div className="flex items-center justify-between mb-4">
                      <div className="space-y-1">
                        <h3 className="text-lg font-bold">Training Completion</h3>
                        <div className="flex flex-wrap gap-x-4 gap-y-1">
                          <div className="flex items-center gap-1.5">
                            <CheckCircle2 className="w-3.5 h-3.5 text-green-600" />
                            <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">
                              Tasks: <span className="text-zinc-900">{completedTasksCount} / {totalTasksCount}</span>
                            </span>
                          </div>
                          <div className="flex items-center gap-1.5">
                            <Trophy className="w-3.5 h-3.5 text-orange-500" />
                            <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">
                              Points: <span className="text-zinc-900">{earnedTrainingPoints.toLocaleString()} / {totalTrainingPoints.toLocaleString()}</span>
                            </span>
                          </div>
                        </div>
                      </div>
                      <span className="text-2xl font-bold text-zinc-800">{trainingCompletion}%</span>
                    </div>
                    <div className="w-full bg-zinc-100 rounded-full h-4">
                      <div className="bg-zinc-800 h-4 rounded-full transition-all duration-1000" style={{ width: `${trainingCompletion}%` }}></div>
                    </div>
                    
                    <div className="mt-8 space-y-4">
                      <div className="space-y-4 border-b pb-4">
                        <h4 className="text-sm font-bold text-zinc-500 uppercase tracking-wider">Day-wise Training Tasks</h4>
                        <div className="flex gap-2 overflow-x-auto pb-2 no-scrollbar -mx-1 px-1">
                          {Array.from(new Set(activeTrainingMaterials.map(tm => tm.day))).sort((a: any, b: any) => a - b).map(day => {
                            const dayTasks = activeTrainingMaterials.filter(tm => tm.day === day);
                            const totalTasks = dayTasks.length;
                            const completedTasks = dayTasks.filter(tm => (agent.trainingPoints?.[tm.id!] || 0) >= tm.points).length;
                            const isFullyCompleted = totalTasks > 0 && completedTasks === totalTasks;

                            return (
                              <button
                                key={day}
                                onClick={() => setSelectedAgentOverviewDay(day)}
                                className={cn(
                                  "px-3 py-2 rounded-xl text-[10px] font-bold transition-all whitespace-nowrap flex flex-col items-center gap-1 min-w-[70px] flex-shrink-0",
                                  selectedAgentOverviewDay === day 
                                    ? "bg-orange-600 text-white shadow-lg scale-105" 
                                    : isFullyCompleted
                                      ? "bg-green-100 text-green-700 border border-green-200 hover:bg-green-200"
                                      : "bg-zinc-100 text-zinc-600 hover:bg-zinc-200 border border-zinc-200"
                                )}
                              >
                                <span>Day {day}</span>
                                <span className={cn(
                                  "text-[9px] opacity-80",
                                  isFullyCompleted && selectedAgentOverviewDay !== day ? "text-green-600" : ""
                                )}>
                                  {completedTasks}/{totalTasks}
                                </span>
                              </button>
                            );
                          })}
                        </div>
                      </div>

                      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 bg-zinc-50 p-4 rounded-xl border border-zinc-200">
                        <div className="flex items-center gap-4">
                          <div className="flex flex-col">
                            <span className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">Status Filter</span>
                            <div className="flex gap-1 mt-1">
                              {(['All', 'Pending', 'Complete'] as const).map(status => (
                                <button
                                  key={status}
                                  onClick={() => setTrainingStatusOverviewFilter(status)}
                                  className={cn(
                                    "px-3 py-1 rounded-md text-[10px] font-bold uppercase transition-all",
                                    trainingStatusOverviewFilter === status
                                      ? "bg-zinc-800 text-white"
                                      : "bg-white text-zinc-600 border border-zinc-200 hover:bg-zinc-100"
                                  )}
                                >
                                  {status}
                                </button>
                              ))}
                            </div>
                          </div>
                        </div>
                        
                        {(() => {
                          const dayTasks = activeTrainingMaterials.filter(tm => tm.day === selectedAgentOverviewDay);
                          const completedCount = dayTasks.filter(tm => (agent.trainingPoints?.[tm.id!] || 0) >= tm.points).length;
                          return (
                            <div className="flex items-center gap-2 bg-white px-4 py-2 rounded-lg border border-zinc-200 shadow-sm">
                              <CheckCircle2 className="w-4 h-4 text-green-600" />
                              <span className="text-sm font-bold text-zinc-700">
                                Day {selectedAgentOverviewDay} Progress: <span className="text-green-600">{completedCount}</span> / <span className="text-zinc-400">{dayTasks.length}</span>
                              </span>
                            </div>
                          );
                        })()}
                      </div>

                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse border border-zinc-200">
                          <thead>
                            <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase">
                              <th className="px-4 py-3 border border-zinc-200">Priority</th>
                              <th className="px-4 py-3 border border-zinc-200">Task Details</th>
                              <th className="px-4 py-3 border border-zinc-200">Module</th>
                              <th className="px-4 py-3 border border-zinc-200">Time (Min)</th>
                              <th className="px-4 py-3 border border-zinc-200">Points Allotted</th>
                              <th className="px-4 py-3 border border-zinc-200">Points Assigned</th>
                              <th className="px-4 py-3 border border-zinc-200">Status</th>
                              <th className="px-4 py-3 border border-zinc-200">Link</th>
                            </tr>
                          </thead>
                          <tbody>
                            {(() => {
                              let dayTasks = activeTrainingMaterials.filter(tm => tm.day === selectedAgentOverviewDay);
                              
                              // Sort by priority (ascending)
                              dayTasks = [...dayTasks].sort((a, b) => (a.priority || 0) - (b.priority || 0));

                              // Apply status filter
                              if (trainingStatusOverviewFilter !== 'All') {
                                dayTasks = dayTasks.filter(tm => {
                                  const assignedPoints = agent.trainingPoints?.[tm.id!] || 0;
                                  const isCompleted = assignedPoints >= tm.points;
                                  return trainingStatusOverviewFilter === 'Complete' ? isCompleted : !isCompleted;
                                });
                              }

                              if (dayTasks.length === 0) {
                                return (
                                  <tr>
                                    <td colSpan={8} className="px-4 py-8 text-center text-zinc-400 italic text-sm">
                                      {trainingStatusOverviewFilter === 'All' 
                                        ? `No training tasks assigned for Day ${selectedAgentOverviewDay}`
                                        : `No ${trainingStatusOverviewFilter.toLowerCase()} tasks for Day ${selectedAgentOverviewDay}`}
                                    </td>
                                  </tr>
                                );
                              }
                              return dayTasks.map(tm => {
                                const assignedPoints = agent.trainingPoints?.[tm.id!] || 0;
                                const isCompleted = assignedPoints >= tm.points;
                                return (
                                  <tr key={tm.id} className="hover:bg-zinc-50 transition-colors">
                                    <td className="px-4 py-3 text-sm border border-zinc-200 font-bold text-zinc-500">
                                      {tm.priority || 0}
                                    </td>
                                    <td className="px-4 py-3 text-sm border border-zinc-200">
                                      <div className="font-medium text-zinc-900">{tm.remarks}</div>
                                    </td>
                                    <td className="px-4 py-3 text-xs border border-zinc-200">
                                      <div className="flex flex-wrap gap-1">
                                        {(Array.isArray(tm.module) ? tm.module : [tm.module]).map((m, idx) => (
                                          <Badge key={idx} className="bg-zinc-100 text-zinc-600">{m}</Badge>
                                        ))}
                                      </div>
                                    </td>
                                    <td className="px-4 py-3 text-sm border border-zinc-200 text-zinc-600">{tm.time}</td>
                                    <td className="px-4 py-3 text-sm border border-zinc-200 font-bold text-zinc-700">{tm.points}</td>
                                    <td className="px-4 py-3 text-sm border border-zinc-200 font-bold text-orange-600">{assignedPoints}</td>
                                    <td className="px-4 py-3 text-xs border border-zinc-200">
                                      {isCompleted ? (
                                        <Badge className="bg-green-100 text-green-700">Completed</Badge>
                                      ) : assignedPoints > 0 ? (
                                        <Badge className="bg-orange-100 text-orange-700">Partial</Badge>
                                      ) : (
                                        <Badge className="bg-zinc-100 text-zinc-400">Pending</Badge>
                                      )}
                                    </td>
                                    <td className="px-4 py-3 text-sm border border-zinc-200">
                                      {tm.link ? (
                                        <a 
                                          href={tm.link} 
                                          target="_blank" 
                                          rel="noopener noreferrer"
                                          className="text-orange-600 hover:text-orange-700 flex items-center gap-1 font-medium"
                                        >
                                          <ExternalLink className="w-3 h-3" />
                                          View
                                        </a>
                                      ) : (
                                        <span className="text-zinc-400">-</span>
                                      )}
                                    </td>
                                  </tr>
                                );
                              });
                            })()}
                          </tbody>
                        </table>
                      </div>
                    </div>
                  </Card>

                  {/* Employee Summary Table (Approx) */}
                  <Card className="overflow-hidden border-zinc-200 shadow-sm">
                    <div className="bg-zinc-50 px-4 py-2 border-b border-zinc-200 flex justify-between items-center">
                      <h3 className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">Employee Summary (Incentive + Quaterly) (Approx)</h3>
                      <div className="flex items-center gap-4 text-[10px] font-bold uppercase tracking-wider">
                        <div className="flex items-center gap-1.5">
                          <div className="w-3 h-3 bg-white border border-zinc-300 rounded-sm"></div>
                          <span className="text-zinc-500">White - Due</span>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <div className="w-3 h-3 bg-purple-100 border border-purple-200 rounded-sm"></div>
                          <span className="text-purple-600">Purple - Eligible to be paid</span>
                        </div>
                        <div className="flex items-center gap-1.5">
                          <div className="w-3 h-3 bg-green-100 border border-green-200 rounded-sm"></div>
                          <span className="text-green-600">Green - Paid</span>
                        </div>
                      </div>
                    </div>
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-white text-[10px] font-bold text-zinc-400 uppercase">
                            <th className="px-4 py-2 border-r">Month</th>
                            <th className="px-4 py-2 border-r">Daily</th>
                            <th className="px-4 py-2 border-r">Weekly</th>
                            <th className="px-4 py-2 border-r">Monthly</th>
                            <th className="px-4 py-2 border-r">Quarterly</th>
                            <th className="px-4 py-2 border-r">Annually</th>
                            <th className="px-4 py-2 border-r">Total Approx Due</th>
                            <th className="px-4 py-2 border-r">Total Approx Paid</th>
                            <th className="px-4 py-2 bg-orange-50 text-orange-600">Approx Balance</th>
                          </tr>
                        </thead>
                        <tbody>
                          {agentOverviewMonthlyIncentiveData.map((row: any) => (
                            <tr key={row.month} className="text-sm font-medium border-b border-zinc-100 hover:bg-zinc-50 transition-colors text-zinc-600">
                              <td className="px-4 py-2 border-r font-bold text-zinc-900">{row.month}</td>
                              <td className={cn("px-4 py-2 border-r", row.typeStatuses.Daily.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Daily.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Daily.toLocaleString()}</td>
                              <td className={cn("px-4 py-2 border-r", row.typeStatuses.Weekly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Weekly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Weekly.toLocaleString()}</td>
                              <td className={cn("px-4 py-2 border-r", row.typeStatuses.Monthly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Monthly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Monthly.toLocaleString()}</td>
                              <td className={cn("px-4 py-2 border-r", row.typeStatuses.Quarterly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Quarterly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Quarterly.toLocaleString()}</td>
                              <td className={cn("px-4 py-2 border-r", row.typeStatuses.Annually.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Annually.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Annually.toLocaleString()}</td>
                              <td 
                                className="px-4 py-2 border-r text-red-600 cursor-pointer hover:bg-red-50 transition-colors font-bold"
                                onClick={() => setAgentLedgerModal({ isOpen: true, employeeId: selectedAgentOverview })}
                              >
                                ₹{row.totalDue.toLocaleString()}
                              </td>
                              <td className="px-4 py-2 border-r text-green-600 font-bold">₹{row.totalPaid.toLocaleString()}</td>
                              <td className="px-4 py-2 bg-orange-50 text-orange-600 font-bold">₹{(row.totalDue - row.totalPaid).toLocaleString()}</td>
                            </tr>
                          ))}
                          <tr className="text-sm font-black text-zinc-900 bg-zinc-50">
                            <td className="px-4 py-3 border-r uppercase">Grand Total</td>
                            <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === selectedAgentOverview && i.type === 'Daily').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                            <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === selectedAgentOverview && i.type === 'Weekly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                            <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === selectedAgentOverview && i.type === 'Monthly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                            <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === selectedAgentOverview && i.type === 'Quarterly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                            <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === selectedAgentOverview && i.type === 'Annually').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                            <td 
                              className="px-4 py-3 border-r text-red-600 cursor-pointer hover:bg-red-100 transition-colors font-bold"
                              onClick={() => setAgentLedgerModal({ isOpen: true, employeeId: selectedAgentOverview })}
                            >
                              ₹{incentives.filter(i => i.employeeId === selectedAgentOverview).reduce((sum, i) => sum + i.amount, 0).toLocaleString()}
                            </td>
                            <td className="px-4 py-3 border-r text-green-600 font-bold">₹{incentivePayments.filter(p => p.employeeId === selectedAgentOverview).reduce((sum, p) => sum + p.amount, 0).toLocaleString()}</td>
                            <td className="px-4 py-3 bg-orange-100 text-orange-700 text-lg font-bold">
                              ₹{(
                                incentives.filter(i => i.employeeId === selectedAgentOverview).reduce((sum, i) => sum + i.amount, 0) - 
                                incentivePayments.filter(p => p.employeeId === selectedAgentOverview).reduce((sum, p) => sum + p.amount, 0)
                              ).toLocaleString()}
                            </td>
                          </tr>
                        </tbody>
                      </table>
                    </div>
                  </Card>
                </div>
              );
            })()}
          </div>
        )}
        {activeTab === 'bdeOverview' && hasPermission('bdeOverview') && (
          <div className="space-y-8">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-center gap-2">
                <BarChart3 className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">BDE Overview</h2>
              </div>
              <div className="flex flex-wrap items-center gap-3 min-w-[200px]">
                <SearchableSelect
                  options={[
                    { value: '', label: 'Select BDE' },
                    ...(getPermissionLevel('bdeOverview') === 'Complete' ? [{ value: 'All', label: 'All BDE' }] : []),
                    ...bdes.filter(bde => {
                      const level = getPermissionLevel('bdeOverview');
                      if (level === 'Complete') return true;
                      if (level === 'Limited') return bde.email?.toLowerCase() === profile?.email?.toLowerCase();
                      return false;
                    }).map(bde => ({ value: bde.name, label: bde.name, key: bde.id }))
                  ]}
                  value={selectedBDEForOverview}
                  onChange={setSelectedBDEForOverview}
                  placeholder="Select BDE"
                />
                <div className="flex bg-zinc-100 p-1 rounded-lg overflow-x-auto">
                  {(['Daily', 'Weekly', '4-Weekly', '8-Weekly', 'Monthly'] as const).map((tf) => (
                    <button
                      key={tf}
                      onClick={() => setBdeOverviewTimeframe(tf)}
                      className={cn(
                        "px-4 py-1.5 text-xs font-bold rounded-md transition-all whitespace-nowrap",
                        bdeOverviewTimeframe === tf 
                          ? "bg-white text-orange-600 shadow-sm" 
                          : "text-zinc-500 hover:text-zinc-700"
                      )}
                    >
                      {tf}
                    </button>
                  ))}
                </div>
                {bdeOverviewTimeframe === 'Monthly' && (
                  <select 
                    className="px-4 py-2 bg-white border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm font-medium"
                    value={selectedMonthForBDEOverview}
                    onChange={(e) => setSelectedMonthForBDEOverview(e.target.value)}
                  >
                    <option value="">Select Month</option>
                    {Array.from(new Set(allSortedWeeks.map(w => w.month))).map(m => (
                      <option key={m} value={m}>{m}</option>
                    ))}
                  </select>
                )}
              </div>
            </div>

            {(() => {
              if (!selectedBDEForOverview) {
                return (
                  <Card className="p-12 flex flex-col items-center justify-center text-zinc-400">
                    <Users className="w-12 h-12 mb-4 opacity-20" />
                    <p className="text-lg font-medium">Please select a BDE to view analytics</p>
                  </Card>
                );
              }

              const bdeAgents = employees.filter(emp => 
                (selectedBDEForOverview === 'All' || emp.bde === selectedBDEForOverview) && 
                emp.status === 'Active'
              );
              const bdeSales = sales.filter(sale => {
                if (selectedBDEForOverview === 'All') return true;
                const agentBde = findEmployeeBySale(sale.agent, sale.agentEmail)?.bde;
                const assocBde = findEmployeeBySale(sale.associateBde, sale.associateBdeEmail)?.bde;
                const bdeByEmail = findBDEBySale(sale.bde, sale.bdeEmail)?.name;
                return agentBde === selectedBDEForOverview || assocBde === selectedBDEForOverview || bdeByEmail === selectedBDEForOverview || sale.bde === selectedBDEForOverview;
              });
              
              const currentWeek = allSortedWeeks[allSortedWeeks.length - 1];
              const last4Weeks = allSortedWeeks.slice(-4);
              const last8Weeks = allSortedWeeks.slice(-8);
              const currentMonth = bdeOverviewTimeframe === 'Monthly' && selectedMonthForBDEOverview 
                ? selectedMonthForBDEOverview 
                : (currentWeek?.month || 'Current Month');
              const monthWeeks = allSortedWeeks.filter(w => w.month === currentMonth);
              const today = new Date().toISOString().split('T')[0];

              let filteredSales = [];
              let timeframeLabel = '';
              let dateRangeLabel = '';

              const bdeComparisonData = bdes.map(bde => {
                const agents = employees.filter(emp => emp.bde === bde.name && emp.status === 'Active');
                
                const bdeSalesForThisBde = sales.filter(sale => {
                  const agentBde = findEmployeeBySale(sale.agent, sale.agentEmail)?.bde;
                  const assocBde = findEmployeeBySale(sale.associateBde, sale.associateBdeEmail)?.bde;
                  const bdeByEmail = findBDEBySale(sale.bde, sale.bdeEmail)?.name;
                  return agentBde === bde.name || assocBde === bde.name || bdeByEmail === bde.name || sale.bde === bde.name;
                });

                let filteredSalesForBde = [];
                if (bdeOverviewTimeframe === 'Daily') {
                  filteredSalesForBde = bdeSalesForThisBde.filter(s => s.date === today);
                } else if (bdeOverviewTimeframe === 'Weekly') {
                  filteredSalesForBde = bdeSalesForThisBde.filter(s => s.week === currentWeek?.weekName);
                } else if (bdeOverviewTimeframe === '4-Weekly') {
                  const last4WeekNames = last4Weeks.map(w => w.weekName);
                  filteredSalesForBde = bdeSalesForThisBde.filter(s => last4WeekNames.includes(s.week));
                } else if (bdeOverviewTimeframe === '8-Weekly') {
                  const last8WeekNames = last8Weeks.map(w => w.weekName);
                  filteredSalesForBde = bdeSalesForThisBde.filter(s => last8WeekNames.includes(s.week));
                } else if (bdeOverviewTimeframe === 'Monthly') {
                  const monthWeekNames = monthWeeks.map(w => w.weekName);
                  filteredSalesForBde = bdeSalesForThisBde.filter(s => monthWeekNames.includes(s.week));
                }

                const bdeTotalSales = filteredSalesForBde.reduce((sum, s) => {
                  let credit = 0;
                  if (findEmployeeBySale(s.agent, s.agentEmail)?.bde === bde.name) {
                    credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                  }
                  if (findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === bde.name) {
                    credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                  }
                  return sum + credit;
                }, 0);

                const bdeApproxMargin = bdeTotalSales * 0.12;

                const bdeExpenses = agents.reduce((sum, emp) => {
                  const monthlyTotalExpense = emp.salary * 3;
                  let adjustedExpense = 0;
                  if (bdeOverviewTimeframe === 'Daily') adjustedExpense = monthlyTotalExpense / 30;
                  else if (bdeOverviewTimeframe === 'Weekly') adjustedExpense = monthlyTotalExpense / 4;
                  else if (bdeOverviewTimeframe === '4-Weekly') adjustedExpense = monthlyTotalExpense;
                  else if (bdeOverviewTimeframe === '8-Weekly') adjustedExpense = monthlyTotalExpense * 2;
                  else if (bdeOverviewTimeframe === 'Monthly') adjustedExpense = monthlyTotalExpense;
                  return sum + adjustedExpense;
                }, 0);

                const bdeNetProfit = bdeApproxMargin - bdeExpenses;

                const bdeTotalSalary = agents.reduce((sum, emp) => {
                  const monthlySalary = emp.salary;
                  let adjustedSalary = 0;
                  if (bdeOverviewTimeframe === 'Daily') adjustedSalary = monthlySalary / 30;
                  else if (bdeOverviewTimeframe === 'Weekly') adjustedSalary = monthlySalary / 4;
                  else if (bdeOverviewTimeframe === '4-Weekly') adjustedSalary = monthlySalary;
                  else if (bdeOverviewTimeframe === '8-Weekly') adjustedSalary = monthlySalary * 2;
                  else if (bdeOverviewTimeframe === 'Monthly') adjustedSalary = monthlySalary;
                  return sum + adjustedSalary;
                }, 0);

                const salesSalaryRatio = bdeTotalSalary > 0 ? bdeTotalSales / bdeTotalSalary : 0;

                return {
                  name: bde.name,
                  totalSales: bdeTotalSales,
                  approxMargin: bdeApproxMargin,
                  expenses: bdeExpenses,
                  netProfit: bdeNetProfit,
                  salesSalaryRatio
                };
              });
              let totalTarget = 0;

              if (bdeOverviewTimeframe === 'Daily') {
                filteredSales = bdeSales.filter(s => s.date === today);
                timeframeLabel = `Today (${today})`;
                dateRangeLabel = today;
                totalTarget = bdeAgents.reduce((sum, emp) => sum + (emp.target / 24), 0) * 0.75;
              } else if (bdeOverviewTimeframe === 'Weekly') {
                filteredSales = bdeSales.filter(s => s.week === currentWeek?.weekName);
                timeframeLabel = currentWeek?.weekName || 'Current Week';
                dateRangeLabel = currentWeek ? `${currentWeek.startDate} to ${currentWeek.endDate}` : '';
                const actualTarget = bdeAgents.reduce((sum, emp) => {
                  const weeklyBaseTarget = emp.target / 4;
                  const wkIndex = allSortedWeeks.findIndex(w => w.weekName === currentWeek?.weekName);
                  const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === emp.joiningWeek);
                  let percentageVal = 100;
                  if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                    const tenureWeek = wkIndex - joiningWkIndex + 1;
                    if (tenureWeek <= 6) percentageVal = (emp as any)[`week${tenureWeek}Target`] || 0;
                  }
                  return sum + (weeklyBaseTarget * percentageVal) / 100;
                }, 0);
                totalTarget = actualTarget * 0.75;
              } else if (bdeOverviewTimeframe === '4-Weekly') {
                const last4WeekNames = last4Weeks.map(w => w.weekName);
                filteredSales = bdeSales.filter(s => last4WeekNames.includes(s.week));
                timeframeLabel = 'Last 4 Weeks';
                dateRangeLabel = last4Weeks.length > 0 ? `${last4Weeks[0].startDate} to ${last4Weeks[last4Weeks.length-1].endDate}` : '';
                totalTarget = bdeAgents.reduce((sum, emp) => sum + emp.target, 0) * 0.75;
              } else if (bdeOverviewTimeframe === '8-Weekly') {
                const last8WeekNames = last8Weeks.map(w => w.weekName);
                filteredSales = bdeSales.filter(s => last8WeekNames.includes(s.week));
                timeframeLabel = 'Last 8 Weeks';
                dateRangeLabel = last8Weeks.length > 0 ? `${last8Weeks[0].startDate} to ${last8Weeks[last8Weeks.length-1].endDate}` : '';
                totalTarget = bdeAgents.reduce((sum, emp) => sum + (emp.target * 2), 0) * 0.75;
              } else if (bdeOverviewTimeframe === 'Monthly') {
                const monthWeekNames = monthWeeks.map(w => w.weekName);
                filteredSales = bdeSales.filter(s => monthWeekNames.includes(s.week));
                timeframeLabel = `Month: ${currentMonth}`;
                dateRangeLabel = monthWeeks.length > 0 ? `${monthWeeks[0].startDate} to ${monthWeeks[monthWeeks.length-1].endDate}` : '';
                totalTarget = bdeAgents.reduce((sum, emp) => sum + emp.target, 0) * 0.75;
              }

              const totalSales = filteredSales.reduce((sum, s) => {
                let credit = 0;
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.agent, s.agentEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                }
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                }
                return sum + credit;
              }, 0);
              const numSales = filteredSales.length;
              const totalSalesValue = totalSales;
              const totalMarginValue = filteredSales.reduce((sum, s) => {
                let credit = 0;
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.agent, s.agentEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.totalMargin || 0) * ((s.agentPercentage ?? 100) / 100);
                }
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.totalMargin || 0) * ((s.assocBdePercentage ?? 0) / 100);
                }
                return sum + credit;
              }, 0);
              const avgPackageValue = numSales > 0 ? totalSalesValue / numSales : 0;

              const advanceSales = filteredSales.filter(s => s.advanceCN === 'Advance');
              const advanceCount = advanceSales.length;
              const advanceValue = advanceSales.reduce((sum, s) => {
                let credit = 0;
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.agent, s.agentEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                }
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                }
                return sum + credit;
              }, 0);
              const advancePerc = totalSalesValue > 0 ? (advanceValue / totalSalesValue) * 100 : 0;

              const cnSales = filteredSales.filter(s => s.advanceCN === 'Credit Note');
              const cnCount = cnSales.length;
              const cnValue = cnSales.reduce((sum, s) => {
                let credit = 0;
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.agent, s.agentEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                }
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                }
                return sum + credit;
              }, 0);
              const cnPerc = totalSalesValue > 0 ? (cnValue / totalSalesValue) * 100 : 0;

              const cancelSales = filteredSales.filter(s => s.advanceCN === 'Cancel');
              const cancelCount = cancelSales.length;
              const cancelValue = cancelSales.reduce((sum, s) => {
                let credit = 0;
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.agent, s.agentEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                }
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                }
                return sum + credit;
              }, 0);
              const cancelPerc = totalSalesValue > 0 ? (cancelValue / totalSalesValue) * 100 : 0;

              const workPendingCount = filteredSales.filter(s => s.workPending === 'Issue').length;

              const approxMargin = totalSales * 0.12;
              
              const totalExpenses = bdeAgents.reduce((sum, emp) => {
                const monthlyTotalExpense = emp.salary * 3;
                
                let adjustedExpense = 0;
                if (bdeOverviewTimeframe === 'Daily') adjustedExpense = monthlyTotalExpense / 30;
                else if (bdeOverviewTimeframe === 'Weekly') adjustedExpense = monthlyTotalExpense / 4;
                else if (bdeOverviewTimeframe === '4-Weekly') adjustedExpense = monthlyTotalExpense;
                else if (bdeOverviewTimeframe === '8-Weekly') adjustedExpense = monthlyTotalExpense * 2;
                else if (bdeOverviewTimeframe === 'Monthly') adjustedExpense = monthlyTotalExpense;
                return sum + adjustedExpense;
              }, 0);

              const profitability = approxMargin - totalExpenses;
              const achievement = totalTarget > 0 ? Math.round((totalSales / totalTarget) * 100) : 0;

              const bdeSalesStats = {
                inhouse: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
                branch: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
                franchisee: { total: 0, value: 0, advance: 0, advanceValue: 0, cn: 0, cnValue: 0, confirmed: 0, confirmedValue: 0, cancel: 0, cancelValue: 0 },
              };

              filteredSales.forEach(s => {
                let category: 'inhouse' | 'branch' | 'franchisee' = 'inhouse';
                if (s.salesBy === 'Branch') category = 'branch';
                else if (s.salesBy === 'Franchisee Sales') category = 'franchisee';

                let val = 0;
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.agent, s.agentEmail)?.bde === selectedBDEForOverview) {
                  val += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                }
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === selectedBDEForOverview) {
                  val += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                }

                if (val > 0) {
                  bdeSalesStats[category].total += 1;
                  bdeSalesStats[category].value += val;

                  if (s.advanceCN === 'Advance') {
                    bdeSalesStats[category].advance += 1;
                    bdeSalesStats[category].advanceValue += val;
                  } else if (s.advanceCN === 'Credit Note') {
                    bdeSalesStats[category].cn += 1;
                    bdeSalesStats[category].cnValue += val;
                  } else if (s.advanceCN === 'Done' || s.advanceCN === 'Confirmed') {
                    bdeSalesStats[category].confirmed += 1;
                    bdeSalesStats[category].confirmedValue += val;
                  } else if (s.advanceCN === 'Cancel') {
                    bdeSalesStats[category].cancel += 1;
                    bdeSalesStats[category].cancelValue += val;
                  }
                }
              });

              // 8-Week Performance Calculations for Header
              const last8WeekNamesHeader = last8Weeks.map(w => w.weekName);
              const eightWeekSalesHeader = bdeSales.filter(s => last8WeekNamesHeader.includes(s.week));
              const eightWeekTotalSalesHeader = eightWeekSalesHeader.reduce((sum, s) => {
                let credit = 0;
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.agent, s.agentEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.agentPercentage ?? 100) / 100);
                }
                if (selectedBDEForOverview === 'All' || findEmployeeBySale(s.associateBde, s.associateBdeEmail)?.bde === selectedBDEForOverview) {
                  credit += (s.packageValue || 0) * ((s.assocBdePercentage ?? 0) / 100);
                }
                return sum + credit;
              }, 0);
              const eightWeekApproxMarginHeader = eightWeekTotalSalesHeader * 0.12;
              
              const eightWeekTotalExpensesHeader = bdeAgents.reduce((sum, emp) => {
                const monthlyTotalExpense = emp.salary * 3;
                return sum + (monthlyTotalExpense / 4) * 8;
              }, 0);

              const eightWeekActualTargetHeader = bdeAgents.reduce((sum, emp) => sum + (emp.target / 4) * 8, 0);
              const eightWeekBdeTargetHeader = eightWeekActualTargetHeader * 0.75;
              
              const eightWeekProfitHeader = eightWeekApproxMarginHeader - eightWeekTotalExpensesHeader;
              const eightWeekAchievementHeader = eightWeekBdeTargetHeader > 0 ? Math.round((eightWeekTotalSalesHeader / eightWeekBdeTargetHeader) * 100) : 0;

              const getBDEStatus = (ach: number) => {
                if (ach >= 100) return { label: 'EXCELLENT', color: 'text-emerald-600', bg: 'bg-emerald-50', border: 'border-emerald-100' };
                if (ach >= 90) return { label: 'GOOD', color: 'text-blue-600', bg: 'bg-blue-50', border: 'border-blue-100' };
                if (ach >= 70) return { label: 'AVERAGE', color: 'text-orange-600', bg: 'bg-orange-50', border: 'border-orange-100' };
                return { label: 'BAD', color: 'text-red-600', bg: 'bg-red-50', border: 'border-red-100' };
              };
              const bdeStatus = getBDEStatus(eightWeekAchievementHeader);

              return (
                <div className="space-y-6">
                  {/* Performance Summary Boxes */}
                  <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-3">
                    <Card 
                      className="p-3 bg-orange-600 text-white border-none shadow-sm cursor-pointer hover:scale-[1.02] transition-transform"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Total Sales Details', sales: filteredSales })}
                    >
                      <p className="text-orange-100 text-[10px] font-bold uppercase tracking-wider">Sales Value</p>
                      <p className="text-lg font-black mt-1">₹{totalSalesValue.toLocaleString()}</p>
                      <p className="text-[9px] text-orange-100 mt-1">{numSales} Trans.</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Margin Details', sales: filteredSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Margin</p>
                      <p className="text-lg font-black mt-1 text-emerald-600">₹{totalMarginValue.toLocaleString()}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">{totalSalesValue > 0 ? ((totalMarginValue / totalSalesValue) * 100).toFixed(1) : 0}%</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Package Value Details', sales: filteredSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Avg Pkg</p>
                      <p className="text-lg font-black mt-1 text-blue-600">₹{Math.round(avgPackageValue).toLocaleString()}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">{numSales} Sales</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Advance Status Details', sales: advanceSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Advance</p>
                      <p className="text-lg font-black mt-1 text-orange-600">{advanceCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">₹{advanceValue.toLocaleString()} ({advancePerc.toFixed(1)}%)</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Credit Note Details', sales: cnSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">CN</p>
                      <p className="text-lg font-black mt-1 text-purple-600">{cnCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">₹{cnValue.toLocaleString()} ({cnPerc.toFixed(1)}%)</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Cancel Status Details', sales: cancelSales })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Cancel</p>
                      <p className="text-lg font-black mt-1 text-red-600">{cancelCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">₹{cancelValue.toLocaleString()} ({cancelPerc.toFixed(1)}%)</p>
                    </Card>
                    <Card 
                      className="p-3 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Work Pending Details', sales: filteredSales.filter(s => s.workPending === 'Issue') })}
                    >
                      <p className="text-zinc-400 text-[10px] font-bold uppercase tracking-wider">Pending</p>
                      <p className="text-lg font-black mt-1 text-amber-600">{workPendingCount}</p>
                      <p className="text-[9px] text-zinc-500 mt-1">Issues</p>
                    </Card>
                  </div>

                  {/* Summary Cards */}
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-4">
                    <Card className="p-4 bg-orange-50 border-orange-100">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-orange-100 rounded-lg">
                          <TrendingUp className="w-5 h-5 text-orange-600" />
                        </div>
                        <div className="flex-1">
                          <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Inhouse Sales</p>
                          <p className="text-2xl font-bold text-zinc-900">{bdeSalesStats.inhouse.total}</p>
                          <div className="mt-2 pt-2 border-t border-orange-200 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                            <div className="text-orange-700">Adv: {bdeSalesStats.inhouse.advance}</div>
                            <div className="text-orange-700">CN: {bdeSalesStats.inhouse.cn}</div>
                            <div className="text-orange-700">Conf: {bdeSalesStats.inhouse.confirmed}</div>
                            <div className="text-orange-700">Can: {bdeSalesStats.inhouse.cancel}</div>
                          </div>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-orange-50 border-orange-100">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-orange-100 rounded-lg">
                          <TrendingUp className="w-5 h-5 text-orange-600" />
                        </div>
                        <div className="flex-1">
                          <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Branch Sales</p>
                          <p className="text-2xl font-bold text-zinc-900">{bdeSalesStats.branch.total}</p>
                          <div className="mt-2 pt-2 border-t border-orange-200 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                            <div className="text-orange-700">Adv: {bdeSalesStats.branch.advance}</div>
                            <div className="text-orange-700">CN: {bdeSalesStats.branch.cn}</div>
                            <div className="text-orange-700">Conf: {bdeSalesStats.branch.confirmed}</div>
                            <div className="text-orange-700">Can: {bdeSalesStats.branch.cancel}</div>
                          </div>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-orange-50 border-orange-100">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-orange-100 rounded-lg">
                          <TrendingUp className="w-5 h-5 text-orange-600" />
                        </div>
                        <div className="flex-1">
                          <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Franchisee Sales</p>
                          <p className="text-2xl font-bold text-zinc-900">{bdeSalesStats.franchisee.total}</p>
                          <div className="mt-2 pt-2 border-t border-orange-200 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                            <div className="text-orange-700">Adv: {bdeSalesStats.franchisee.advance}</div>
                            <div className="text-orange-700">CN: {bdeSalesStats.franchisee.cn}</div>
                            <div className="text-orange-700">Conf: {bdeSalesStats.franchisee.confirmed}</div>
                            <div className="text-orange-700">Can: {bdeSalesStats.franchisee.cancel}</div>
                          </div>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-zinc-900 text-white">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <IndianRupee className="w-5 h-5 text-orange-400" />
                        </div>
                        <div className="flex-1">
                          <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Inhouse Value</p>
                          <p className="text-2xl font-bold text-white">₹{bdeSalesStats.inhouse.value.toLocaleString()}</p>
                          <div className="mt-2 pt-2 border-t border-zinc-800 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                            <div className="text-zinc-400">Adv: ₹{bdeSalesStats.inhouse.advanceValue.toLocaleString()}</div>
                            <div className="text-zinc-400">CN: ₹{bdeSalesStats.inhouse.cnValue.toLocaleString()}</div>
                            <div className="text-zinc-400">Conf: ₹{bdeSalesStats.inhouse.confirmedValue.toLocaleString()}</div>
                            <div className="text-zinc-400">Can: ₹{bdeSalesStats.inhouse.cancelValue.toLocaleString()}</div>
                          </div>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-zinc-900 text-white">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <IndianRupee className="w-5 h-5 text-orange-400" />
                        </div>
                        <div className="flex-1">
                          <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Branch Value</p>
                          <p className="text-2xl font-bold text-white">₹{bdeSalesStats.branch.value.toLocaleString()}</p>
                          <div className="mt-2 pt-2 border-t border-zinc-800 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                            <div className="text-zinc-400">Adv: ₹{bdeSalesStats.branch.advanceValue.toLocaleString()}</div>
                            <div className="text-zinc-400">CN: ₹{bdeSalesStats.branch.cnValue.toLocaleString()}</div>
                            <div className="text-zinc-400">Conf: ₹{bdeSalesStats.branch.confirmedValue.toLocaleString()}</div>
                            <div className="text-zinc-400">Can: ₹{bdeSalesStats.branch.cancelValue.toLocaleString()}</div>
                          </div>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-zinc-900 text-white">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <IndianRupee className="w-5 h-5 text-orange-400" />
                        </div>
                        <div className="flex-1">
                          <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Franchisee Value</p>
                          <p className="text-2xl font-bold text-white">₹{bdeSalesStats.franchisee.value.toLocaleString()}</p>
                          <div className="mt-2 pt-2 border-t border-zinc-800 grid grid-cols-2 gap-1 text-[9px] font-bold uppercase">
                            <div className="text-zinc-400">Adv: ₹{bdeSalesStats.franchisee.advanceValue.toLocaleString()}</div>
                            <div className="text-zinc-400">CN: ₹{bdeSalesStats.franchisee.cnValue.toLocaleString()}</div>
                            <div className="text-zinc-400">Conf: ₹{bdeSalesStats.franchisee.confirmedValue.toLocaleString()}</div>
                            <div className="text-zinc-400">Can: ₹{bdeSalesStats.franchisee.cancelValue.toLocaleString()}</div>
                          </div>
                        </div>
                      </div>
                    </Card>
                  </div>

                  {/* 8 Weeks Performance Highlight */}
                  <Card className={cn("p-8 border-2 shadow-lg transition-all hover:scale-[1.01]", bdeStatus.bg, bdeStatus.border)}>
                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-8 items-center">
                      <div className="space-y-2 text-center lg:text-left lg:border-r lg:border-zinc-200 pr-0 lg:pr-8">
                        <p className={cn("text-sm font-bold uppercase tracking-widest", bdeStatus.color)}>8 Weeks Bravo Point Achievement</p>
                        <div className="flex items-center justify-center lg:justify-start gap-4">
                          <h1 className={cn("text-7xl font-black tracking-tighter", bdeStatus.color)}>
                            {eightWeekAchievementHeader}%
                          </h1>
                          <div className={cn("px-3 py-1 rounded-full text-[10px] font-bold uppercase border-2", bdeStatus.color, bdeStatus.border, "bg-white/50")}>
                            {bdeStatus.label}
                          </div>
                        </div>
                        <p className="text-zinc-500 text-[10px] font-medium">Based on performance over the last 8 weeks</p>
                      </div>

                      <div className="space-y-2 text-center lg:text-left lg:border-r lg:border-zinc-200 pr-0 lg:pr-8">
                        <p className="text-sm font-bold uppercase tracking-widest text-zinc-500">Total Contribution (8 Weeks)</p>
                        <h2 className={cn("text-5xl font-black tracking-tighter", eightWeekProfitHeader >= 0 ? "text-green-600" : "text-red-600")}>
                          ₹{eightWeekProfitHeader.toLocaleString()}
                        </h2>
                        <p className="text-zinc-500 text-[10px] font-medium">Net contribution to company after workstation costs</p>
                      </div>

                      <div className="space-y-4 text-center lg:text-left">
                        <div>
                          <p className="text-xs font-bold uppercase tracking-widest text-zinc-500 mb-1">Current Floor Timing</p>
                          <div className="flex items-center justify-center lg:justify-start gap-2">
                            <Clock className="w-4 h-4 text-zinc-400" />
                            <span className="text-2xl font-black text-zinc-800">
                              {eightWeekAchievementHeader > 90 ? "8:30 Minutes" : eightWeekAchievementHeader >= 70 ? "9:30 Minutes" : "10:30 Minutes"}
                            </span>
                          </div>
                        </div>
                        <div>
                          <p className="text-xs font-bold uppercase tracking-widest text-zinc-500 mb-1">Sunday Requirement</p>
                          <div className="flex items-center justify-center lg:justify-start gap-2">
                            <CalendarIcon className="w-4 h-4 text-zinc-400" />
                            <span className={cn("text-sm font-bold", eightWeekAchievementHeader < 70 ? "text-red-600" : "text-green-600")}>
                              {eightWeekAchievementHeader < 70 ? "2nd/4th Sunday Working Required" : "Standard Offs (Sundays Off)"}
                            </span>
                          </div>
                        </div>
                        <div>
                          <p className="text-xs font-bold uppercase tracking-widest text-zinc-500 mb-1">HR Review Status</p>
                          <div className="flex items-center justify-center lg:justify-start gap-2">
                            <Shield className="w-4 h-4 text-zinc-400" />
                            <span className={cn("text-sm font-bold", eightWeekAchievementHeader < 30 ? "text-red-600" : "text-green-600")}>
                              {eightWeekAchievementHeader < 30 ? (
                                <span className="flex items-center gap-1 animate-pulse">
                                  <AlertCircle className="w-3 h-3" /> Talk to HR
                                </span>
                              ) : "Clear"}
                            </span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </Card>

                  {/* Daily Sales Analysis Chart */}
                  {(() => {
                    const dailySalesDataRaw = filteredSales.reduce((acc: any, sale) => {
                      const date = sale.date || 'Unknown';
                      if (!acc[date]) acc[date] = { date, count: 0, value: 0 };
                      acc[date].count += 1;
                      acc[date].value += (sale.packageValue || 0);
                      return acc;
                    }, {});

                    const dailySalesData = Object.values(dailySalesDataRaw)
                      .sort((a: any, b: any) => new Date(a.date as string).getTime() - new Date(b.date as string).getTime());

                    return (
                      <Card className="p-6">
                        <h3 className="text-sm font-bold text-zinc-900 uppercase mb-6 flex items-center gap-2">
                          <CalendarIcon className="w-4 h-4 text-orange-600" />
                          Daily Sales Analysis (Value & Count)
                        </h3>
                        <div className="h-[500px] w-full">
                          <ResponsiveContainer width="100%" height="100%">
                            <BarChart data={dailySalesData} margin={{ top: 130, right: 30, left: 20, bottom: 60 }}>
                              <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                              <XAxis 
                                dataKey="date" 
                                fontSize={10} 
                                fontWeight="bold"
                                angle={-45}
                                textAnchor="end"
                                interval={0}
                                tickFormatter={(val) => val === 'Unknown' ? val : new Date(val).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })}
                              />
                              <YAxis fontSize={10} tickFormatter={(val) => `₹${(val/100000).toFixed(2)}L`} />
                              <Tooltip 
                                contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                                formatter={(value: number, name: string) => [name === 'value' ? `₹${value.toLocaleString()}` : value, name === 'value' ? 'Amount' : 'Sales Count']}
                              />
                              <Bar dataKey="value" name="Amount" fill="#ea580c" radius={[4, 4, 0, 0]}>
                                <LabelList 
                                  dataKey="value" 
                                  position="top" 
                                  content={(props: any) => {
                                    const { x, y, width, value, index } = props;
                                    const dataPoint = dailySalesData[index] as any;
                                    const count = dataPoint?.count || 0;
                                    return (
                                      <text 
                                        x={x + width / 2} 
                                        y={y - 45} 
                                        fill="#ea580c" 
                                        textAnchor="middle" 
                                        fontSize={14} 
                                        fontWeight="black"
                                        transform={`rotate(-90, ${x + width / 2}, ${y - 45})`}
                                      >
                                        ₹{(value / 100000).toFixed(2)}L ({count})
                                      </text>
                                    );
                                  }}
                                />
                              </Bar>
                            </BarChart>
                          </ResponsiveContainer>
                        </div>
                      </Card>
                    );
                  })()}

                  <div className="bg-zinc-50 p-4 rounded-xl border border-zinc-200 flex flex-col md:flex-row md:items-center justify-between gap-2">
                    <div>
                      <p className="text-xs font-bold text-zinc-500 uppercase">Data sorted for:</p>
                      <p className="text-lg font-bold text-orange-600">{timeframeLabel}</p>
                    </div>
                    <div className="text-right">
                      <p className="text-xs font-bold text-zinc-500 uppercase">Date Range:</p>
                      <p className="text-sm font-medium text-zinc-700">{dateRangeLabel}</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    <Card 
                      className="p-6 border-l-4 border-orange-500 cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setBdeOverviewSummaryModal({
                        isOpen: true,
                        type: 'sales',
                        title: 'Total Sales Summary',
                        data: { totalSales, filteredSales }
                      })}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-xs font-bold text-zinc-500 uppercase">Total Sales</p>
                        <TrendingUp className="w-4 h-4 text-orange-500" />
                      </div>
                      <p className="text-2xl font-bold">₹{totalSales.toLocaleString()}</p>
                      <p className="text-[10px] text-zinc-400 mt-1">{timeframeLabel}</p>
                    </Card>
                    <Card 
                      className="p-6 border-l-4 border-blue-500 cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setBdeOverviewSummaryModal({
                        isOpen: true,
                        type: 'margin',
                        title: 'Approx Margin Summary',
                        data: { approxMargin, totalSales }
                      })}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-xs font-bold text-zinc-500 uppercase">Approx margin (12%)</p>
                        <PieChart className="w-4 h-4 text-blue-500" />
                      </div>
                      <p className="text-2xl font-bold">₹{approxMargin.toLocaleString()}</p>
                      <p className="text-[10px] text-zinc-400 mt-1">12% of total sales</p>
                    </Card>
                    <Card 
                      className="p-6 border-l-4 border-red-500 cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setBdeOverviewSummaryModal({
                        isOpen: true,
                        type: 'expenses',
                        title: 'Expenses Summary',
                        data: { totalExpenses, bdeAgents, timeframe: bdeOverviewTimeframe }
                      })}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-xs font-bold text-zinc-500 uppercase">Expenses</p>
                        <DollarSign className="w-4 h-4 text-red-500" />
                      </div>
                      <p className="text-2xl font-bold">₹{totalExpenses.toLocaleString()}</p>
                      <p className="text-[10px] text-zinc-400 mt-1">Salary + Workstation Cost</p>
                    </Card>
                    <Card 
                      className="p-6 border-l-4 border-emerald-500 cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setBdeOverviewSummaryModal({
                        isOpen: true,
                        type: 'profit',
                        title: 'Net Profit Summary',
                        data: { profitability, approxMargin, totalExpenses }
                      })}
                    >
                      <div className="flex items-center justify-between mb-2">
                        <p className="text-xs font-bold text-zinc-500 uppercase">Net Profit</p>
                        <Target className="w-4 h-4 text-emerald-500" />
                      </div>
                      <p className={cn("text-2xl font-bold", profitability >= 0 ? "text-emerald-600" : "text-red-600")}>
                        ₹{profitability.toLocaleString()}
                      </p>
                      <p className="text-[10px] text-zinc-400 mt-1">Approx Margin - Expenses</p>
                    </Card>
                  </div>

                  {selectedBDEForOverview === 'All' && (
                    <Card className="p-6 overflow-x-auto">
                      <div className="flex items-center justify-between mb-6">
                        <h3 className="text-sm font-bold text-zinc-900 uppercase flex items-center gap-2">
                          <Users className="w-4 h-4 text-orange-600" />
                          BDE Comparison Table
                        </h3>
                        <button
                          onClick={() => downloadBDEComparison(bdeComparisonData, timeframeLabel)}
                          className="flex items-center gap-2 px-3 py-1.5 bg-emerald-600 text-white rounded-lg text-xs font-bold hover:bg-emerald-700 transition-colors"
                        >
                          <Download className="w-3 h-3" />
                          Export Excel
                        </button>
                      </div>
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-zinc-50 border-b border-zinc-200">
                            <th className="p-3 text-[10px] font-bold text-zinc-500 uppercase tracking-wider">BDE Name</th>
                            <th className="p-3 text-[10px] font-bold text-zinc-500 uppercase tracking-wider text-right">Total Sales</th>
                            <th className="p-3 text-[10px] font-bold text-zinc-500 uppercase tracking-wider text-right">Approx Margin (12%)</th>
                            <th className="p-3 text-[10px] font-bold text-zinc-500 uppercase tracking-wider text-right">Expenses</th>
                            <th className="p-3 text-[10px] font-bold text-zinc-500 uppercase tracking-wider text-right">Net Profit</th>
                            <th className="p-3 text-[10px] font-bold text-zinc-500 uppercase tracking-wider text-right">Sales/Salary Ratio</th>
                          </tr>
                        </thead>
                        <tbody>
                          {bdeComparisonData.sort((a, b) => b.totalSales - a.totalSales).map((data, idx) => (
                            <tr key={idx} className="border-b border-zinc-100 hover:bg-zinc-50 transition-colors">
                              <td className="p-3 text-xs font-bold text-zinc-700">{data.name}</td>
                              <td className="p-3 text-xs font-bold text-zinc-900 text-right">₹{data.totalSales.toLocaleString()}</td>
                              <td className="p-3 text-xs font-bold text-blue-600 text-right">₹{data.approxMargin.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
                              <td className="p-3 text-xs font-bold text-red-600 text-right">₹{data.expenses.toLocaleString()}</td>
                              <td className={cn("p-3 text-xs font-bold text-right", data.netProfit >= 0 ? "text-emerald-600" : "text-red-600")}>
                                ₹{data.netProfit.toLocaleString(undefined, { minimumFractionDigits: 2, maximumFractionDigits: 2 })}
                              </td>
                              <td className="p-3 text-xs font-bold text-zinc-900 text-right">{data.salesSalaryRatio.toFixed(2)}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </Card>
                  )}

                  <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                    <Card className="p-6 lg:col-span-1">
                      <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
                        <Target className="w-5 h-5 text-orange-600" />
                        Target Achievement
                      </h3>
                      <div className="flex flex-col items-center justify-center py-8">
                        <div className="relative w-40 h-40">
                          <svg className="w-full h-full" viewBox="0 0 36 36">
                            <path
                              className="text-zinc-100"
                              strokeDasharray="100, 100"
                              d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="3"
                            />
                            <path
                              className="text-orange-500"
                              strokeDasharray={`${Math.min(100, achievement)}, 100`}
                              d="M18 2.0845 a 15.9155 15.9155 0 0 1 0 31.831 a 15.9155 15.9155 0 0 1 0 -31.831"
                              fill="none"
                              stroke="currentColor"
                              strokeWidth="3"
                              strokeLinecap="round"
                            />
                          </svg>
                          <div className="absolute inset-0 flex flex-col items-center justify-center">
                            <span className="text-3xl font-bold">{achievement}%</span>
                            <span className="text-[10px] text-zinc-500 uppercase font-bold">Achieved</span>
                          </div>
                        </div>
                        <div className="mt-6 w-full space-y-2">
                          <div className="flex justify-between text-sm">
                            <span className="text-zinc-500">BDE Target (75%)</span>
                            <span className="font-bold">₹{totalTarget.toLocaleString()}</span>
                          </div>
                          <div className="flex justify-between text-sm">
                            <span className="text-zinc-500">Actual Sales</span>
                            <span className="font-bold">₹{totalSales.toLocaleString()}</span>
                          </div>
                        </div>
                      </div>
                    </Card>

                    <Card className="p-6 lg:col-span-2">
                      <h3 className="text-lg font-bold mb-4 flex items-center gap-2">
                        <Users className="w-5 h-5 text-orange-600" />
                        Agent Performance
                      </h3>
                      <div className="overflow-x-auto">
                        <table className="w-full text-left border-collapse">
                          <thead>
                            <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase tracking-wider">
                              <th className="px-4 py-3 border-b">Agent Name</th>
                              <th className="px-4 py-3 border-b text-right">Target</th>
                              <th className="px-4 py-3 border-b text-right">Sales</th>
                              <th className="px-4 py-3 border-b text-center">100%</th>
                              <th className="px-4 py-3 border-b text-center">70%</th>
                              <th className="px-4 py-3 border-b text-center">50%</th>
                              <th className="px-4 py-3 border-b text-center">30%</th>
                              <th className="px-4 py-3 border-b text-center">0%</th>
                              <th className="px-4 py-3 border-b">Assoc BDEs</th>
                              <th className="px-4 py-3 border-b text-right">Net Profit</th>
                              <th className="px-4 py-3 border-b text-center">Achievement</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-zinc-100">
                            {bdeAgents.map(agent => {
                              const agentSalesData = filteredSales.filter(s => 
                                s.agent?.toLowerCase().trim() === agent.name.toLowerCase().trim() ||
                                (agent.email && s.agentEmail?.toLowerCase() === agent.email.toLowerCase()) ||
                                s.associateBde?.toLowerCase().trim() === agent.name.toLowerCase().trim() ||
                                (agent.email && s.associateBdeEmail?.toLowerCase() === agent.email.toLowerCase())
                              );
                              const agentTotalSales = agentSalesData.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'packageValue'), 0);
                              const agentTotalMargin = agentSalesData.reduce((sum, s) => sum + getAgentCreditForSale(s, agent, 'totalMargin'), 0);
                              
                              const monthlyTotalExpense = agent.salary * 3;
                              let agentAdjustedExpense = 0;
                              if (bdeOverviewTimeframe === 'Daily') agentAdjustedExpense = monthlyTotalExpense / 30;
                              else if (bdeOverviewTimeframe === 'Weekly') agentAdjustedExpense = monthlyTotalExpense / 4;
                              else if (bdeOverviewTimeframe === '4-Weekly') agentAdjustedExpense = monthlyTotalExpense;
                              else if (bdeOverviewTimeframe === '8-Weekly') agentAdjustedExpense = monthlyTotalExpense * 2;
                              else if (bdeOverviewTimeframe === 'Monthly') agentAdjustedExpense = monthlyTotalExpense;

                              const agentNetProfit = (agentTotalSales * 0.12) - agentAdjustedExpense;
                              
                              let agentTarget = 0;
                              if (bdeOverviewTimeframe === 'Daily') agentTarget = agent.target / 24;
                              else if (bdeOverviewTimeframe === 'Weekly') {
                                const weeklyBaseTarget = agent.target / 4;
                                const wkIndex = allSortedWeeks.findIndex(w => w.weekName === currentWeek?.weekName);
                                const joiningWkIndex = allSortedWeeks.findIndex(w => w.weekName === agent.joiningWeek);
                                let percentageVal = 100;
                                if (joiningWkIndex !== -1 && wkIndex >= joiningWkIndex) {
                                  const tenureWeek = wkIndex - joiningWkIndex + 1;
                                  if (tenureWeek <= 6) percentageVal = (agent as any)[`week${tenureWeek}Target`] || 0;
                                }
                                agentTarget = (weeklyBaseTarget * percentageVal) / 100;
                              } else if (bdeOverviewTimeframe === '4-Weekly') {
                                agentTarget = agent.target;
                              } else if (bdeOverviewTimeframe === '8-Weekly') {
                                agentTarget = agent.target * 2;
                              } else if (bdeOverviewTimeframe === 'Monthly') {
                                agentTarget = agent.target;
                              }

                              const agentAchievement = agentTarget > 0 ? Math.round((agentTotalSales / agentTarget) * 100) : 0;
                              const sales100 = agentSalesData.filter(s => s.agentPercentage === 100).length;
                              const sales70 = agentSalesData.filter(s => s.agentPercentage === 70).length;
                              const sales50 = agentSalesData.filter(s => s.agentPercentage === 50).length;
                              const sales30 = agentSalesData.filter(s => s.agentPercentage === 30).length;
                              const sales0 = agentSalesData.filter(s => s.agentPercentage === 0).length;
                              
                              const assocBDEs = Array.from(new Set(agentSalesData.filter(s => s.associateBde).map(s => s.associateBde))).join(', ');

                              return (
                                <tr key={agent.id} className="hover:bg-zinc-50/50 transition-colors">
                                  <td 
                                    className="px-4 py-3 cursor-pointer group"
                                    onClick={() => setBdeOverviewSummaryModal({
                                      isOpen: true,
                                      type: 'agent',
                                      title: `Agent Summary: ${agent.name}`,
                                      data: { agent, agentTotalSales, agentTotalMargin, agentNetProfit, agentTarget, agentAchievement, agentSalesData }
                                    })}
                                  >
                                    <p className="text-sm font-bold text-zinc-900 group-hover:text-orange-600 transition-colors">{agent.name}</p>
                                    <p className="text-[10px] text-zinc-400">{agent.employeeCode}</p>
                                  </td>
                                  <td className="px-4 py-3 text-sm font-medium text-right text-zinc-500">₹{agentTarget.toLocaleString()}</td>
                                  <td className="px-4 py-3 text-sm font-medium text-right">₹{agentTotalSales.toLocaleString()}</td>
                                  <td className="px-4 py-3 text-sm text-center font-medium text-zinc-600">{sales100 || '-'}</td>
                                  <td className="px-4 py-3 text-sm text-center font-medium text-zinc-600">{sales70 || '-'}</td>
                                  <td className="px-4 py-3 text-sm text-center font-medium text-zinc-600">{sales50 || '-'}</td>
                                  <td className="px-4 py-3 text-sm text-center font-medium text-zinc-600">{sales30 || '-'}</td>
                                  <td className="px-4 py-3 text-sm text-center font-medium text-zinc-600">{sales0 || '-'}</td>
                                  <td className="px-4 py-3 text-sm text-zinc-600 italic">{assocBDEs || '-'}</td>
                                  <td className={cn("px-4 py-3 text-sm font-medium text-right", agentNetProfit >= 0 ? "text-emerald-600" : "text-red-600")}>
                                    ₹{agentNetProfit.toLocaleString()}
                                  </td>
                                  <td className="px-4 py-3 text-center">
                                    <div className="flex items-center justify-center gap-2">
                                      <div className="w-16 bg-zinc-100 rounded-full h-1.5 overflow-hidden">
                                        <div 
                                          className={cn("h-full rounded-full", agentAchievement >= 100 ? "bg-emerald-500" : agentAchievement >= 70 ? "bg-orange-500" : "bg-red-500")}
                                          style={{ width: `${Math.min(100, agentAchievement)}%` }}
                                        />
                                      </div>
                                      <span className={cn("text-xs font-bold", agentAchievement >= 100 ? "text-emerald-600" : agentAchievement >= 70 ? "text-orange-600" : "text-red-600")}>
                                        {agentAchievement}%
                                      </span>
                                    </div>
                                  </td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </Card>
                  </div>

                  <Card className="p-6">
                    <h3 className="text-lg font-bold mb-6 flex items-center gap-2">
                      <BarChart3 className="w-5 h-5 text-orange-600" />
                      Multi-Timeframe Performance & Breakeven
                    </h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 xl:grid-cols-5 gap-4">
                      {[
                        { label: 'Weekly', weeks: 1 },
                        { label: '4-Weekly', weeks: 4 },
                        { label: '8-Weekly', weeks: 8 },
                        { label: '12-Weekly', weeks: 12 },
                        { label: '24-Weekly', weeks: 24 }
                      ].map((tf) => {
                        const tfWeeks = allSortedWeeks.slice(-tf.weeks);
                        const tfWeekNames = tfWeeks.map(w => w.weekName);
                        const tfSales = bdeSales.filter(s => tfWeekNames.includes(s.week));
                        const tfTotalSales = tfSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);
                        const tfApproxMargin = tfTotalSales * 0.12;
                        
                        const tfTotalExpenses = bdeAgents.reduce((sum, emp) => {
                          const monthlyTotalExpense = emp.salary * 3;
                          return sum + (monthlyTotalExpense / 4) * tf.weeks;
                        }, 0);

                        const tfActualTarget = bdeAgents.reduce((sum, emp) => sum + (emp.target / 4) * tf.weeks, 0);
                        const tfBdeTarget = tfActualTarget * 0.75;
                        const tfBreakevenSales = tfTotalExpenses / 0.12;
                        const tfProfit = tfApproxMargin - tfTotalExpenses;
                        const tfAchievement = tfBdeTarget > 0 ? Math.round((tfTotalSales / tfBdeTarget) * 100) : 0;

                        return (
                          <Card key={tf.label} className="p-5 border-zinc-200 hover:shadow-lg transition-all duration-300 group">
                            <div className="flex items-center justify-between mb-4">
                              <h4 className="font-bold text-zinc-900 group-hover:text-orange-600 transition-colors">{tf.label}</h4>
                              <div className="px-2 py-0.5 bg-orange-50 text-orange-600 text-[10px] font-bold rounded-full uppercase">
                                {tf.weeks} {tf.weeks === 1 ? 'Week' : 'Weeks'}
                              </div>
                            </div>

                            <div className="flex flex-col items-center justify-center mb-6">
                              <div className="relative w-28 h-28">
                                <svg className="w-full h-full -rotate-90" viewBox="0 0 36 36">
                                  <circle
                                    cx="18" cy="18" r="15.9155"
                                    fill="none" stroke="#f4f4f5" strokeWidth="3"
                                  />
                                  <circle
                                    cx="18" cy="18" r="15.9155"
                                    fill="none" stroke="currentColor" strokeWidth="3"
                                    strokeDasharray={`${Math.min(100, tfAchievement)}, 100`}
                                    strokeLinecap="round"
                                    className={cn(
                                      "transition-all duration-1000",
                                      tfAchievement >= 100 ? "text-emerald-500" : tfAchievement >= 70 ? "text-orange-500" : "text-red-500"
                                    )}
                                  />
                                </svg>
                                <div className="absolute inset-0 flex flex-col items-center justify-center">
                                  <span className="text-xl font-bold">{tfAchievement}%</span>
                                  <span className="text-[8px] text-zinc-400 uppercase font-bold">Achieved</span>
                                </div>
                              </div>
                            </div>

                            <div className="space-y-2.5">
                              <div className="flex justify-between items-center">
                                <span className="text-[10px] font-bold text-zinc-500 uppercase">Sales</span>
                                <span className="text-xs font-bold text-zinc-900">₹{tfTotalSales.toLocaleString()}</span>
                              </div>
                              <div className="flex justify-between items-center">
                                <span className="text-[10px] font-bold text-zinc-500 uppercase">Margin - 12% of sales</span>
                                <span className="text-xs font-bold text-blue-600">₹{tfApproxMargin.toLocaleString()}</span>
                              </div>
                              <div className="flex justify-between items-center">
                                <span className="text-[10px] font-bold text-zinc-500 uppercase">Expenses - Total of workstation cost</span>
                                <span className="text-xs font-bold text-red-600">₹{tfTotalExpenses.toLocaleString()}</span>
                              </div>
                              <div className="flex justify-between items-center pt-1 border-t border-zinc-50">
                                <span className="text-[10px] font-bold text-zinc-500 uppercase">New profit - Margin - Workstation cost</span>
                                <span className={cn("text-xs font-bold", tfProfit >= 0 ? "text-emerald-600" : "text-red-600")}>
                                  ₹{tfProfit.toLocaleString()}
                                </span>
                              </div>
                              <div className="flex justify-between items-center">
                                <span className="text-[10px] font-bold text-zinc-500 uppercase">BDE target = Total target *75%</span>
                                <span className="text-xs font-bold text-orange-600">₹{tfBdeTarget.toLocaleString()}</span>
                              </div>
                              <div className="flex justify-between items-center p-2 bg-zinc-50 rounded-lg mt-2">
                                <span className="text-[9px] font-bold text-zinc-500 uppercase">Breakeven sales</span>
                                <span className="text-xs font-bold text-zinc-900">₹{Math.round(tfBreakevenSales).toLocaleString()}</span>
                              </div>
                            </div>
                          </Card>
                        );
                      })}
                    </div>
                  </Card>

                  {/* Compressed Sales Table */}
                  <Card className="p-6">
                    <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 mb-6">
                      <h3 className="text-lg font-bold flex items-center gap-2">
                        <TrendingUp className="w-5 h-5 text-orange-600" />
                        Sales Details ({timeframeLabel})
                      </h3>
                      <div className="flex flex-wrap items-center gap-3">
                        <select 
                          className="px-3 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-xs font-bold uppercase"
                          value={bdeOverviewWorkPendingFilter}
                          onChange={(e) => setBdeOverviewWorkPendingFilter(e.target.value)}
                        >
                          <option value="All">Work Pending: All</option>
                          <option value="Issue">Issue</option>
                          <option value="Done">Done</option>
                        </select>
                        <select 
                          className="px-3 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-xs font-bold uppercase"
                          value={bdeOverviewStatusFilter}
                          onChange={(e) => setBdeOverviewStatusFilter(e.target.value)}
                        >
                          <option value="All">Status: All</option>
                          <option value="Advance">Advance</option>
                          <option value="Credit Note">Credit Note</option>
                          <option value="Done">Done</option>
                          <option value="Confirmed">Confirmed</option>
                          <option value="Cancel">Cancel</option>
                        </select>
                        <div className="relative w-full md:w-64">
                          <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-400" />
                          <input 
                            type="text"
                            placeholder="Search sales..."
                            className="w-full pl-10 pr-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-sm"
                            value={bdeOverviewSalesSearch}
                            onChange={(e) => setBdeOverviewSalesSearch(e.target.value)}
                          />
                        </div>
                      </div>
                    </div>
                    
                    <div className="overflow-x-auto">
                      <table className="w-full text-left border-collapse">
                        <thead>
                          <tr className="bg-zinc-50 text-[10px] font-bold text-zinc-500 uppercase tracking-wider">
                            <th className="px-4 py-3 border-b">Log</th>
                          <th className="px-4 py-3 border-b">Date</th>
                            <th className="px-4 py-3 border-b">Trip ID</th>
                            <th className="px-4 py-3 border-b">Guest Name</th>
                            <th className="px-4 py-3 border-b">Agent</th>
                            <th className="px-4 py-3 border-b">Assoc BDE</th>
                            <th className="px-4 py-3 border-b text-center">% Agent</th>
                            <th className="px-4 py-3 border-b text-center">% Assoc BDE</th>
                            <th className="px-4 py-3 border-b">Destination</th>
                            <th className="px-4 py-3 border-b text-right">Value</th>
                            <th className="px-4 py-3 border-b">P - Hotel</th>
                            <th className="px-4 py-3 border-b">P Flight</th>
                            <th className="px-4 py-3 border-b">Land</th>
                            <th className="px-4 py-3 border-b">HFL Issue</th>
                            <th className="px-4 py-3 border-b">Work Pending</th>
                            <th className="px-4 py-3 border-b text-center">Status</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-zinc-100">
                          {(() => {
                            const bdeOverviewFilteredSales = filteredSales
                              .filter(sale => {
                                // Search filter
                                const matchesSearch = !bdeOverviewSalesSearch || (() => {
                                  const search = bdeOverviewSalesSearch.toLowerCase();
                                  return (
                                    sale.guestName?.toLowerCase().includes(search) ||
                                    sale.tripId?.toLowerCase().includes(search) ||
                                    sale.destination?.toLowerCase().includes(search) ||
                                    sale.agent?.toLowerCase().includes(search)
                                  );
                                })();

                                // Work Pending filter
                                const matchesWorkPending = bdeOverviewWorkPendingFilter === 'All' || 
                                  sale.workPending === bdeOverviewWorkPendingFilter;

                                // Status filter
                                const matchesStatus = bdeOverviewStatusFilter === 'All' || 
                                  sale.advanceCN === bdeOverviewStatusFilter;

                                return matchesSearch && matchesWorkPending && matchesStatus;
                              })
                              .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

                            if (bdeOverviewFilteredSales.length === 0) {
                              return (
                                <tr>
                                  <td colSpan={12} className="px-4 py-8 text-center text-zinc-400 text-sm italic">
                                    No sales found for this period
                                  </td>
                                </tr>
                              );
                            }

                            return bdeOverviewFilteredSales.map(sale => (
                              <tr key={sale.id} className="hover:bg-zinc-50/50 transition-colors">
                                <td className="px-4 py-3 border-b text-center">
                                  <button 
                                    onClick={() => handleOpenRemarks(sale)}
                                    className={cn(
                                      "p-1.5 rounded-lg transition-all flex items-center gap-1 mx-auto",
                                      isRemarkUnread(sale)
                                        ? "bg-orange-600 text-white shadow-lg shadow-orange-600/20 animate-pulse" 
                                        : "bg-zinc-100 text-zinc-400 hover:bg-zinc-200 hover:text-zinc-600"
                                    )}
                                  >
                                    <MessageSquare className="w-3.5 h-3.5" />
                                    <span className="text-[10px] font-black">{sale.remarksCount || 0}</span>
                                  </button>
                                </td>
                                <td className="px-4 py-3 text-xs text-zinc-600 whitespace-nowrap">{sale.date}</td>
                                <td className="px-4 py-3 text-xs font-bold text-zinc-900">{sale.tripId}</td>
                                <td className="px-4 py-3 text-xs font-medium text-zinc-700">{sale.guestName}</td>
                                <td className="px-4 py-3 text-xs text-zinc-600">{sale.agent}</td>
                                <td className="px-4 py-3 text-xs text-zinc-600">{sale.associateBde}</td>
                                <td className="px-4 py-3 text-xs text-center font-medium">{sale.agentPercentage}%</td>
                                <td className="px-4 py-3 text-xs text-center font-medium">{sale.assocBdePercentage}%</td>
                                <td className="px-4 py-3 text-xs text-zinc-600">{sale.destination}</td>
                                <td className="px-4 py-3 text-xs font-bold text-right">₹{sale.packageValue?.toLocaleString()}</td>
                                <td className={cn("px-4 py-3 text-xs", sale.pHotel === 'Issue' ? "text-red-600 font-bold" : "text-zinc-600")}>{sale.pHotel || '-'}</td>
                                <td className={cn("px-4 py-3 text-xs", sale.pFlight === 'Issue' ? "text-red-600 font-bold" : "text-zinc-600")}>{sale.pFlight || '-'}</td>
                                <td className={cn("px-4 py-3 text-xs", sale.land === 'Issue' ? "text-red-600 font-bold" : "text-zinc-600")}>{sale.land || '-'}</td>
                                <td className={cn("px-4 py-3 text-xs", sale.hflIssue === 'Issue' ? "text-red-600 font-bold" : "text-zinc-600")}>{sale.hflIssue || '-'}</td>
                                <td className={cn("px-4 py-3 text-xs", sale.workPending === 'Issue' ? "text-red-600 font-bold" : "text-zinc-600")}>{sale.workPending || '-'}</td>
                                <td className="px-4 py-3 text-center">
                                  <Badge className={cn(
                                    "text-[9px] px-1.5 py-0.5",
                                    sale.advanceCN === 'Advance' ? "bg-blue-100 text-blue-700" :
                                    sale.advanceCN === 'Credit Note' ? "bg-purple-100 text-purple-700" :
                                    sale.advanceCN === 'Done' || sale.advanceCN === 'Confirmed' ? "bg-green-100 text-green-700" :
                                    sale.advanceCN === 'Cancel' ? "bg-red-100 text-red-700" :
                                    "bg-zinc-100 text-zinc-700"
                                  )}>
                                    {sale.advanceCN || 'Pending'}
                                  </Badge>
                                </td>
                              </tr>
                            ));
                          })()}
                        </tbody>
                      </table>
                    </div>
                  </Card>

                  {/* BDE Overview Summary Modal */}
                  <Modal
                    isOpen={bdeOverviewSummaryModal.isOpen}
                    onClose={() => setBdeOverviewSummaryModal({ ...bdeOverviewSummaryModal, isOpen: false })}
                    title={bdeOverviewSummaryModal.title}
                  >
                    <div className="space-y-6">
                      {bdeOverviewSummaryModal.type === 'sales' && (
                        <div className="space-y-4">
                          <div className="grid grid-cols-2 gap-4">
                            <div className="p-4 bg-orange-50 rounded-xl border border-orange-100">
                              <p className="text-xs font-bold text-orange-600 uppercase">Total Sales</p>
                              <p className="text-xl font-bold">₹{bdeOverviewSummaryModal.data?.totalSales.toLocaleString()}</p>
                            </div>
                            <div className="p-4 bg-zinc-50 rounded-xl border border-zinc-100">
                              <p className="text-xs font-bold text-zinc-600 uppercase">Transactions</p>
                              <p className="text-xl font-bold">{bdeOverviewSummaryModal.data?.filteredSales.length}</p>
                            </div>
                          </div>
                          <div className="max-h-60 overflow-y-auto border border-zinc-100 rounded-xl">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-zinc-50 sticky top-0">
                                <tr>
                                  <th className="px-3 py-2 border-b font-bold uppercase">Date</th>
                                  <th className="px-3 py-2 border-b font-bold uppercase">Trip ID</th>
                                  <th className="px-3 py-2 border-b font-bold uppercase">Guest Name</th>
                                  <th className="px-3 py-2 border-b font-bold uppercase">Agent</th>
                                  <th className="px-3 py-2 border-b font-bold uppercase text-right">Value</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-zinc-50">
                                {[...(bdeOverviewSummaryModal.data?.filteredSales || [])]
                                  .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
                                  .map((s: any, i: number) => (
                                    <tr key={i}>
                                      <td className="px-3 py-2 whitespace-nowrap">{s.date}</td>
                                      <td className="px-3 py-2 font-mono text-orange-600">{s.tripId}</td>
                                      <td className="px-3 py-2">{s.guestName}</td>
                                      <td className="px-3 py-2">{s.agent}</td>
                                      <td className="px-3 py-2 text-right font-medium">₹{s.packageValue?.toLocaleString()}</td>
                                    </tr>
                                  ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {bdeOverviewSummaryModal.type === 'margin' && (
                        <div className="space-y-4">
                          <div className="p-6 bg-blue-50 rounded-2xl border border-blue-100 text-center">
                            <p className="text-sm font-bold text-blue-600 uppercase mb-1">Calculated Approx Margin</p>
                            <p className="text-3xl font-bold text-blue-700">₹{bdeOverviewSummaryModal.data?.approxMargin.toLocaleString()}</p>
                            <p className="text-xs text-blue-500 mt-2 italic">Based on 12% of Total Sales (₹{bdeOverviewSummaryModal.data?.totalSales.toLocaleString()})</p>
                          </div>
                          <div className="p-4 bg-zinc-50 rounded-xl border border-zinc-100">
                            <h4 className="text-xs font-bold text-zinc-500 uppercase mb-3">Margin Breakdown</h4>
                            <div className="space-y-2">
                              <div className="flex justify-between text-sm">
                                <span>Total Sales</span>
                                <span className="font-medium">₹{bdeOverviewSummaryModal.data?.totalSales.toLocaleString()}</span>
                              </div>
                              <div className="flex justify-between text-sm">
                                <span>Margin Percentage</span>
                                <span className="font-medium">12%</span>
                              </div>
                              <div className="border-t border-zinc-200 pt-2 flex justify-between font-bold">
                                <span>Total Approx Margin</span>
                                <span className="text-blue-600">₹{bdeOverviewSummaryModal.data?.approxMargin.toLocaleString()}</span>
                              </div>
                            </div>
                          </div>
                        </div>
                      )}

                      {bdeOverviewSummaryModal.type === 'expenses' && (
                        <div className="space-y-4">
                          <div className="p-6 bg-red-50 rounded-2xl border border-red-100 text-center">
                            <p className="text-sm font-bold text-red-600 uppercase mb-1">Total Expenses</p>
                            <p className="text-3xl font-bold text-red-700">₹{bdeOverviewSummaryModal.data?.totalExpenses.toLocaleString()}</p>
                            <p className="text-xs text-red-500 mt-2 italic">Sum of 100% Workstation Costs (Salary * 3)</p>
                          </div>
                          <div className="max-h-60 overflow-y-auto border border-zinc-100 rounded-xl">
                            <table className="w-full text-left text-xs">
                              <thead className="bg-zinc-50 sticky top-0">
                                <tr>
                                  <th className="px-3 py-2 border-b font-bold uppercase">Agent</th>
                                  <th className="px-3 py-2 border-b font-bold uppercase text-right">Total Expense</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-zinc-50">
                                {bdeOverviewSummaryModal.data?.bdeAgents.map((agent: any) => {
                                  const monthlyTotalExpense = agent.salary * 3;
                                  let adjustedExpense = 0;
                                  const tf = bdeOverviewSummaryModal.data?.timeframe;
                                  if (tf === 'Daily') adjustedExpense = monthlyTotalExpense / 30;
                                  else if (tf === 'Weekly') adjustedExpense = monthlyTotalExpense / 4;
                                  else if (tf === '4-Weekly') adjustedExpense = monthlyTotalExpense;
                                  else if (tf === '8-Weekly') adjustedExpense = monthlyTotalExpense * 2;
                                  else if (tf === 'Monthly') adjustedExpense = monthlyTotalExpense;

                                  return (
                                    <tr key={agent.id}>
                                      <td className="px-3 py-2 font-medium">{agent.name}</td>
                                      <td className="px-3 py-2 text-right">₹{adjustedExpense.toLocaleString()}</td>
                                    </tr>
                                  );
                                })}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      {bdeOverviewSummaryModal.type === 'profit' && (
                        <div className="space-y-4">
                          <div className={cn(
                            "p-6 rounded-2xl border text-center",
                            bdeOverviewSummaryModal.data?.profitability >= 0 
                              ? "bg-emerald-50 border-emerald-100" 
                              : "bg-red-50 border-red-100"
                          )}>
                            <p className={cn(
                              "text-sm font-bold uppercase mb-1",
                              bdeOverviewSummaryModal.data?.profitability >= 0 ? "text-emerald-600" : "text-red-600"
                            )}>Net Profitability</p>
                            <p className={cn(
                              "text-3xl font-bold",
                              bdeOverviewSummaryModal.data?.profitability >= 0 ? "text-emerald-700" : "text-red-700"
                            )}>₹{bdeOverviewSummaryModal.data?.profitability.toLocaleString()}</p>
                          </div>
                          <div className="p-4 bg-zinc-50 rounded-xl border border-zinc-100 space-y-3">
                            <div className="flex justify-between text-sm">
                              <span className="text-zinc-500">Approx Margin (+)</span>
                              <span className="font-bold text-blue-600">₹{bdeOverviewSummaryModal.data?.approxMargin.toLocaleString()}</span>
                            </div>
                            <div className="flex justify-between text-sm">
                              <span className="text-zinc-500">Total Expenses (-)</span>
                              <span className="font-bold text-red-600">₹{bdeOverviewSummaryModal.data?.totalExpenses.toLocaleString()}</span>
                            </div>
                            <div className="border-t border-zinc-200 pt-2 flex justify-between font-bold text-lg">
                              <span>Net Result</span>
                              <span className={bdeOverviewSummaryModal.data?.profitability >= 0 ? "text-emerald-600" : "text-red-600"}>
                                ₹{bdeOverviewSummaryModal.data?.profitability.toLocaleString()}
                              </span>
                            </div>
                          </div>
                        </div>
                      )}

                      {bdeOverviewSummaryModal.type === 'agent' && (
                        <div className="space-y-4">
                          <div className="flex items-center gap-4 p-4 bg-zinc-50 rounded-2xl border border-zinc-100">
                            <div className="w-12 h-12 bg-orange-100 rounded-full flex items-center justify-center text-orange-600 font-bold text-xl">
                              {bdeOverviewSummaryModal.data?.agent.name.charAt(0)}
                            </div>
                            <div>
                              <p className="text-lg font-bold">{bdeOverviewSummaryModal.data?.agent.name}</p>
                              <p className="text-xs text-zinc-500">{bdeOverviewSummaryModal.data?.agent.employeeCode} • {bdeOverviewSummaryModal.data?.agent.position}</p>
                            </div>
                          </div>
                          <div className="grid grid-cols-2 gap-3">
                            <div className="p-3 bg-zinc-50 rounded-xl border border-zinc-100">
                              <p className="text-[10px] font-bold text-zinc-500 uppercase">Target</p>
                              <p className="text-sm font-bold">₹{bdeOverviewSummaryModal.data?.agentTarget.toLocaleString()}</p>
                            </div>
                            <div className="p-3 bg-zinc-50 rounded-xl border border-zinc-100">
                              <p className="text-[10px] font-bold text-zinc-500 uppercase">Sales</p>
                              <p className="text-sm font-bold">₹{bdeOverviewSummaryModal.data?.agentTotalSales.toLocaleString()}</p>
                            </div>
                            <div className="p-3 bg-zinc-50 rounded-xl border border-zinc-100">
                              <p className="text-[10px] font-bold text-zinc-500 uppercase">Achievement</p>
                              <p className="text-sm font-bold">{bdeOverviewSummaryModal.data?.agentAchievement}%</p>
                            </div>
                            <div className="p-3 bg-zinc-50 rounded-xl border border-zinc-100">
                              <p className="text-[10px] font-bold text-zinc-500 uppercase">Net Profit</p>
                              <p className={cn("text-sm font-bold", bdeOverviewSummaryModal.data?.agentNetProfit >= 0 ? "text-emerald-600" : "text-red-600")}>
                                ₹{bdeOverviewSummaryModal.data?.agentNetProfit.toLocaleString()}
                              </p>
                            </div>
                          </div>
                          <div className="max-h-40 overflow-y-auto border border-zinc-100 rounded-xl">
                            <table className="w-full text-left text-[10px]">
                              <thead className="bg-zinc-50 sticky top-0">
                                <tr>
                                  <th className="px-2 py-1 border-b font-bold uppercase">Date</th>
                                  <th className="px-2 py-1 border-b font-bold uppercase">Trip ID</th>
                                  <th className="px-2 py-1 border-b font-bold uppercase">Guest Name</th>
                                  <th className="px-2 py-1 border-b font-bold uppercase text-right">Value</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-zinc-50">
                                {[...(bdeOverviewSummaryModal.data?.agentSalesData || [])]
                                  .sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime())
                                  .map((s: any, i: number) => (
                                    <tr key={i}>
                                      <td className="px-2 py-1 whitespace-nowrap">{s.date}</td>
                                      <td className="px-2 py-1 font-mono text-orange-600">{s.tripId}</td>
                                      <td className="px-2 py-1">{s.guestName}</td>
                                      <td className="px-2 py-1 text-right font-medium">₹{s.packageValue?.toLocaleString()}</td>
                                    </tr>
                                  ))}
                              </tbody>
                            </table>
                          </div>
                        </div>
                      )}

                      <div className="flex justify-end pt-4 border-t border-zinc-100">
                        <Button onClick={() => setBdeOverviewSummaryModal({ ...bdeOverviewSummaryModal, isOpen: false })}>
                          Close Summary
                        </Button>
                      </div>
                    </div>
                  </Modal>

                  <div className="mt-8">
                    <Last7DaysSalesOverview sales={sales} employees={employees} />
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {activeTab === 'statsOverview' && (
          <div className="space-y-8">
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-8 h-8 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Stats Overview</h2>
              </div>
              <div className="flex flex-wrap items-center gap-3">
                <select 
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl focus:ring-2 focus:ring-orange-500 outline-none text-xs font-bold uppercase shadow-sm"
                  value={statsOverviewSourceFilter}
                  onChange={(e) => setStatsOverviewSourceFilter(e.target.value)}
                >
                  <option value="All">Source: All</option>
                  {Array.from(new Set(sales.map(s => s.source))).filter(Boolean).sort().map(s => (
                    <option key={s} value={s}>{s}</option>
                  ))}
                </select>
                <div className="min-w-[150px]">
                  <SearchableSelect
                    options={[
                      { value: 'All', label: 'BDE: All' },
                      ...bdes.sort((a, b) => a.name.localeCompare(b.name)).map(b => ({ value: b.name, label: b.name, key: b.id }))
                    ]}
                    value={statsOverviewBdeFilter}
                    onChange={setStatsOverviewBdeFilter}
                    placeholder="BDE: All"
                  />
                </div>
                <select 
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl focus:ring-2 focus:ring-orange-500 outline-none text-xs font-bold uppercase shadow-sm"
                  value={statsOverviewMonthFilter}
                  onChange={(e) => {
                    setStatsOverviewMonthFilter(e.target.value);
                    setStatsOverviewWeekFilter('All');
                  }}
                >
                  <option value="All">Month: All</option>
                  {allMonths.map(m => (
                    <option key={m} value={m}>{m}</option>
                  ))}
                </select>
                <select 
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl focus:ring-2 focus:ring-orange-500 outline-none text-xs font-bold uppercase shadow-sm"
                  value={statsOverviewWeekFilter}
                  onChange={(e) => setStatsOverviewWeekFilter(e.target.value)}
                >
                  <option value="All">Week: All</option>
                  {allSortedWeeks
                    .filter(w => statsOverviewMonthFilter === 'All' || w.month === statsOverviewMonthFilter)
                    .map(w => (
                      <option key={w.id} value={w.weekName}>{w.weekName}</option>
                    ))}
                </select>
                <select 
                  className="px-3 py-2 bg-white border border-zinc-200 rounded-xl focus:ring-2 focus:ring-orange-500 outline-none text-xs font-bold uppercase shadow-sm"
                  value={statsOverviewStatusFilter}
                  onChange={(e) => setStatsOverviewStatusFilter(e.target.value)}
                >
                  <option value="All">Status: All</option>
                  <option value="Advance">Advance</option>
                  <option value="Credit Note">Credit Note</option>
                  <option value="Done">Done</option>
                  <option value="Confirmed">Confirmed</option>
                  <option value="Cancel">Cancel</option>
                </select>
              </div>
            </div>

            {(() => {
              const baseFilteredSales = sales.filter(sale => {
                const matchesSource = statsOverviewSourceFilter === 'All' || sale.source === statsOverviewSourceFilter;
                const matchesStatus = statsOverviewStatusFilter === 'All' || sale.advanceCN === statsOverviewStatusFilter;
                const matchesWeek = statsOverviewWeekFilter === 'All' || sale.week === statsOverviewWeekFilter;
                
                const saleMonth = getSaleMonth(sale, weeks);
                const matchesMonth = statsOverviewMonthFilter === 'All' || (saleMonth === statsOverviewMonthFilter);
                
                return matchesSource && matchesStatus && matchesWeek && matchesMonth;
              });

              const filteredStatsSales = baseFilteredSales.flatMap(sale => {
                const entries = [];
                const agentBde = findEmployeeBySale(sale.agent, sale.agentEmail)?.bde || findBDEBySale(sale.bde, sale.bdeEmail)?.name || sale.bde || 'Unknown';
                const agentCredit = (sale.packageValue || 0) * ((sale.agentPercentage ?? 100) / 100);
                entries.push({ ...sale, bde: agentBde, packageValue: agentCredit, isPrimary: true });
                
                if (sale.associateBde) {
                  const assocBde = findEmployeeBySale(sale.associateBde, sale.associateBdeEmail)?.bde || findBDEBySale(sale.bde, sale.bdeEmail)?.name || sale.bde || 'Unknown';
                  const assocCredit = (sale.packageValue || 0) * ((sale.assocBdePercentage ?? 0) / 100);
                  if (assocCredit > 0) {
                    entries.push({ ...sale, bde: assocBde, packageValue: assocCredit, isPrimary: false });
                  }
                }
                return entries;
              }).filter(entry => statsOverviewBdeFilter === 'All' || entry.bde === statsOverviewBdeFilter);

              // Data for Destination Chart
              const destDataRaw = filteredStatsSales.reduce((acc: any, sale) => {
                const dest = (sale.destination || 'Unknown').trim().toLowerCase();
                if (!acc[dest]) acc[dest] = { name: dest, salesCount: 0, totalValue: 0, totalPax: 0 };
                if (sale.isPrimary) acc[dest].salesCount += 1;
                acc[dest].totalValue += (sale.packageValue || 0);
                if (sale.isPrimary) acc[dest].totalPax += (sale.noOfPax || 0);
                return acc;
              }, {});

              const totalSalesCountAll = filteredStatsSales.filter(s => s.isPrimary).length;
              const totalValueAll = filteredStatsSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);

              const destData = Object.values(destDataRaw).map((val: any) => ({
                name: val.name.charAt(0).toUpperCase() + val.name.slice(1),
                salesCount: val.salesCount,
                totalValue: val.totalValue,
                totalPax: val.totalPax,
                avgSalesValue: val.salesCount > 0 ? Math.round(val.totalValue / val.salesCount) : 0,
                salesCountPerc: totalSalesCountAll > 0 ? (val.salesCount / totalSalesCountAll) * 100 : 0,
                totalValuePerc: totalValueAll > 0 ? (val.totalValue / totalValueAll) * 100 : 0
              }))
              .sort((a, b) => {
                if (statsOverviewDestMetric === 'totalValue') return b.totalValue - a.totalValue;
                if (statsOverviewDestMetric === 'salesCount') return b.salesCount - a.salesCount;
                if (statsOverviewDestMetric === 'totalPax') return b.totalPax - a.totalPax;
                return b.avgSalesValue - a.avgSalesValue;
              })
              .slice(0, 10);

              // Data for Source Chart
              const sourceDataRaw = filteredStatsSales.reduce((acc: any, sale) => {
                const source = (sale.source || 'Unknown').trim();
                const salesBy = sale.salesBy || 'Inhouse';
                
                if (!acc[source]) {
                  acc[source] = { 
                    name: source, 
                    count: 0, 
                    value: 0,
                    inhouseCount: 0,
                    inhouseValue: 0,
                    branchCount: 0,
                    branchValue: 0,
                    franchiseeCount: 0,
                    franchiseeValue: 0
                  };
                }
                
                acc[source].count += 1;
                acc[source].value += (sale.packageValue || 0);
                
                if (salesBy === 'Inhouse') {
                  acc[source].inhouseCount += 1;
                  acc[source].inhouseValue += (sale.packageValue || 0);
                } else if (salesBy === 'Branch') {
                  acc[source].branchCount += 1;
                  acc[source].branchValue += (sale.packageValue || 0);
                } else if (salesBy === 'Franchisee Sales') {
                  acc[source].franchiseeCount += 1;
                  acc[source].franchiseeValue += (sale.packageValue || 0);
                }
                
                return acc;
              }, {});

              const totalSourceCount = filteredStatsSales.length;
              const totalSourceValue = filteredStatsSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);

              const sourceData = Object.values(sourceDataRaw).map((val: any) => ({
                ...val,
                countPerc: totalSourceCount > 0 ? (val.count / totalSourceCount) * 100 : 0,
                valuePerc: totalSourceValue > 0 ? (val.value / totalSourceValue) * 100 : 0
              })).sort((a: any, b: any) => b.count - a.count);

              // Data for Trend Chart (by Week)
              const trendData = Object.entries(
                filteredStatsSales.reduce((acc: any, sale) => {
                  const weekId = sale.week || 'Unknown';
                  if (!acc[weekId]) {
                    const weekInfo = weeks.find(w => w.weekName === weekId);
                    acc[weekId] = { 
                      weekId, 
                      weekName: weekInfo ? weekInfo.weekName : weekId,
                      startDate: weekInfo ? weekInfo.startDate : '9999-99-99',
                      sales: 0, 
                      margin: 0 
                    };
                  }
                  acc[weekId].sales += (sale.packageValue || 0);
                  acc[weekId].margin += (sale.totalMargin || 0);
                  return acc;
                }, {})
              ).map(([_, val]) => val)
               .sort((a: any, b: any) => {
                 return new Date(a.startDate).getTime() - new Date(b.startDate).getTime();
               });

              // Data for Status Chart
              const statusDataRaw = filteredStatsSales.reduce((acc: any, sale) => {
                const status = (sale.advanceCN || 'Pending').trim();
                const salesBy = sale.salesBy || 'Inhouse';
                
                if (!acc[status]) {
                  acc[status] = { 
                    name: status, 
                    count: 0, 
                    value: 0,
                    inhouseCount: 0,
                    inhouseValue: 0,
                    branchCount: 0,
                    branchValue: 0,
                    franchiseeCount: 0,
                    franchiseeValue: 0
                  };
                }
                
                acc[status].count += 1;
                acc[status].value += (sale.packageValue || 0);
                
                if (salesBy === 'Inhouse') {
                  acc[status].inhouseCount += 1;
                  acc[status].inhouseValue += (sale.packageValue || 0);
                } else if (salesBy === 'Branch') {
                  acc[status].branchCount += 1;
                  acc[status].branchValue += (sale.packageValue || 0);
                } else if (salesBy === 'Franchisee Sales') {
                  acc[status].franchiseeCount += 1;
                  acc[status].franchiseeValue += (sale.packageValue || 0);
                }
                
                return acc;
              }, {});

              const totalStatusCount = filteredStatsSales.length;
              const totalStatusValue = filteredStatsSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);

              const statusData = Object.values(statusDataRaw).map((val: any) => ({
                ...val,
                countPerc: totalStatusCount > 0 ? (val.count / totalStatusCount) * 100 : 0,
                valuePerc: totalStatusValue > 0 ? (val.value / totalStatusValue) * 100 : 0
              })).sort((a: any, b: any) => b.count - a.count);

              const totalSalesValue = filteredStatsSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);
              const totalMarginValue = filteredStatsSales.reduce((sum, s) => sum + (s.totalMargin || 0), 0);
              const avgPackageValue = filteredStatsSales.length > 0 ? totalSalesValue / filteredStatsSales.length : 0;
              const numSales = filteredStatsSales.length;

              const getBreakdown = (salesList: any[]) => {
                return salesList.reduce((acc: any, s: any) => {
                  const source = s.salesBy || 'Inhouse';
                  if (!acc[source]) acc[source] = { count: 0, value: 0, margin: 0 };
                  acc[source].count += 1;
                  acc[source].value += (s.packageValue || 0);
                  acc[source].margin += (s.totalMargin || 0);
                  return acc;
                }, {
                  'Inhouse': { count: 0, value: 0, margin: 0 },
                  'Branch': { count: 0, value: 0, margin: 0 },
                  'Franchisee Sales': { count: 0, value: 0, margin: 0 }
                });
              };

              const salesBreakdown = getBreakdown(filteredStatsSales);
              const advanceSales = filteredStatsSales.filter(s => s.advanceCN === 'Advance');
              const advanceBreakdown = getBreakdown(advanceSales);
              const advanceCount = advanceSales.length;
              const advanceValue = advanceSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);
              const advancePerc = totalSalesValue > 0 ? (advanceValue / totalSalesValue) * 100 : 0;

              const cnSales = filteredStatsSales.filter(s => s.advanceCN === 'Credit Note');
              const cnBreakdown = getBreakdown(cnSales);
              const cnCount = cnSales.length;
              const cnValue = cnSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);
              const cnPerc = totalSalesValue > 0 ? (cnValue / totalSalesValue) * 100 : 0;

              const cancelSales = filteredStatsSales.filter(s => s.advanceCN === 'Cancel');
              const cancelBreakdown = getBreakdown(cancelSales);
              const cancelCount = cancelSales.length;
              const cancelValue = cancelSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);
              const cancelPerc = totalSalesValue > 0 ? (cancelValue / totalSalesValue) * 100 : 0;

              const pendingSales = filteredStatsSales.filter(s => s.workPending === 'Issue');
              const pendingBreakdown = getBreakdown(pendingSales);
              const workPendingCount = pendingSales.length;

              // Top 5 Destinations by Source
              const topDestinationsBySource = Object.keys(sourceDataRaw).map(sourceName => {
                const sourceSales = filteredStatsSales.filter(s => (s.source || 'Unknown').trim() === sourceName);
                const destStats = sourceSales.reduce((acc: any, s) => {
                  const rawDest = (s.destination || 'Unknown').trim();
                  const destKey = rawDest.toLowerCase();
                  if (!acc[destKey]) {
                    acc[destKey] = { name: rawDest, count: 0, amount: 0, margin: 0 };
                  }
                  acc[destKey].count += 1;
                  acc[destKey].amount += (s.packageValue || 0);
                  acc[destKey].margin += (s.totalMargin || 0);
                  return acc;
                }, {});
                
                const top5 = Object.values(destStats)
                  .sort((a: any, b: any) => b.count - a.count)
                  .slice(0, 5);
                  
                return { source: sourceName, top5 };
              }).sort((a, b) => {
                const aCount = sourceDataRaw[a.source]?.count || 0;
                const bCount = sourceDataRaw[b.source]?.count || 0;
                return bCount - aCount;
              });

              // Base sales for comparison (ignores week/month/BDE filters but keeps source/status)
              const baseSalesForPackageAnalysis = sales.filter(sale => {
                const matchesSource = statsOverviewSourceFilter === 'All' || sale.source === statsOverviewSourceFilter;
                const matchesStatus = statsOverviewStatusFilter === 'All' || sale.advanceCN === statsOverviewStatusFilter;
                return matchesSource && matchesStatus;
              });

              // Data for Daily Sales Chart
              const dailySalesDataRaw = filteredStatsSales.reduce((acc: any, sale) => {
                const date = sale.date || 'Unknown';
                if (!acc[date]) acc[date] = { date, count: 0, value: 0 };
                acc[date].count += 1;
                acc[date].value += (sale.packageValue || 0);
                return acc;
              }, {});

              const dailySalesData = Object.values(dailySalesDataRaw)
                .sort((a: any, b: any) => new Date(a.date).getTime() - new Date(b.date).getTime());

              // Package Amount Analysis Data
              const packageRanges = [
                { label: 'Up to ₹50k', min: 0, max: 50000 },
                { label: '₹50k - ₹1.5L', min: 50001, max: 150000 },
                { label: '₹1.5L - ₹3L', min: 150001, max: 300000 },
                { label: 'Above ₹3L', min: 300001, max: Infinity }
              ];

              // Determine comparison frames
              let comparisonFrames: string[] = [];
              if (statsOverviewPackageAnalysisDimension === 'BDE') {
                const bdeSales = baseSalesForPackageAnalysis.reduce((acc: any, s) => {
                  const bde = s.bde || 'Unknown';
                  acc[bde] = (acc[bde] || 0) + (s.packageValue || 0);
                  return acc;
                }, {});
                comparisonFrames = Object.entries(bdeSales)
                  .sort((a: any, b: any) => b[1] - a[1])
                  .slice(0, 5)
                  .map(entry => entry[0]);
              } else if (statsOverviewPackageAnalysisDimension === 'Monthly') {
                comparisonFrames = allMonths.slice(-5);
              } else if (statsOverviewPackageAnalysisDimension === 'Weekly') {
                comparisonFrames = allSortedWeeks.slice(-5).map(w => w.weekName);
              } else if (statsOverviewPackageAnalysisDimension === 'Destination') {
                const destSales = baseSalesForPackageAnalysis.reduce((acc: any, s) => {
                  const rawDest = (s.destination || 'Unknown').trim();
                  const destKey = rawDest.toLowerCase();
                  if (!acc[destKey]) {
                    acc[destKey] = { name: rawDest, value: 0 };
                  }
                  acc[destKey].value += (s.packageValue || 0);
                  return acc;
                }, {});
                comparisonFrames = Object.values(destSales)
                  .sort((a: any, b: any) => b.value - a.value)
                  .slice(0, 5)
                  .map((entry: any) => entry.name);
              } else if (statsOverviewPackageAnalysisDimension === 'Source') {
                const sourceSales = baseSalesForPackageAnalysis.reduce((acc: any, s) => {
                  const rawSource = (s.source || 'Unknown').trim();
                  const sourceKey = rawSource.toLowerCase();
                  if (!acc[sourceKey]) {
                    acc[sourceKey] = { name: rawSource, value: 0 };
                  }
                  acc[sourceKey].value += (s.packageValue || 0);
                  return acc;
                }, {});
                comparisonFrames = Object.values(sourceSales)
                  .sort((a: any, b: any) => b.value - a.value)
                  .slice(0, 5)
                  .map((entry: any) => entry.name);
              }

              const packageAnalysisData = packageRanges.map(range => {
                const salesInRange = baseSalesForPackageAnalysis.filter(s => (s.packageValue || 0) >= range.min && (s.packageValue || 0) <= range.max);
                
                const frameData: any = {};
                comparisonFrames.forEach((frame, idx) => {
                  let frameSales = [];
                  if (statsOverviewPackageAnalysisDimension === 'BDE') {
                    frameSales = salesInRange.filter(s => s.bde === frame);
                  } else if (statsOverviewPackageAnalysisDimension === 'Monthly') {
                    frameSales = salesInRange.filter(s => {
                      const saleMonth = getSaleMonth(s, weeks);
                      return saleMonth === frame;
                    });
                  } else if (statsOverviewPackageAnalysisDimension === 'Weekly') {
                    frameSales = salesInRange.filter(s => s.week === frame);
                  } else if (statsOverviewPackageAnalysisDimension === 'Destination') {
                    frameSales = salesInRange.filter(s => (s.destination || 'Unknown').trim().toLowerCase() === frame.toLowerCase());
                  } else if (statsOverviewPackageAnalysisDimension === 'Source') {
                    frameSales = salesInRange.filter(s => (s.source || 'Unknown').trim().toLowerCase() === frame.toLowerCase());
                  }
                  
                  const value = frameSales.reduce((sum, s) => sum + (s.packageValue || 0), 0);
                  const margin = frameSales.reduce((sum, s) => sum + (s.totalMargin || 0), 0);
                  const inhouseCount = frameSales.filter(s => (s.salesBy || 'Inhouse') === 'Inhouse').length;
                  const branchCount = frameSales.filter(s => (s.salesBy || 'Inhouse') === 'Branch').length;
                  const franchiseeCount = frameSales.filter(s => (s.salesBy || 'Inhouse') === 'Franchisee Sales').length;
                  
                  frameData[`frame${idx}Count`] = frameSales.length;
                  frameData[`frame${idx}Value`] = value;
                  frameData[`frame${idx}Margin`] = margin;
                  frameData[`frame${idx}MarginPerc`] = value > 0 ? (margin / value) * 100 : 0;
                  frameData[`frame${idx}InhouseCount`] = inhouseCount;
                  frameData[`frame${idx}BranchCount`] = branchCount;
                  frameData[`frame${idx}FranchiseeCount`] = franchiseeCount;
                  frameData[`frame${idx}Name`] = frame;
                });

                const totalValue = salesInRange.reduce((sum, s) => sum + (s.packageValue || 0), 0);
                const totalMargin = salesInRange.reduce((sum, s) => sum + (s.totalMargin || 0), 0);
                const totalInhouseCount = salesInRange.filter(s => (s.salesBy || 'Inhouse') === 'Inhouse').length;
                const totalBranchCount = salesInRange.filter(s => (s.salesBy || 'Inhouse') === 'Branch').length;
                const totalFranchiseeCount = salesInRange.filter(s => (s.salesBy || 'Inhouse') === 'Franchisee Sales').length;
                
                return {
                  range: range.label,
                  count: salesInRange.length,
                  value: totalValue,
                  margin: totalMargin,
                  marginPerc: totalValue > 0 ? (totalMargin / totalValue) * 100 : 0,
                  totalInhouseCount,
                  totalBranchCount,
                  totalFranchiseeCount,
                  ...frameData
                };
              });

              const COLORS = ['#ea580c', '#3b82f6', '#10b981', '#f59e0b', '#6366f1', '#ec4899', '#8b5cf6', '#14b8a6'];

              return (
                <div className="space-y-8">
                  {/* Last 7 Days Overview Table */}
                  <Last7DaysSalesOverview 
                    sales={sales} 
                    employees={employees}
                    sourceFilter={statsOverviewSourceFilter} 
                    bdeFilter={statsOverviewBdeFilter} 
                  />

                  {/* Summary Cards */}
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-6">
                    <Card 
                      className="p-6 bg-orange-600 text-white border-none shadow-lg shadow-orange-200 cursor-pointer hover:scale-[1.02] transition-transform"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Total Sales Details', sales: filteredStatsSales })}
                    >
                      <p className="text-orange-100 text-xs font-bold uppercase tracking-wider">Total Sales Value</p>
                      <p className="text-3xl font-black mt-2">₹{totalSalesValue.toLocaleString()}</p>
                      <div className="mt-4 flex items-center gap-2 text-orange-100 text-xs">
                        <TrendingUp className="w-4 h-4" />
                        <span>{numSales} Transactions</span>
                      </div>
                      <div className="mt-4 pt-4 border-t border-white/20 grid grid-cols-3 gap-2 text-[12px] font-bold uppercase">
                        <div>
                          <p className="text-orange-200 mb-1">Inhouse</p>
                          <p>₹{salesBreakdown['Inhouse'].value.toLocaleString()}</p>
                        </div>
                        <div>
                          <p className="text-orange-200 mb-1">Branch</p>
                          <p>₹{salesBreakdown['Branch'].value.toLocaleString()}</p>
                        </div>
                        <div>
                          <p className="text-orange-200 mb-1">Franchisee</p>
                          <p>₹{salesBreakdown['Franchisee Sales'].value.toLocaleString()}</p>
                        </div>
                      </div>
                    </Card>
                    <Card 
                      className="p-6 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Margin Details', sales: filteredStatsSales })}
                    >
                      <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Total Margin</p>
                      <p className="text-3xl font-black mt-2 text-emerald-600">₹{totalMarginValue.toLocaleString()}</p>
                      <p className="text-[10px] text-zinc-500 mt-2 italic">Avg Margin: {totalSalesValue > 0 ? ((totalMarginValue / totalSalesValue) * 100).toFixed(1) : 0}%</p>
                      <div className="mt-4 pt-4 border-t border-zinc-100 grid grid-cols-3 gap-2 text-[12px] font-bold uppercase">
                        <div>
                          <p className="text-zinc-400 mb-1">Inhouse</p>
                          <p className="text-zinc-700">₹{salesBreakdown['Inhouse'].margin.toLocaleString()}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Branch</p>
                          <p className="text-zinc-700">₹{salesBreakdown['Branch'].margin.toLocaleString()}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Franchisee</p>
                          <p className="text-zinc-700">₹{salesBreakdown['Franchisee Sales'].margin.toLocaleString()}</p>
                        </div>
                      </div>
                    </Card>
                    <Card 
                      className="p-6 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Package Value Details', sales: filteredStatsSales })}
                    >
                      <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Avg Package Value</p>
                      <p className="text-3xl font-black mt-2 text-blue-600">₹{Math.round(avgPackageValue).toLocaleString()}</p>
                      <p className="text-[10px] text-zinc-500 mt-2 italic">{numSales} Sales</p>
                      <div className="mt-4 pt-4 border-t border-zinc-100 grid grid-cols-3 gap-2 text-[12px] font-bold uppercase">
                        <div>
                          <p className="text-zinc-400 mb-1">Inhouse</p>
                          <p className="text-zinc-700">₹{salesBreakdown['Inhouse'].count > 0 ? Math.round(salesBreakdown['Inhouse'].value / salesBreakdown['Inhouse'].count).toLocaleString() : 0}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Branch</p>
                          <p className="text-zinc-700">₹{salesBreakdown['Branch'].count > 0 ? Math.round(salesBreakdown['Branch'].value / salesBreakdown['Branch'].count).toLocaleString() : 0}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Franchisee</p>
                          <p className="text-zinc-700">₹{salesBreakdown['Franchisee Sales'].count > 0 ? Math.round(salesBreakdown['Franchisee Sales'].value / salesBreakdown['Franchisee Sales'].count).toLocaleString() : 0}</p>
                        </div>
                      </div>
                    </Card>
                    <Card 
                      className="p-6 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Advance Status Details', sales: advanceSales })}
                    >
                      <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Advance Status</p>
                      <p className="text-3xl font-black mt-2 text-orange-600">{advanceCount}</p>
                      <p className="text-[10px] text-zinc-500 mt-2 italic">Value: ₹{advanceValue.toLocaleString()} ({advancePerc.toFixed(1)}%)</p>
                      <div className="mt-4 pt-4 border-t border-zinc-100 grid grid-cols-3 gap-2 text-[12px] font-bold uppercase">
                        <div>
                          <p className="text-zinc-400 mb-1">Inhouse</p>
                          <p className="text-zinc-700">{advanceBreakdown['Inhouse'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Branch</p>
                          <p className="text-zinc-700">{advanceBreakdown['Branch'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Franchisee</p>
                          <p className="text-zinc-700">{advanceBreakdown['Franchisee Sales'].count}</p>
                        </div>
                      </div>
                    </Card>
                    <Card 
                      className="p-6 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Credit Note Details', sales: cnSales })}
                    >
                      <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Credit Note Status</p>
                      <p className="text-3xl font-black mt-2 text-purple-600">{cnCount}</p>
                      <p className="text-[10px] text-zinc-500 mt-2 italic">Value: ₹{cnValue.toLocaleString()} ({cnPerc.toFixed(1)}%)</p>
                      <div className="mt-4 pt-4 border-t border-zinc-100 grid grid-cols-3 gap-2 text-[12px] font-bold uppercase">
                        <div>
                          <p className="text-zinc-400 mb-1">Inhouse</p>
                          <p className="text-zinc-700">{cnBreakdown['Inhouse'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Branch</p>
                          <p className="text-zinc-700">{cnBreakdown['Branch'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Franchisee</p>
                          <p className="text-zinc-700">{cnBreakdown['Franchisee Sales'].count}</p>
                        </div>
                      </div>
                    </Card>
                    <Card 
                      className="p-6 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Cancel Status Details', sales: cancelSales })}
                    >
                      <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Cancel Status</p>
                      <p className="text-3xl font-black mt-2 text-red-600">{cancelCount}</p>
                      <p className="text-[10px] text-zinc-500 mt-2 italic">Value: ₹{cancelValue.toLocaleString()} ({cancelPerc.toFixed(1)}%)</p>
                      <div className="mt-4 pt-4 border-t border-zinc-100 grid grid-cols-3 gap-2 text-[12px] font-bold uppercase">
                        <div>
                          <p className="text-zinc-400 mb-1">Inhouse</p>
                          <p className="text-zinc-700">{cancelBreakdown['Inhouse'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Branch</p>
                          <p className="text-zinc-700">{cancelBreakdown['Branch'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Franchisee</p>
                          <p className="text-zinc-700">{cancelBreakdown['Franchisee Sales'].count}</p>
                        </div>
                      </div>
                    </Card>
                    <Card 
                      className="p-6 bg-white border-zinc-200 shadow-sm cursor-pointer hover:shadow-md transition-shadow"
                      onClick={() => setSummaryDetailModal({ isOpen: true, title: 'Work Pending Details', sales: filteredStatsSales.filter(s => s.workPending === 'Issue') })}
                    >
                      <p className="text-zinc-400 text-xs font-bold uppercase tracking-wider">Work Pending Issues</p>
                      <p className="text-3xl font-black mt-2 text-amber-600">{workPendingCount}</p>
                      <p className="text-[10px] text-zinc-500 mt-2 italic">Issues in Sales Master</p>
                      <div className="mt-4 pt-4 border-t border-zinc-100 grid grid-cols-3 gap-2 text-[12px] font-bold uppercase">
                        <div>
                          <p className="text-zinc-400 mb-1">Inhouse</p>
                          <p className="text-zinc-700">{pendingBreakdown['Inhouse'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Branch</p>
                          <p className="text-zinc-700">{pendingBreakdown['Branch'].count}</p>
                        </div>
                        <div>
                          <p className="text-zinc-400 mb-1">Franchisee</p>
                          <p className="text-zinc-700">{pendingBreakdown['Franchisee Sales'].count}</p>
                        </div>
                      </div>
                    </Card>
                  </div>

                  {/* Daily Sales Analysis Chart */}
                  <Card className="p-6">
                    <h3 className="text-sm font-bold text-zinc-900 uppercase mb-6 flex items-center gap-2">
                      <CalendarIcon className="w-4 h-4 text-orange-600" />
                      Daily Sales Analysis (Value & Count)
                    </h3>
                    <div className="h-[500px] w-full">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={dailySalesData} margin={{ top: 130, right: 30, left: 20, bottom: 60 }}>
                          <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                          <XAxis 
                            dataKey="date" 
                            fontSize={10} 
                            fontWeight="bold"
                            angle={-45}
                            textAnchor="end"
                            interval={0}
                            tickFormatter={(val) => val === 'Unknown' ? val : new Date(val).toLocaleDateString('en-IN', { day: '2-digit', month: 'short' })}
                          />
                          <YAxis fontSize={10} tickFormatter={(val) => `₹${(val/100000).toFixed(2)}L`} />
                          <Tooltip 
                            contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                            formatter={(value: number, name: string) => [name === 'value' ? `₹${value.toLocaleString()}` : value, name === 'value' ? 'Amount' : 'Sales Count']}
                          />
                          <Bar dataKey="value" name="Amount" fill="#ea580c" radius={[4, 4, 0, 0]}>
                            <LabelList 
                              dataKey="value" 
                              position="top" 
                              content={(props: any) => {
                                const { x, y, width, value, index } = props;
                                if (!dailySalesData[index]) return null;
                                const count = (dailySalesData[index] as any).count;
                                return (
                                  <g transform={`translate(${x + width / 2},${y - 45})`}>
                                    <text 
                                      x={0} 
                                      y={0} 
                                      fill="#ea580c" 
                                      textAnchor="start" 
                                      fontSize={14} 
                                      fontWeight="black"
                                      transform="rotate(-90)"
                                    >
                                      {`₹${(value / 100000).toFixed(2)}L (${count})`}
                                    </text>
                                  </g>
                                );
                              }}
                            />
                          </Bar>
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </Card>

                  {/* Charts Grid */}
                  <div className="grid grid-cols-1 lg:grid-cols-2 gap-8">
                    {/* Sales Trend */}
                    <Card className="p-6">
                      <h3 className="text-sm font-bold text-zinc-900 uppercase mb-6 flex items-center gap-2">
                        <TrendingUp className="w-4 h-4 text-orange-600" />
                        Sales & Margin Trend
                      </h3>
                      <div className="h-80 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={trendData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis 
                              dataKey="weekName" 
                              fontSize={10} 
                              fontWeight="bold"
                            />
                            <YAxis fontSize={10} tickFormatter={(val) => `₹${(val/100000).toFixed(2)}L`} />
                            <Tooltip 
                              contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                              formatter={(value: number, name: string, props: any) => {
                                const data = props.payload as any;
                                if (name === 'Margin' && data && data.sales > 0) {
                                  const percentage = ((value / data.sales) * 100).toFixed(1);
                                  return [`₹${value.toLocaleString()} (${percentage}%)`, name];
                                }
                                return [`₹${value.toLocaleString()}`, name];
                              }}
                            />
                            <Legend verticalAlign="top" height={36}/>
                             <Bar dataKey="sales" name="Sales" fill="#ea580c" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey="sales" 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value } = props;
                                  return (
                                    <text x={x + width / 2} y={y - 5} fill="#ea580c" textAnchor="middle" fontSize={11} fontWeight="black">
                                      ₹{(value / 100000).toFixed(2)}L
                                    </text>
                                  );
                                }}
                              />
                            </Bar>
                            <Bar dataKey="margin" name="Margin" fill="#10b981" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey="margin" 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  if (!trendData[index]) return null;
                                  const data = trendData[index] as any;
                                  const percentage = data && data.sales > 0 ? ((data.margin / data.sales) * 100).toFixed(1) : 0;
                                  return (
                                    <text x={x + width / 2} y={y - 20} fill="#059669" textAnchor="middle" fontSize={11} fontWeight="black">
                                      <tspan x={x + width / 2} dy="0">₹{(value / 100000).toFixed(2)}L</tspan>
                                      <tspan x={x + width / 2} dy="10">({percentage}%)</tspan>
                                    </text>
                                  );
                                }}
                              />
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </Card>

                    {/* Destination Analysis */}
                    <Card className="p-6">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                        <h3 className="text-sm font-bold text-zinc-900 uppercase flex items-center gap-2">
                          <Star className="w-4 h-4 text-orange-600" />
                          Top 10 Destinations
                        </h3>
                        <div className="flex flex-wrap gap-2">
                          {[
                            { id: 'totalValue', label: 'Value' },
                            { id: 'salesCount', label: 'Count' },
                            { id: 'totalPax', label: 'Pax' },
                            { id: 'avgSalesValue', label: 'Avg/Sales' }
                          ].map((m) => (
                            <button
                              key={m.id}
                              onClick={() => setStatsOverviewDestMetric(m.id as any)}
                              className={`px-3 py-1 rounded-full text-[10px] font-bold transition-all ${
                                statsOverviewDestMetric === m.id
                                  ? 'bg-orange-600 text-white shadow-md'
                                  : 'bg-zinc-100 text-zinc-500 hover:bg-zinc-200'
                              }`}
                            >
                              {m.label}
                            </button>
                          ))}
                        </div>
                      </div>
                      <div className="h-80 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={destData} margin={{ top: 35, right: 30, left: 20, bottom: 5 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="name" fontSize={10} fontWeight="bold" />
                            <YAxis 
                              fontSize={10} 
                              tickFormatter={(val) => {
                                if (statsOverviewDestMetric === 'totalValue' || statsOverviewDestMetric === 'avgSalesValue') return `₹${(val/100000).toFixed(2)}L`;
                                return val;
                              }} 
                            />
                            <Tooltip 
                              contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                              formatter={(value: number) => {
                                if (statsOverviewDestMetric === 'totalValue' || statsOverviewDestMetric === 'avgSalesValue') return [`₹${value.toLocaleString()}`, 'Value'];
                                return [value.toLocaleString(), 'Value'];
                              }}
                            />
                            <Bar dataKey={statsOverviewDestMetric} fill="#3b82f6" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey={statsOverviewDestMetric} 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  const data = destData[index] as any;
                                  if (!data) return null;

                                  let mainLabel = '';
                                  let subLabel = '';

                                  if (statsOverviewDestMetric === 'totalValue') {
                                    mainLabel = `₹${(value/100000).toFixed(2)}L`;
                                    subLabel = `(${data.totalValuePerc.toFixed(1)}%)`;
                                  } else if (statsOverviewDestMetric === 'salesCount') {
                                    mainLabel = `${value}`;
                                    subLabel = `(${data.salesCountPerc.toFixed(1)}%)`;
                                  } else if (statsOverviewDestMetric === 'avgSalesValue') {
                                    mainLabel = `₹${(value/100000).toFixed(2)}L`;
                                  } else {
                                    mainLabel = value.toString();
                                  }

                                  return (
                                    <text x={x + width / 2} y={y - 20} fill="#1e40af" textAnchor="middle" fontSize={9} fontWeight="bold">
                                      <tspan x={x + width / 2} dy="0">{mainLabel}</tspan>
                                      {subLabel && <tspan x={x + width / 2} dy="10">{subLabel}</tspan>}
                                    </text>
                                  );
                                }}
                              />
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                    </Card>

                    {/* Source Distribution */}
                    <Card className="p-6">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                        <h3 className="text-sm font-bold text-zinc-900 uppercase flex items-center gap-2">
                          <History className="w-4 h-4 text-orange-600" />
                          Lead Source Distribution
                        </h3>
                        <div className="flex gap-2">
                          <button
                            onClick={() => setStatsOverviewSourceMetric('count')}
                            className={`px-3 py-1 rounded-full text-[10px] font-bold transition-all ${
                              statsOverviewSourceMetric === 'count'
                                ? 'bg-orange-600 text-white shadow-md'
                                : 'bg-zinc-100 text-zinc-500 hover:bg-zinc-200'
                            }`}
                          >
                            Count
                          </button>
                          <button
                            onClick={() => setStatsOverviewSourceMetric('value')}
                            className={`px-3 py-1 rounded-full text-[10px] font-bold transition-all ${
                              statsOverviewSourceMetric === 'value'
                                ? 'bg-orange-600 text-white shadow-md'
                                : 'bg-zinc-100 text-zinc-500 hover:bg-zinc-200'
                            }`}
                          >
                            Value
                          </button>
                        </div>
                      </div>
                      <div className="h-80 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={sourceData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="name" fontSize={10} fontWeight="bold" />
                            <YAxis 
                              fontSize={10} 
                              tickFormatter={(val) => {
                                if (statsOverviewSourceMetric === 'value') return `₹${(val/100000).toFixed(2)}L`;
                                return val;
                              }} 
                            />
                            <Tooltip 
                              contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                              formatter={(value: number) => {
                                if (statsOverviewSourceMetric === 'value') return [`₹${value.toLocaleString()}`, 'Value'];
                                return [value.toLocaleString(), 'Count'];
                              }}
                            />
                            <Legend verticalAlign="top" height={36}/>
                            <Bar 
                              dataKey={statsOverviewSourceMetric === 'count' ? 'inhouseCount' : 'inhouseValue'} 
                              name="Inhouse" 
                              stackId="a" 
                              fill="#ea580c" 
                              radius={[0, 0, 0, 0]} 
                            />
                            <Bar 
                              dataKey={statsOverviewSourceMetric === 'count' ? 'branchCount' : 'branchValue'} 
                              name="Branch" 
                              stackId="a" 
                              fill="#3b82f6" 
                              radius={[0, 0, 0, 0]} 
                            />
                            <Bar 
                              dataKey={statsOverviewSourceMetric === 'count' ? 'franchiseeCount' : 'franchiseeValue'} 
                              name="Franchisee" 
                              stackId="a" 
                              fill="#10b981" 
                              radius={[4, 4, 0, 0]} 
                            />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>
                      
                      {/* Detailed Source Table */}
                      <div className="mt-8 overflow-x-auto">
                        <table className="w-full text-left text-[10px]">
                          <thead className="bg-zinc-50 text-zinc-500 uppercase font-bold">
                            <tr>
                              <th className="px-4 py-2">Source</th>
                              <th className="px-4 py-2">Count (%)</th>
                              <th className="px-4 py-2">Value (%)</th>
                              <th className="px-4 py-2">Inhouse (C/V)</th>
                              <th className="px-4 py-2">Branch (C/V)</th>
                              <th className="px-4 py-2">Franchisee (C/V)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-zinc-100">
                            {sourceData.map((s: any) => (
                              <tr key={s.name} className="hover:bg-zinc-50 transition-colors">
                                <td className="px-4 py-2 font-bold text-zinc-900">{s.name}</td>
                                <td className="px-4 py-2">
                                  <span className="font-bold">{s.count}</span>
                                  <span className="text-zinc-400 ml-1">({s.countPerc.toFixed(1)}%)</span>
                                </td>
                                <td className="px-4 py-2">
                                  <span className="font-bold">₹{(s.value/100000).toFixed(2)}L</span>
                                  <span className="text-zinc-400 ml-1">({s.valuePerc.toFixed(1)}%)</span>
                                </td>
                                <td className="px-4 py-2 text-zinc-600">
                                  {s.inhouseCount} / ₹{(s.inhouseValue/100000).toFixed(2)}L
                                </td>
                                <td className="px-4 py-2 text-zinc-600">
                                  {s.branchCount} / ₹{(s.branchValue/100000).toFixed(2)}L
                                </td>
                                <td className="px-4 py-2 text-zinc-600">
                                  {s.franchiseeCount} / ₹{(s.franchiseeValue/100000).toFixed(2)}L
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </Card>

                    {/* Status Breakdown */}
                    <Card className="p-6">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                        <h3 className="text-sm font-bold text-zinc-900 uppercase flex items-center gap-2">
                          <CheckCircle2 className="w-4 h-4 text-orange-600" />
                          Sales Status Breakdown
                        </h3>
                        <div className="flex gap-2">
                          <button
                            onClick={() => setStatsOverviewStatusMetric('count')}
                            className={`px-3 py-1 rounded-full text-[10px] font-bold transition-all ${
                              statsOverviewStatusMetric === 'count'
                                ? 'bg-orange-600 text-white shadow-md'
                                : 'bg-zinc-100 text-zinc-500 hover:bg-zinc-200'
                            }`}
                          >
                            Count
                          </button>
                          <button
                            onClick={() => setStatsOverviewStatusMetric('value')}
                            className={`px-3 py-1 rounded-full text-[10px] font-bold transition-all ${
                              statsOverviewStatusMetric === 'value'
                                ? 'bg-orange-600 text-white shadow-md'
                                : 'bg-zinc-100 text-zinc-500 hover:bg-zinc-200'
                            }`}
                          >
                            Value
                          </button>
                        </div>
                      </div>
                      <div className="h-80 w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={statusData} margin={{ top: 20, right: 30, left: 20, bottom: 5 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="name" fontSize={10} fontWeight="bold" />
                            <YAxis 
                              fontSize={10} 
                              tickFormatter={(val) => {
                                if (statsOverviewStatusMetric === 'value') return `₹${(val/100000).toFixed(2)}L`;
                                return val;
                              }} 
                            />
                            <Tooltip 
                              contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                              formatter={(value: number) => {
                                if (statsOverviewStatusMetric === 'value') return [`₹${value.toLocaleString()}`, 'Value'];
                                return [value.toLocaleString(), 'Count'];
                              }}
                            />
                            <Legend verticalAlign="top" height={36}/>
                            <Bar 
                              dataKey={statsOverviewStatusMetric === 'count' ? 'inhouseCount' : 'inhouseValue'} 
                              name="Inhouse" 
                              stackId="a" 
                              fill="#ea580c" 
                            />
                            <Bar 
                              dataKey={statsOverviewStatusMetric === 'count' ? 'branchCount' : 'branchValue'} 
                              name="Branch" 
                              stackId="a" 
                              fill="#3b82f6" 
                            />
                            <Bar 
                              dataKey={statsOverviewStatusMetric === 'count' ? 'franchiseeCount' : 'franchiseeValue'} 
                              name="Franchisee" 
                              stackId="a" 
                              fill="#10b981" 
                              radius={[4, 4, 0, 0]} 
                            />
                          </BarChart>
                        </ResponsiveContainer>
                      </div>

                      {/* Detailed Status Table */}
                      <div className="mt-8 overflow-x-auto">
                        <table className="w-full text-left text-[10px]">
                          <thead className="bg-zinc-50 text-zinc-500 uppercase font-bold">
                            <tr>
                              <th className="px-4 py-2">Status</th>
                              <th className="px-4 py-2">Count (%)</th>
                              <th className="px-4 py-2">Value (%)</th>
                              <th className="px-4 py-2">Inhouse (C/V)</th>
                              <th className="px-4 py-2">Branch (C/V)</th>
                              <th className="px-4 py-2">Franchisee (C/V)</th>
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-zinc-100">
                            {statusData.map((s: any) => (
                              <tr key={s.name} className="hover:bg-zinc-50 transition-colors">
                                <td className="px-4 py-2 font-bold text-zinc-900">{s.name}</td>
                                <td className="px-4 py-2">
                                  <span className="font-bold">{s.count}</span>
                                  <span className="text-zinc-400 ml-1">({s.countPerc.toFixed(1)}%)</span>
                                </td>
                                <td className="px-4 py-2">
                                  <span className="font-bold">₹{(s.value/100000).toFixed(2)}L</span>
                                  <span className="text-zinc-400 ml-1">({s.valuePerc.toFixed(1)}%)</span>
                                </td>
                                <td className="px-4 py-2 text-zinc-600">
                                  {s.inhouseCount} / ₹{(s.inhouseValue/100000).toFixed(2)}L
                                </td>
                                <td className="px-4 py-2 text-zinc-600">
                                  {s.branchCount} / ₹{(s.branchValue/100000).toFixed(2)}L
                                </td>
                                <td className="px-4 py-2 text-zinc-600">
                                  {s.franchiseeCount} / ₹{(s.franchiseeValue/100000).toFixed(2)}L
                                </td>
                              </tr>
                            ))}
                          </tbody>
                        </table>
                      </div>
                    </Card>
                  </div>

                  {/* Top 5 Destinations by Source */}
                  <Card className="p-6">
                    <div className="flex items-center justify-between mb-6">
                      <h3 className="text-sm font-bold text-zinc-900 uppercase flex items-center gap-2">
                        <MapPin className="w-4 h-4 text-orange-600" />
                        Top 5 Destinations by Source
                      </h3>
                    </div>
                    <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-3 gap-4">
                      {topDestinationsBySource.filter(s => s.top5.length > 0).map((sourceGroup) => (
                        <div key={sourceGroup.source} className="bg-zinc-50 rounded-xl p-4 border border-zinc-100">
                          <h4 className="text-xs font-black text-zinc-900 uppercase mb-3 pb-2 border-b border-zinc-200 flex justify-between items-center">
                            <span>{sourceGroup.source}</span>
                            <span className="text-[10px] text-zinc-400 font-normal">Top 5</span>
                          </h4>
                          <div className="space-y-2">
                            <div className="grid grid-cols-12 text-[12px] font-bold text-zinc-400 uppercase pb-1 border-b border-zinc-100">
                              <div className="col-span-5">Destination</div>
                              <div className="col-span-2 text-center">Sales</div>
                              <div className="col-span-3 text-right">Amount</div>
                              <div className="col-span-2 text-right">Margin</div>
                            </div>
                            {sourceGroup.top5.map((dest: any, idx: number) => (
                              <div key={dest.name} className="grid grid-cols-12 items-center text-[13px] py-1.5 border-b border-zinc-50 last:border-0">
                                <div className="col-span-5 flex items-center gap-2 min-w-0">
                                  <span className="w-5 h-5 flex-shrink-0 flex items-center justify-center bg-zinc-200 text-zinc-600 rounded-full text-[9px] font-bold">
                                    {idx + 1}
                                  </span>
                                  <span className="font-medium text-zinc-700 truncate">{dest.name}</span>
                                </div>
                                <div className="col-span-2 text-center font-bold text-zinc-900">
                                  {dest.count}
                                </div>
                                <div className="col-span-3 text-right font-bold text-orange-600">
                                  ₹{(dest.amount / 100000).toFixed(1)}L
                                </div>
                                <div className="col-span-2 text-right flex flex-col items-end leading-tight">
                                  <span className="font-bold text-emerald-600">₹{(dest.margin / 1000).toFixed(0)}k</span>
                                  <span className="text-[9px] text-emerald-500 font-bold">
                                    {dest.amount > 0 ? ((dest.margin / dest.amount) * 100).toFixed(1) : 0}%
                                  </span>
                                </div>
                              </div>
                            ))}
                          </div>
                        </div>
                      ))}
                    </div>
                  </Card>

                  {/* Package Amount Analysis */}
                  <Card className="p-6">
                      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-4 mb-6">
                        <h3 className="text-sm font-bold text-zinc-900 uppercase flex items-center gap-2">
                          <IndianRupee className="w-4 h-4 text-orange-600" />
                          Package Amount Analysis
                        </h3>
                        <div className="flex items-center gap-2 bg-zinc-100 p-1 rounded-xl">
                          {(['BDE', 'Monthly', 'Weekly', 'Destination', 'Source'] as const).map((dim) => (
                            <button
                              key={dim}
                              onClick={() => setStatsOverviewPackageAnalysisDimension(dim)}
                              className={cn(
                                "px-4 py-1.5 rounded-lg text-[10px] font-bold transition-all",
                                statsOverviewPackageAnalysisDimension === dim 
                                  ? "bg-white text-orange-600 shadow-sm" 
                                  : "text-zinc-500 hover:bg-zinc-200"
                              )}
                            >
                              {dim}
                            </button>
                          ))}
                        </div>
                      </div>
                      <div className="h-[600px] w-full">
                        <ResponsiveContainer width="100%" height="100%">
                          <BarChart data={packageAnalysisData} margin={{ top: 40, right: 30, left: 20, bottom: 5 }}>
                            <CartesianGrid strokeDasharray="3 3" vertical={false} stroke="#f1f5f9" />
                            <XAxis dataKey="range" fontSize={10} fontWeight="bold" />
                            <YAxis fontSize={10} />
                            <Tooltip 
                              contentStyle={{ borderRadius: '12px', border: 'none', boxShadow: '0 10px 15px -3px rgb(0 0 0 / 0.1)' }}
                              formatter={(value: number, name: string, props: any) => {
                                const frameMatch = name.match(/frame(\d+)Count/);
                                const frameIdx = frameMatch ? parseInt(frameMatch[1]) : 0;
                                const frameName = props.payload[`frame${frameIdx}Name`];
                                return [value.toLocaleString(), frameName || name];
                              }}
                            />
                            <Legend 
                              verticalAlign="top" 
                              height={36}
                              formatter={(value, entry: any) => {
                                const frameMatch = value.match(/frame(\d+)Count/);
                                const frameIdx = frameMatch ? parseInt(frameMatch[1]) : 0;
                                return packageAnalysisData[0]?.[`frame${frameIdx}Name`] || value;
                              }}
                            />
                            <Bar dataKey="frame0Count" fill="#1e3a8a" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey="frame0Count" 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  const data = packageAnalysisData[index];
                                  if (!data) return null;
                                  return (
                                    <g>
                                      <text x={x + width / 2} y={y - 20} fill="#1e3a8a" textAnchor="middle" fontSize={10} fontWeight="bold">
                                        {value}
                                      </text>
                                      <text x={x + width / 2} y={y - 8} fill="#1e3a8a" textAnchor="middle" fontSize={8}>
                                        ₹{(data.frame0Value/100000).toFixed(1)}L
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>
                            <Bar dataKey="frame1Count" fill="#1d4ed8" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey="frame1Count" 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  const data = packageAnalysisData[index];
                                  if (!data) return null;
                                  return (
                                    <g>
                                      <text x={x + width / 2} y={y - 20} fill="#1d4ed8" textAnchor="middle" fontSize={10} fontWeight="bold">
                                        {value}
                                      </text>
                                      <text x={x + width / 2} y={y - 8} fill="#1d4ed8" textAnchor="middle" fontSize={8}>
                                        ₹{(data.frame1Value/100000).toFixed(1)}L
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>
                            <Bar dataKey="frame2Count" fill="#3b82f6" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey="frame2Count" 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  const data = packageAnalysisData[index];
                                  if (!data) return null;
                                  return (
                                    <g>
                                      <text x={x + width / 2} y={y - 20} fill="#3b82f6" textAnchor="middle" fontSize={10} fontWeight="bold">
                                        {value}
                                      </text>
                                      <text x={x + width / 2} y={y - 8} fill="#3b82f6" textAnchor="middle" fontSize={8}>
                                        ₹{(data.frame2Value/100000).toFixed(1)}L
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>
                            <Bar dataKey="frame3Count" fill="#60a5fa" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey="frame3Count" 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  const data = packageAnalysisData[index];
                                  if (!data) return null;
                                  return (
                                    <g>
                                      <text x={x + width / 2} y={y - 20} fill="#60a5fa" textAnchor="middle" fontSize={10} fontWeight="bold">
                                        {value}
                                      </text>
                                      <text x={x + width / 2} y={y - 8} fill="#60a5fa" textAnchor="middle" fontSize={8}>
                                        ₹{(data.frame3Value/100000).toFixed(1)}L
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>
                            <Bar dataKey="frame4Count" fill="#93c5fd" radius={[4, 4, 0, 0]}>
                              <LabelList 
                                dataKey="frame4Count" 
                                position="top" 
                                content={(props: any) => {
                                  const { x, y, width, value, index } = props;
                                  const data = packageAnalysisData[index];
                                  if (!data) return null;
                                  return (
                                    <g>
                                      <text x={x + width / 2} y={y - 20} fill="#93c5fd" textAnchor="middle" fontSize={10} fontWeight="bold">
                                        {value}
                                      </text>
                                      <text x={x + width / 2} y={y - 8} fill="#93c5fd" textAnchor="middle" fontSize={8}>
                                        ₹{(data.frame4Value/100000).toFixed(1)}L
                                      </text>
                                    </g>
                                  );
                                }}
                              />
                            </Bar>
                          </BarChart>
                        </ResponsiveContainer>
                      </div>

                      {/* Detailed Package Analysis Table */}
                      <div className="mt-8 overflow-x-auto">
                        <table className="w-full text-left text-[12px] border-collapse">
                          <thead className="bg-zinc-50 text-zinc-500 uppercase font-bold">
                            <tr>
                              <th className="px-4 py-3 border border-zinc-200">Package Range</th>
                              <th className="px-4 py-3 border border-zinc-200">Total (C/V/M%)</th>
                              {comparisonFrames.map((frame, idx) => (
                                <th key={idx} className="px-4 py-3 border border-zinc-200 bg-zinc-100 text-zinc-900">
                                  {frame}
                                </th>
                              ))}
                            </tr>
                          </thead>
                          <tbody className="divide-y divide-zinc-100">
                            {packageAnalysisData.map((p: any) => {
                              const framePercentages = comparisonFrames.map((_, idx) => p[`frame${idx}MarginPerc`] || 0);
                              const sortedPercentages = [...framePercentages].sort((a, b) => a - b);
                              
                              const getBgColor = (perc: number) => {
                                const rank = sortedPercentages.indexOf(perc);
                                // Lowest (rank 0) -> darkest, Highest (rank 4) -> lightest
                                if (rank === 0) return 'bg-blue-900 text-white';
                                if (rank === 1) return 'bg-blue-700 text-white';
                                if (rank === 2) return 'bg-blue-500 text-white';
                                if (rank === 3) return 'bg-blue-300 text-zinc-900';
                                return 'bg-blue-100 text-zinc-900';
                              };

                              return (
                                <tr key={p.range} className="hover:bg-zinc-50 transition-colors">
                                  <td className="px-4 py-3 font-bold text-zinc-900 border border-zinc-200">{p.range}</td>
                                  <td className="px-4 py-3 border border-zinc-200">
                                    <div className="flex flex-col">
                                      <span className="font-bold">{p.count} Sales</span>
                                      <span className="text-[10px] text-zinc-500 italic">(In: {p.totalInhouseCount} / Br: {p.totalBranchCount} / Fr: {p.totalFranchiseeCount})</span>
                                      <span className="text-zinc-600">₹{(p.value/100000).toFixed(2)}L</span>
                                      <span className="text-emerald-600 font-bold">{p.marginPerc.toFixed(1)}% Margin</span>
                                    </div>
                                  </td>
                                  {comparisonFrames.map((_, idx) => {
                                    const count = p[`frame${idx}Count`];
                                    const value = p[`frame${idx}Value`];
                                    const marginPerc = p[`frame${idx}MarginPerc`];
                                    const inhouse = p[`frame${idx}InhouseCount`];
                                    const branch = p[`frame${idx}BranchCount`];
                                    const franchisee = p[`frame${idx}FranchiseeCount`];
                                    const bgColorClass = getBgColor(marginPerc);

                                    return (
                                      <td key={idx} className={cn("px-4 py-3 border border-zinc-200", bgColorClass)}>
                                        <div className="flex flex-col">
                                          <span className="font-bold">{count} Sales</span>
                                          <span className="text-[10px] opacity-80 italic">(In: {inhouse} / Br: {branch} / Fr: {franchisee})</span>
                                          <span className="opacity-90">₹{(value/100000).toFixed(2)}L</span>
                                          <span className="font-bold">{marginPerc.toFixed(1)}% M</span>
                                        </div>
                                      </td>
                                    );
                                  })}
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    </Card>
                </div>
              );
            })()}
          </div>
        )}

        {activeTab === 'issueOverview' && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <AlertCircle className="w-6 h-6 text-red-600" />
                <h2 className="text-3xl font-bold tracking-tight">Issue Overview</h2>
              </div>
            </div>

            <div className="flex items-center gap-2 bg-white p-1 rounded-xl border border-zinc-200 w-fit">
              {(['All', 'Hotel', 'Flight', 'Land', 'HFL', 'Work Pending', 'Credit Note'] as const).map((tab) => (
                <button
                  key={tab}
                  onClick={() => setIssueOverviewActiveTab(tab)}
                  className={cn(
                    "px-4 py-2 rounded-lg text-sm font-bold transition-all flex items-center gap-2",
                    issueOverviewActiveTab === tab 
                      ? "bg-orange-600 text-white shadow-md" 
                      : "text-zinc-600 hover:bg-zinc-50"
                  )}
                >
                  <span>{tab}</span>
                  <span className={cn(
                    "px-1.5 py-0.5 rounded-md text-[10px] font-bold",
                    issueOverviewActiveTab === tab 
                      ? "bg-white/20 text-white" 
                      : "bg-zinc-100 text-zinc-500"
                  )}>
                    {issueOverviewCounts[tab]}
                  </span>
                </button>
              ))}
            </div>

            <Card className="p-0 overflow-hidden border-zinc-200 shadow-xl">
              <div className="overflow-x-auto">
                <table className="w-full text-left border-collapse">
                  <thead className="bg-zinc-100 text-zinc-600 uppercase font-bold text-[10px] tracking-wider border-b border-zinc-200">
                    <tr>
                      <th className="px-4 py-4 border-r border-zinc-200 text-center">Log</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Date</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Trip ID</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Guest Name</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Agent</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Destination</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Value</th>
                      <th className="px-4 py-4 border-r border-zinc-200">P-Hotel</th>
                      <th className="px-4 py-4 border-r border-zinc-200">P-Flight</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Land</th>
                      <th className="px-4 py-4 border-r border-zinc-200">HFL Issue</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Work Pending</th>
                      <th className="px-4 py-4 border-r border-zinc-200">Status</th>
                      <th className="px-4 py-4">Trip Date</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-zinc-200">
                    {filteredIssueSales.map((sale) => (
                      <tr key={sale.id} className="hover:bg-zinc-50 transition-colors group">
                        <td className="px-4 py-4 border-r border-zinc-200 text-center">
                          <button
                            onClick={() => handleOpenRemarks(sale)}
                            className={cn(
                              "inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border transition-all group/btn",
                              isRemarkUnread(sale) 
                                ? "bg-orange-600 border-orange-600 text-white shadow-lg shadow-orange-600/20" 
                                : "border-zinc-200 hover:bg-white hover:border-orange-300 hover:text-orange-600"
                            )}
                          >
                            <MessageSquare className={cn(
                              "w-3.5 h-3.5",
                              isRemarkUnread(sale) ? "text-white" : "text-zinc-400 group-hover/btn:text-orange-500"
                            )} />
                            <span className={cn(
                              "text-xs font-bold",
                              isRemarkUnread(sale) ? "text-white" : "text-zinc-900 group-hover/btn:text-orange-600"
                            )}>{sale.remarksCount || 0}</span>
                          </button>
                        </td>
                        <td className="px-4 py-4 text-sm text-zinc-600 font-medium whitespace-nowrap border-r border-zinc-200">{sale.date}</td>
                        <td className="px-4 py-4 text-sm font-bold text-zinc-900 border-r border-zinc-200">{sale.tripId}</td>
                        <td className="px-4 py-4 text-sm text-zinc-600 font-medium border-r border-zinc-200">{sale.guestName}</td>
                        <td className="px-4 py-4 text-sm text-zinc-600 font-medium border-r border-zinc-200">{sale.agent}</td>
                        <td className="px-4 py-4 text-sm text-zinc-600 font-medium capitalize border-r border-zinc-200">{sale.destination}</td>
                        <td className="px-4 py-4 text-sm font-bold text-zinc-900 whitespace-nowrap border-r border-zinc-200">
                          ₹{(sale.packageValue || 0).toLocaleString()}
                        </td>
                        <td className="px-4 py-4 border-r border-zinc-200">
                          <span className={cn(
                            "text-xs font-bold px-2 py-1 rounded-md",
                            sale.pHotel === 'Issue' ? "bg-red-50 text-red-600" : "bg-zinc-50 text-zinc-500"
                          )}>
                            {sale.pHotel || 'Done'}
                          </span>
                        </td>
                        <td className="px-4 py-4 border-r border-zinc-200">
                          <span className={cn(
                            "text-xs font-bold px-2 py-1 rounded-md",
                            sale.pFlight === 'Issue' ? "bg-red-50 text-red-600" : "bg-zinc-50 text-zinc-500"
                          )}>
                            {sale.pFlight || 'Done'}
                          </span>
                        </td>
                        <td className="px-4 py-4 border-r border-zinc-200">
                          <span className={cn(
                            "text-xs font-bold px-2 py-1 rounded-md",
                            sale.land === 'Issue' ? "bg-red-50 text-red-600" : "bg-zinc-50 text-zinc-500"
                          )}>
                            {sale.land || 'Done'}
                          </span>
                        </td>
                        <td className="px-4 py-4 border-r border-zinc-200">
                          <span className={cn(
                            "text-xs font-bold px-2 py-1 rounded-md",
                            sale.hflIssue === 'Issue' ? "bg-red-50 text-red-600" : "bg-zinc-50 text-zinc-500"
                          )}>
                            {sale.hflIssue || 'Done'}
                          </span>
                        </td>
                        <td className="px-4 py-4 border-r border-zinc-200">
                          <span className={cn(
                            "text-xs font-bold px-2 py-1 rounded-md",
                            sale.workPending === 'Issue' ? "bg-red-50 text-red-600" : "bg-zinc-50 text-zinc-500"
                          )}>
                            {sale.workPending || 'Done'}
                          </span>
                        </td>
                        <td className="px-4 py-4 border-r border-zinc-200">
                          <Badge className={cn(
                            "text-[10px] px-2 py-0.5",
                            sale.advanceCN === 'Done' ? "bg-green-100 text-green-700" :
                            sale.advanceCN === 'Confirmed' ? "bg-blue-100 text-blue-700" :
                            sale.advanceCN === 'Cancel' ? "bg-red-100 text-red-700" :
                            "bg-orange-100 text-orange-700"
                          )}>
                            {sale.advanceCN || 'Pending'}
                          </Badge>
                        </td>
                        <td className="px-4 py-4 text-sm text-zinc-600 font-bold whitespace-nowrap">{sale.tripDate}</td>
                      </tr>
                    ))}
                    {filteredIssueSales.length === 0 && (
                      <tr>
                        <td colSpan={14} className="px-6 py-12 text-center text-zinc-500 font-medium">
                          No issues found for the selected filters.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </Card>

            <div className="pt-8 border-t border-zinc-100">
              <Last7DaysSalesOverview sales={sales} employees={employees} />
            </div>

            <div className="pt-8 border-t border-zinc-100">
                <div className="flex items-center gap-2">
                  <TrendingUp className="w-6 h-6 text-orange-600" />
                  <h3 className="text-2xl font-bold tracking-tight text-zinc-900">Sales Overview</h3>
                </div>
                <div className="flex items-center gap-2">
                  <Filter className="w-4 h-4 text-zinc-400" />
                  <select 
                    value={issueOverviewWeekFilter}
                    onChange={(e) => setIssueOverviewWeekFilter(e.target.value)}
                    className="bg-white border border-zinc-200 rounded-lg px-3 py-1.5 text-sm focus:ring-2 focus:ring-orange-500 outline-none"
                  >
                    <option value="All">All (Last 2 Weeks)</option>
                    <option value="Current">Current Week ({currentWeekName})</option>
                    <option value="Last">Last Week ({lastWeekName})</option>
                  </select>
                </div>
              </div>
              
              <Card className="p-0 overflow-hidden border-zinc-200 shadow-xl">
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead className="bg-zinc-100 text-zinc-600 uppercase font-bold text-[10px] tracking-wider border-b border-zinc-200">
                      <tr>
                        <th className="px-4 py-4 border-r border-zinc-200 text-center">Log</th>
                        <th className="px-4 py-4 border-r border-zinc-200">Date</th>
                        <th className="px-4 py-4 border-r border-zinc-200">Trip ID</th>
                        <th className="px-4 py-4 border-r border-zinc-200">Guest Name</th>
                        <th className="px-4 py-4 border-r border-zinc-200">Agent</th>
                        <th className="px-4 py-4 border-r border-zinc-200">BDE</th>
                        <th className="px-4 py-4 border-r border-zinc-200">Destination</th>
                        <th className="px-4 py-4 border-r border-zinc-200 text-right">Value</th>
                        <th className="px-4 py-4 border-r border-zinc-200 text-center">Status</th>
                        <th className="px-4 py-4">Trip Date</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-zinc-200">
                      {filteredIssueOverviewAllSales.map((sale) => (
                        <tr key={sale.id} className="hover:bg-zinc-50 transition-colors group">
                          <td className="px-4 py-4 border-r border-zinc-200 text-center">
                            <button
                              onClick={() => handleOpenRemarks(sale)}
                              className={cn(
                                "inline-flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border transition-all group/btn",
                                isRemarkUnread(sale) 
                                  ? "bg-orange-600 border-orange-600 text-white shadow-lg shadow-orange-600/20" 
                                  : "border-zinc-200 hover:bg-white hover:border-orange-300 hover:text-orange-600"
                              )}
                            >
                              <MessageSquare className={cn(
                                "w-3.5 h-3.5",
                                isRemarkUnread(sale) ? "text-white" : "text-zinc-400 group-hover/btn:text-orange-500"
                              )} />
                              <span className={cn(
                                "text-xs font-bold",
                                isRemarkUnread(sale) ? "text-white" : "text-zinc-900 group-hover/btn:text-orange-600"
                              )}>{sale.remarksCount || 0}</span>
                            </button>
                          </td>
                          <td className="px-4 py-4 text-sm text-zinc-600 font-medium whitespace-nowrap border-r border-zinc-200">{sale.date}</td>
                          <td className="px-4 py-4 text-sm font-bold text-zinc-900 border-r border-zinc-200">{sale.tripId}</td>
                          <td className="px-4 py-4 text-sm text-zinc-600 font-medium border-r border-zinc-200">{sale.guestName}</td>
                          <td className="px-4 py-4 text-sm text-zinc-600 font-medium border-r border-zinc-200">{sale.agent}</td>
                          <td className="px-4 py-4 text-sm text-zinc-600 font-medium border-r border-zinc-200">{sale.bde}</td>
                          <td className="px-4 py-4 text-sm text-zinc-600 font-medium capitalize border-r border-zinc-200">{sale.destination}</td>
                          <td className="px-4 py-4 text-sm font-bold text-zinc-900 whitespace-nowrap border-r border-zinc-200 text-right">
                            ₹{(sale.packageValue || 0).toLocaleString()}
                          </td>
                          <td className="px-4 py-4 border-r border-zinc-200 text-center">
                            <Badge className={cn(
                              "text-[10px] px-2 py-0.5",
                              sale.advanceCN === 'Done' ? "bg-green-100 text-green-700" :
                              sale.advanceCN === 'Confirmed' ? "bg-blue-100 text-blue-700" :
                              sale.advanceCN === 'Cancel' ? "bg-red-100 text-red-700" :
                              "bg-orange-100 text-orange-700"
                            )}>
                              {sale.advanceCN || 'Pending'}
                            </Badge>
                          </td>
                          <td className="px-4 py-4 text-sm text-zinc-600 font-bold whitespace-nowrap">{sale.tripDate}</td>
                        </tr>
                      ))}
                      {filteredIssueOverviewAllSales.length === 0 && (
                        <tr>
                          <td colSpan={10} className="px-6 py-12 text-center text-zinc-500 font-medium">
                            No sales found for the selected filters.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </Card>
          </div>
        )}

        {activeTab === 'userManagement' && isAdmin && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Shield className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">User Management</h2>
              </div>
              <div className="flex items-center gap-2">
                <Button 
                  onClick={() => setIsRolesModalOpen(true)} 
                  variant="outline"
                  className="border-zinc-200"
                >
                  <Shield className="w-4 h-4" />
                  Manage Roles
                </Button>
                <Button onClick={handleCreateUser} className="bg-orange-600">
                  <Plus className="w-4 h-4" />
                  Add User
                </Button>
              </div>
            </div>

            <Card className="p-6 bg-zinc-900 text-white border-none shadow-xl">
              <div className="flex items-center justify-between">
                <div>
                  <h3 className="text-lg font-bold flex items-center gap-2">
                    <Database className="w-5 h-5 text-orange-400" />
                    System Maintenance
                  </h3>
                  <p className="text-zinc-400 text-sm mt-1">
                    Manage server-side aggregations and database health.
                  </p>
                </div>
                <div className="flex items-center gap-3">
                  <Button 
                    variant="outline"
                    onClick={refreshAllData}
                    isLoading={isRefreshingStatic}
                    className="border-zinc-700 text-zinc-300 hover:bg-zinc-800"
                  >
                    <RefreshCw className={cn("w-4 h-4 mr-2", isRefreshingStatic && "animate-spin")} />
                    Refresh System Data
                  </Button>
                  <Button 
                    onClick={async () => {
                      if (window.confirm('Are you sure you want to backfill all sales summaries? This will scan all sales records.')) {
                        try {
                          await backfillSalesSummaries();
                          alert('Backfill completed successfully!');
                        } catch (err: any) {
                          alert('Backfill failed: ' + err.message);
                        }
                      }
                    }}
                    className="bg-orange-600 hover:bg-orange-700 text-white border-none"
                  >
                    <RefreshCw className="w-4 h-4 mr-2" />
                    Backfill Sales Summaries
                  </Button>
                </div>
              </div>
            </Card>

            <div className="flex items-center justify-between bg-white p-4 rounded-2xl border border-zinc-200 shadow-sm">
              <div className="flex items-center gap-4">
                <div className="flex items-center gap-2">
                  <Filter className="w-4 h-4 text-zinc-400" />
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by Role</label>
                  <select 
                    className="px-3 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-orange-500 font-medium"
                    value={userRoleFilter}
                    onChange={(e) => setUserRoleFilter(e.target.value)}
                  >
                    <option value="All">All Roles</option>
                    <option value="Agent">Agent (Default)</option>
                    {roles.filter(r => r.name !== 'Agent').map(r => (
                      <option key={r.id} value={r.name}>{r.name}</option>
                    ))}
                  </select>
                </div>
              </div>
              <div className="text-xs text-zinc-400 font-medium italic">
                Showing {allUsers.filter(u => (u.role || 'Agent') === userRoleFilter || userRoleFilter === 'All').length} users
              </div>
            </div>

            <Card className="overflow-x-auto">
              <table className="w-full text-left border-collapse border border-zinc-200">
                <thead>
                  <tr className="bg-zinc-50">
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">User</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Email</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Matched</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Role</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Permissions</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200 text-center">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {allUsers.filter(u => {
                    const role = u.role || 'Agent';
                    return userRoleFilter === 'All' ? true : role === userRoleFilter;
                  }).map(u => (
                    <tr key={u.uid} className="hover:bg-zinc-50 transition-colors">
                      <td className="px-4 py-3 border border-zinc-200">
                        <div className="flex items-center gap-3">
                          <img src={u.photoURL} className="w-8 h-8 rounded-full" alt="" referrerPolicy="no-referrer" />
                          <span className="text-sm font-bold">{u.displayName}</span>
                        </div>
                      </td>
                      <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{u.email}</td>
                      <td className="px-4 py-3 border border-zinc-200">
                        {employees.some(emp => emp.email?.toLowerCase() === u.email?.toLowerCase()) ? (
                          <Badge className="bg-green-100 text-green-700">Matched</Badge>
                        ) : (
                          <Badge className="bg-red-100 text-red-700">No match found</Badge>
                        )}
                      </td>
                      <td className="px-4 py-3 border border-zinc-200">
                        <Badge className={
                          u.role?.toLowerCase() === 'admin' ? 'bg-purple-100 text-purple-700' : 
                          u.role?.toLowerCase() === 'manager' ? 'bg-blue-100 text-blue-700' : 
                          u.role?.toLowerCase() === 'extra team' ? 'bg-orange-100 text-orange-700' :
                          u.role?.toLowerCase() === 'trainee' ? 'bg-green-100 text-green-700' :
                          'bg-zinc-100 text-zinc-500'
                        }>
                          {u.role || 'Agent'}
                        </Badge>
                      </td>
                      <td className="px-4 py-3 border border-zinc-200">
                        <div className="flex flex-wrap gap-1">
                          {u.permissions && typeof u.permissions === 'object' && !Array.isArray(u.permissions) && Object.keys(u.permissions).length > 0 ? (
                            Object.entries(u.permissions).map(([pId, level]) => {
                              const p = allPermissions.find(ap => ap.id === pId);
                              return (
                                <Badge key={pId} className={cn(
                                  "lowercase text-[9px] flex items-center gap-1",
                                  level === 'Complete' ? "bg-green-50 text-green-600 border border-green-100" : "bg-orange-50 text-orange-600 border border-orange-100"
                                )}>
                                  {p?.label || pId}
                                  <span className="opacity-60 font-black">({level})</span>
                                </Badge>
                              );
                            })
                          ) : null}
                        </div>
                      </td>
                      <td className="px-4 py-3 border border-zinc-200">
                        <div className="flex items-center justify-center gap-2">
                          <button 
                            onClick={() => handleEditUser(u)}
                            className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                            title="Edit User"
                          >
                            <Edit2 className="w-4 h-4" />
                          </button>
                          <button 
                            onClick={() => handleDeleteUser(u.uid, u.displayName || u.email)}
                            className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                            title="Delete User"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>

            <Modal 
              isOpen={isUserModalOpen} 
              onClose={() => setIsUserModalOpen(false)}
              title={isEditingUser ? 'Edit User' : 'Add New User'}
            >
              <form onSubmit={handleUpdateUser} className="space-y-6">
                <div className="space-y-4">
                  {isEditingUser ? (
                    <div className="flex items-center gap-4 p-4 bg-zinc-50 rounded-2xl">
                      <img src={userForm.photoURL} className="w-12 h-12 rounded-full" alt="" referrerPolicy="no-referrer" />
                      <div>
                        <h4 className="font-bold">{userForm.displayName}</h4>
                        <p className="text-sm text-zinc-500">{userForm.email}</p>
                      </div>
                    </div>
                  ) : (
                    <>
                      <div className="space-y-2">
                        <label className="text-xs font-bold text-zinc-500 uppercase">Email Address</label>
                        <input 
                          type="email" 
                          required
                          placeholder="e.g. user@example.com"
                          className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                          value={userForm.email || ''}
                          onChange={(e) => setUserForm({...userForm, email: e.target.value})}
                        />
                      </div>
                      <div className="space-y-2">
                        <label className="text-xs font-bold text-zinc-500 uppercase">Display Name</label>
                        <input 
                          type="text" 
                          required
                          placeholder="e.g. John Doe"
                          className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                          value={userForm.displayName || ''}
                          onChange={(e) => setUserForm({...userForm, displayName: e.target.value})}
                        />
                      </div>
                    </>
                  )}

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Role</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={userForm.role}
                      onChange={(e) => {
                        const roleName = e.target.value;
                        const roleObj = roles.find(r => r.name === roleName);
                        setUserForm({
                          ...userForm, 
                          role: roleName,
                          permissions: roleObj ? roleObj.permissions : {}
                        });
                      }}
                    >
                      <option value="">Select Role</option>
                      {roles.map(r => (
                        <option key={r.id} value={r.name}>{r.name}</option>
                      ))}
                    </select>
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Permissions (Auto-assigned by Role)</label>
                    <div className="grid grid-cols-1 gap-2 p-4 bg-zinc-50 rounded-2xl border border-zinc-200 max-h-48 overflow-y-auto">
                      {userForm.permissions && typeof userForm.permissions === 'object' && !Array.isArray(userForm.permissions) && Object.keys(userForm.permissions).length > 0 ? (
                        Object.entries(userForm.permissions).map(([pId, level]) => {
                          const p = allPermissions.find(ap => ap.id === pId);
                          return (
                            <div key={pId} className="flex items-center justify-between text-xs bg-white p-2 rounded-lg border border-zinc-100">
                              <div className="flex items-center gap-2">
                                <div className={cn(
                                  "w-2 h-2 rounded-full",
                                  level === 'Complete' ? "bg-green-500" : "bg-orange-400"
                                )} />
                                <span className="text-zinc-600 font-medium">{p?.label || pId}</span>
                              </div>
                              <Badge className={cn(
                                "text-[9px] uppercase font-black",
                                level === 'Complete' ? "bg-green-100 text-green-700" : "bg-orange-100 text-orange-700"
                              )}>{level}</Badge>
                            </div>
                          );
                        })
                      ) : (
                        <p className="text-xs text-zinc-400 italic text-center py-4">No permissions assigned to this role</p>
                      )}
                    </div>
                  </div>
                </div>

                {error && <p className="text-xs text-red-500">{error}</p>}
                {success && <p className="text-xs text-green-500">{success}</p>}

                <div className="flex gap-3 pt-4">
                  <Button type="submit" className="flex-1" isLoading={isSavingUser}>
                    {isEditingUser ? 'Update User' : 'Create User'}
                  </Button>
                  <Button 
                    type="button" 
                    variant="ghost" 
                    onClick={() => setIsUserModalOpen(false)}
                  >
                    Cancel
                  </Button>
                </div>
              </form>
            </Modal>

            {/* Roles Management Modal */}
            <Modal
              isOpen={isRolesModalOpen}
              onClose={() => setIsRolesModalOpen(false)}
              title="Role Management"
            >
              <div className="space-y-6">
                <div className="flex justify-end">
                  <Button 
                    onClick={() => {
                      setSelectedRole(null);
                      setIsRoleEditModalOpen(true);
                    }}
                    className="bg-orange-600"
                  >
                    <Plus className="w-4 h-4" />
                    New Role
                  </Button>
                </div>

                <div className="space-y-3">
                  {roles.map(role => (
                    <div key={role.id} className="flex items-center justify-between p-4 bg-zinc-50 rounded-2xl border border-zinc-200">
                      <div>
                        <h4 className="font-bold">{role.name}</h4>
                        <div className="flex items-center gap-2">
                          <p className="text-xs text-zinc-500">{Object.keys(role.permissions).length} permissions</p>
                          <span className="text-zinc-300">•</span>
                          <p className="text-xs text-orange-600 font-medium">Default: {allPermissions.find(p => p.id === role.defaultPage)?.label || 'Dashboard'}</p>
                        </div>
                      </div>
                      <div className="flex items-center gap-2">
                        <button 
                          onClick={() => {
                            setSelectedRole(role);
                            setIsRoleEditModalOpen(true);
                          }}
                          className="p-1.5 hover:bg-orange-100 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                        >
                          <Edit2 className="w-4 h-4" />
                        </button>
                        <button 
                          onClick={() => handleDeleteRole(role.id, role.name)}
                          className="p-1.5 hover:bg-red-100 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>
            </Modal>

            {/* Role Edit Modal */}
            <Modal
              isOpen={isRoleEditModalOpen}
              onClose={() => setIsRoleEditModalOpen(false)}
              title={selectedRole ? 'Edit Role' : 'Create Role'}
            >
              <form 
                onSubmit={async (e) => {
                  e.preventDefault();
                  const formData = new FormData(e.currentTarget);
                  const name = formData.get('name') as string;
                  const defaultPage = formData.get('defaultPage') as string;
                  
                  const selectedPerms: Record<string, 'Limited' | 'Complete'> = {};
                  allPermissions.forEach(p => {
                    const level = formData.get(`perm-${p.id}`) as 'None' | 'Limited' | 'Complete';
                    if (level && level !== 'None') {
                      selectedPerms[p.id] = level;
                    }
                  });

                  try {
                    if (selectedRole) {
                      await updateDoc(doc(db, 'roles', selectedRole.id), {
                        name,
                        defaultPage,
                        permissions: selectedPerms
                      });
                      
                      // Sync all users with this role
                      const usersToUpdate = allUsers.filter(u => u.role === selectedRole.name);
                      for (const u of usersToUpdate) {
                        await updateDoc(doc(db, 'users', u.uid), {
                          permissions: selectedPerms
                        });
                      }
                    } else {
                      await addDoc(collection(db, 'roles'), {
                        name,
                        defaultPage,
                        permissions: selectedPerms,
                        createdAt: serverTimestamp()
                      });
                    }
                    setIsRoleEditModalOpen(false);
                  } catch (err) {
                    console.error("Error saving role:", err);
                  }
                }}
                className="space-y-6"
              >
                <div className="space-y-4">
                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Role Name</label>
                    <input 
                      name="name"
                      type="text" 
                      required
                      defaultValue={selectedRole?.name || ''}
                      placeholder="e.g. Supervisor"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    />
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Default Landing Page</label>
                    <select 
                      name="defaultPage"
                      defaultValue={selectedRole?.defaultPage || 'dashboard'}
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none font-medium"
                    >
                      {allPermissions.map(p => (
                        <option key={p.id} value={p.id}>{p.label}</option>
                      ))}
                    </select>
                    <p className="text-[10px] text-zinc-400 px-2 italic">
                      * This page will be shown when users with this role log in.
                    </p>
                  </div>

                  <div className="space-y-2">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Permissions & Access Level</label>
                    <div className="space-y-2 p-4 bg-zinc-50 rounded-2xl border border-zinc-200 max-h-80 overflow-y-auto">
                      {allPermissions.map(p => (
                        <div key={p.id} className="flex items-center justify-between p-3 bg-white border border-zinc-100 rounded-xl shadow-sm">
                          <span className="text-sm font-medium text-zinc-700">{p.label}</span>
                          <select 
                            name={`perm-${p.id}`}
                            defaultValue={selectedRole?.permissions?.[p.id] || 'None'}
                            className="text-xs px-2 py-1.5 bg-zinc-50 border border-zinc-200 rounded-lg outline-none focus:ring-2 focus:ring-orange-500 font-bold"
                          >
                            <option value="None">No Access</option>
                            <option value="Limited">Limited</option>
                            <option value="Complete">Complete</option>
                          </select>
                        </div>
                      ))}
                    </div>
                    <p className="text-[10px] text-zinc-400 px-2 italic">
                      * Limited: View own data only. Complete: Full access to all data.
                    </p>
                  </div>
                </div>

                <div className="flex justify-end gap-3 pt-4 border-t border-zinc-100">
                  <Button type="button" variant="outline" onClick={() => setIsRoleEditModalOpen(false)}>
                    Cancel
                  </Button>
                  <Button type="submit" className="bg-orange-600">
                    {selectedRole ? 'Update Role' : 'Create Role'}
                  </Button>
                </div>
              </form>
            </Modal>
          </div>
        )}
        {activeTab === 'trainingOverview' && hasPermission('trainingOverview') && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <h2 className="text-xl font-bold">Training Overview</h2>
              <div className="flex items-center gap-4">
                <div className="flex items-center gap-2 min-w-[200px]">
                  <label className="text-xs font-bold text-zinc-500 uppercase whitespace-nowrap">Agent</label>
                  <SearchableSelect
                    options={[
                      { value: '', label: 'All Agents' },
                      ...employees
                        .filter(e => e.status === 'Active')
                        .sort((a, b) => a.name.localeCompare(b.name))
                        .map(e => ({ value: e.name, label: e.name, key: e.id }))
                    ]}
                    value={trainingAgentFilter}
                    onChange={setTrainingAgentFilter}
                    placeholder="All Agents"
                  />
                </div>
                <div className="flex items-center gap-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Module</label>
                  <select 
                    className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-orange-500"
                    value={trainingModuleFilter}
                    onChange={(e) => setTrainingModuleFilter(e.target.value)}
                  >
                    <option value="">All Modules</option>
                    <option value="Sales">Sales</option>
                    <option value="Operations">Operations</option>
                    <option value="Lead">Lead</option>
                  </select>
                </div>
                <div className="flex items-center gap-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Emp. Status</label>
                  <select 
                    className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-orange-500"
                    value={trainingEmployeeStatusFilter}
                    onChange={(e) => setTrainingEmployeeStatusFilter(e.target.value as any)}
                  >
                    <option value="All">All</option>
                    <option value="Active">Active</option>
                    <option value="Deactive">Deactive</option>
                  </select>
                </div>
                <div className="flex items-center gap-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Tr. Status</label>
                  <select 
                    className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-orange-500"
                    value={trainingStatusFilter}
                    onChange={(e) => setTrainingStatusFilter(e.target.value as any)}
                  >
                    <option value="All">All Status</option>
                    <option value="Active">Active</option>
                    <option value="Complete">Complete</option>
                  </select>
                </div>
                <Button 
                  variant="ghost" 
                  onClick={() => { 
                    setTrainingAgentFilter(''); 
                    setTrainingModuleFilter(''); 
                    setTrainingStatusFilter('Active');
                    setTrainingEmployeeStatusFilter('Active');
                  }}
                  className="text-zinc-500"
                >
                  <Filter className="w-4 h-4 mr-2" />
                  Clear
                </Button>
              </div>
            </div>

            {/* Total Status Bar */}
            {(() => {
              const activeTMs = trainingMaterials.filter(tm => tm.isActive);
              const filteredEmployees = employees.filter(e => {
                const matchesStatus = trainingEmployeeStatusFilter === 'All' ? true : e.status === trainingEmployeeStatusFilter;
                const level = getPermissionLevel('trainingOverview');
                const canSee = level === 'Complete' || e.email?.toLowerCase() === profile?.email?.toLowerCase();
                return matchesStatus && canSee;
              });
              
              let totalPossiblePoints = 0;
              let totalAchievedPoints = 0;
              let totalTasks = 0;
              const uniqueAgents = new Set<string>();

              const moduleCounts = {
                Sales: 0,
                Lead: 0,
                Operations: 0
              };

              filteredEmployees.forEach(emp => {
                const empModules = (emp.completedTrainings || []) as any[];
                
                // Count employees per module (regardless of training status filter for this specific metric)
                empModules.forEach(mod => {
                  if (mod in moduleCounts) {
                    moduleCounts[mod as keyof typeof moduleCounts]++;
                  }
                });

                empModules.forEach(mod => {
                  const status = emp.trainingModuleStatus?.[mod] || 'Active';
                  if (trainingStatusFilter !== 'All' && status !== trainingStatusFilter) return;
                  if (trainingModuleFilter && mod !== trainingModuleFilter) return;
                  if (trainingAgentFilter && emp.name !== trainingAgentFilter) return;

                  uniqueAgents.add(emp.id!);

                  const moduleTMs = activeTMs.filter(tm => {
                    const tmMods = Array.isArray(tm.module) ? tm.module : [tm.module].filter(Boolean);
                    const expandedMods = tmMods.flatMap(m => m === 'Both' ? ['Sales', 'Operations'] : [m]);
                    return expandedMods.includes(mod);
                  });

                  totalTasks += moduleTMs.length;
                  totalPossiblePoints += moduleTMs.reduce((s, tm) => s + (tm.points || 0), 0);
                  const pointsMap = emp.trainingPoints || {};
                  totalAchievedPoints += moduleTMs.reduce((s, tm) => s + (pointsMap[tm.id!] || 0), 0);
                });
              });

              const overallCompletion = totalPossiblePoints > 0 ? Math.round((totalAchievedPoints / totalPossiblePoints) * 100) : 0;

              return (
                <div className="space-y-4">
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                    <Card className="p-4 bg-zinc-900 text-white border-none">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <Users className="w-5 h-5 text-blue-400" />
                        </div>
                        <div>
                          <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">No. of Agents</p>
                          <p className="text-2xl font-bold">{uniqueAgents.size}</p>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-zinc-900 text-white border-none">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <CheckCircle2 className="w-5 h-5 text-green-400" />
                        </div>
                        <div>
                          <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">No. of Tasks</p>
                          <p className="text-2xl font-bold">{totalTasks}</p>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-zinc-900 text-white border-none">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <Award className="w-5 h-5 text-orange-400" />
                        </div>
                        <div>
                          <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Total Points</p>
                          <p className="text-2xl font-bold">{totalPossiblePoints}</p>
                        </div>
                      </div>
                    </Card>
                    <Card className="p-4 bg-zinc-900 text-white border-none">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <Trophy className="w-5 h-5 text-yellow-400" />
                        </div>
                        <div>
                          <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Points Achieved</p>
                          <p className="text-2xl font-bold">{totalAchievedPoints}</p>
                        </div>
                      </div>
                    </Card>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <Card className="p-4 bg-white border border-zinc-200">
                      <p className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider mb-1">Sales Module Employees</p>
                      <p className="text-xl font-bold text-zinc-900">{moduleCounts.Sales}</p>
                    </Card>
                    <Card className="p-4 bg-white border border-zinc-200">
                      <p className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider mb-1">Lead Module Employees</p>
                      <p className="text-xl font-bold text-zinc-900">{moduleCounts.Lead}</p>
                    </Card>
                    <Card className="p-4 bg-white border border-zinc-200">
                      <p className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider mb-1">Operations Module Employees</p>
                      <p className="text-xl font-bold text-zinc-900">{moduleCounts.Operations}</p>
                    </Card>
                  </div>

                  <Card className="p-6 bg-zinc-900 text-white border-none">
                    <div className="flex items-center justify-between mb-4">
                      <div className="flex items-center gap-3">
                        <div className="p-2 bg-zinc-800 rounded-lg">
                          <Award className="w-6 h-6 text-orange-400" />
                        </div>
                        <div>
                          <p className="text-xs font-bold text-zinc-400 uppercase tracking-wider">Overall Training Completion</p>
                          <p className="text-3xl font-bold text-white">{overallCompletion}%</p>
                        </div>
                      </div>
                    </div>
                    <div className="w-full bg-zinc-800 h-4 rounded-full overflow-hidden">
                      <motion.div 
                        initial={{ width: 0 }}
                        animate={{ width: `${overallCompletion}%` }}
                        transition={{ duration: 1, ease: "easeOut" }}
                        className="bg-gradient-to-r from-orange-600 to-orange-400 h-full"
                      />
                    </div>
                  </Card>
                </div>
              );
            })()}

            <Card className="overflow-x-auto">
              <table className="w-full text-left border-collapse border border-zinc-200">
                <thead>
                  <tr className="bg-zinc-50">
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Agent name</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Joining Date</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Module</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Total Points</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Total tasks</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Total Task Done</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Total points achieved</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Total Task Pending</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Status</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Update</th>
                  </tr>
                </thead>
                <tbody>
                  {employees
                    .filter(e => trainingEmployeeStatusFilter === 'All' ? true : e.status === trainingEmployeeStatusFilter)
                    .filter(e => !trainingAgentFilter || e.name === trainingAgentFilter)
                    .sort((a, b) => a.name.localeCompare(b.name))
                    .flatMap(agent => {
                      const modules: ('Sales' | 'Operations' | 'Lead')[] = (agent.completedTrainings || []) as any;
                      return modules
                        .filter(mod => !trainingModuleFilter || mod === trainingModuleFilter)
                        .map(mod => {
                          const activeTMs = trainingMaterials.filter(tm => tm.isActive);
                          const moduleTMs = activeTMs.filter(tm => {
                            const tmMods = Array.isArray(tm.module) ? tm.module : [tm.module].filter(Boolean);
                            const expandedMods = tmMods.flatMap(m => m === 'Both' ? ['Sales', 'Operations'] : [m]);
                            return expandedMods.includes(mod);
                          });
                          
                          const totalPoints = moduleTMs.reduce((sum, tm) => sum + (tm.points || 0), 0);
                          const totalTasks = moduleTMs.length;
                          
                          const pointsMap = agent.trainingPoints || {};
                          const pointsAchieved = moduleTMs.reduce((sum, tm) => sum + (pointsMap[tm.id!] || 0), 0);
                          const tasksDone = moduleTMs.filter(tm => (pointsMap[tm.id!] || 0) > 0).length;
                          const tasksPending = totalTasks - tasksDone;
                          
                          const status = agent.trainingModuleStatus?.[mod] || 'Active';

                          if (trainingStatusFilter !== 'All' && status !== trainingStatusFilter) return null;
                          if (totalTasks === 0) return null;

                          return (
                            <tr key={`${agent.id}-${mod}`} className="hover:bg-zinc-50 transition-colors">
                              <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{agent.name}</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">{agent.joiningDate}</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">{mod}</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">{totalPoints}</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">{totalTasks}</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">{tasksDone}</td>
                              <td className="px-4 py-3 text-sm font-bold text-orange-600 border border-zinc-200">{pointsAchieved}</td>
                              <td className="px-4 py-3 text-sm border border-zinc-200">{tasksPending}</td>
                              <td className="px-4 py-3 border border-zinc-200">
                                <Badge className={status === 'Complete' ? 'bg-green-100 text-green-700' : 'bg-blue-100 text-blue-700'}>
                                  {status}
                                </Badge>
                              </td>
                              <td className="px-4 py-3 border border-zinc-200">
                                <Button 
                                  variant="outline" 
                                  size="sm"
                                  onClick={() => {
                                    setSelectedAgentForTraining(agent);
                                    setSelectedModuleForTraining(mod);
                                    setIsTrainingUpdateModalOpen(true);
                                  }}
                                >
                                  Update
                                </Button>
                              </td>
                            </tr>
                          );
                        });
                    })
                    .filter(Boolean)}
                </tbody>
              </table>
            </Card>
          </div>
        )}

        {activeTab === 'training' && hasPermission('training') && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <h2 className="text-xl font-bold">Training Master</h2>
              {getPermissionLevel('training') === 'Complete' && (
                <Button onClick={() => { setTrainingForm({ module: ['Sales'], points: 1, day: 1, time: 15, isActive: true, link: '', priority: 0 }); setIsTrainingModalOpen(true); }}>Add Training Material</Button>
              )}
            </div>
            <Card className="overflow-x-auto">
              <table className="w-full text-left border-collapse border border-zinc-200">
                <thead>
                  <tr className="bg-zinc-50">
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Priority</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Remarks</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Module</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Points</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Day</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Time (Min)</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Link</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Status</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {[...trainingMaterials].sort((a, b) => (a.priority || 0) - (b.priority || 0)).map(tm => (
                    <tr key={tm.id} className="hover:bg-zinc-50 transition-colors">
                      <td className="px-4 py-3 text-sm border border-zinc-200">{tm.priority || 0}</td>
                      <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{tm.remarks}</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200">{Array.isArray(tm.module) ? tm.module.join(', ') : tm.module}</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200">{tm.points}</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200">{tm.day}</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200">{tm.time}</td>
                      <td className="px-4 py-3 text-sm border border-zinc-200">
                        {tm.link && (
                          <a href={tm.link} target="_blank" rel="noopener noreferrer" className="text-orange-600 hover:underline">Link</a>
                        )}
                      </td>
                      <td className="px-4 py-3 border border-zinc-200">
                        <Badge className={tm.isActive ? 'bg-green-100 text-green-700' : 'bg-red-100 text-red-700'}>
                          {tm.isActive ? 'Active' : 'Deactive'}
                        </Badge>
                      </td>
                      <td className="px-4 py-3 border border-zinc-200">
                        <div className="flex items-center gap-2">
                          {getPermissionLevel('training') === 'Complete' && (
                            <>
                              <button onClick={() => { 
                                setTrainingForm({
                                  ...tm,
                                  module: Array.isArray(tm.module) ? tm.module : [tm.module].filter(Boolean) as any,
                                  points: Number(tm.points) || 1,
                                  day: Number(tm.day) || 1,
                                  time: Number(tm.time) || 15,
                                  isActive: tm.isActive !== undefined ? tm.isActive : true,
                                  priority: Number(tm.priority) || 0
                                }); 
                                setIsTrainingModalOpen(true); 
                              }} className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors">
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => tm.id && setDeleteConfirmation({ id: tm.id, type: 'trainingMaterials', name: tm.remarks })}
                                className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </>
                          )}
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </Card>
          </div>
        )}

      {/* Training Update Modal */}
      <Modal
        isOpen={isTrainingUpdateModalOpen}
        onClose={() => setIsTrainingUpdateModalOpen(false)}
        title={`Update Training: ${selectedAgentForTraining?.name} - ${selectedModuleForTraining}`}
      >
        <div className="space-y-6">
          <div className="flex items-center justify-between p-4 bg-zinc-900 rounded-xl text-white">
            <div>
              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Module Progress</p>
              <p className="text-2xl font-bold">
                {(() => {
                  const activeTMs = trainingMaterials.filter(tm => tm.isActive);
                  const moduleTMs = activeTMs.filter(tm => {
                    const tmMods = Array.isArray(tm.module) ? tm.module : [tm.module].filter(Boolean);
                    const expandedMods = tmMods.flatMap(m => m === 'Both' ? ['Sales', 'Operations'] : [m]);
                    return expandedMods.includes(selectedModuleForTraining as any);
                  });
                  const totalPoints = moduleTMs.reduce((sum, tm) => sum + (tm.points || 0), 0);
                  const pointsMap = selectedAgentForTraining?.trainingPoints || {};
                  const pointsAchieved = moduleTMs.reduce((sum, tm) => sum + (pointsMap[tm.id!] || 0), 0);
                  return `${pointsAchieved} / ${totalPoints}`;
                })()}
              </p>
            </div>
            <div className="text-center">
              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-1">Module Status</p>
              <select 
                className="px-3 py-1 bg-zinc-800 border border-zinc-700 rounded-lg text-sm outline-none focus:ring-2 focus:ring-orange-500 text-white"
                value={selectedAgentForTraining?.trainingModuleStatus?.[selectedModuleForTraining as string] || 'Active'}
                onChange={async (e) => {
                  if (!selectedAgentForTraining?.id || !selectedModuleForTraining) return;
                  const val = e.target.value as 'Active' | 'Complete';
                  const newStatus = { ...(selectedAgentForTraining.trainingModuleStatus || {}), [selectedModuleForTraining]: val };
                  setSelectedAgentForTraining({ ...selectedAgentForTraining, trainingModuleStatus: newStatus });
                  try {
                    await updateDoc(doc(db, 'employees', selectedAgentForTraining.id), { trainingModuleStatus: newStatus });
                  } catch (error) {
                    handleFirestoreError(error, OperationType.UPDATE, 'employees');
                  }
                }}
              >
                <option value="Active">Active</option>
                <option value="Complete">Complete</option>
              </select>
            </div>
            <div className="text-center">
              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider mb-1">Joining Date</p>
              <input 
                type="date"
                className="px-3 py-1 bg-zinc-800 border border-zinc-700 rounded-lg text-sm outline-none focus:ring-2 focus:ring-orange-500 text-white"
                value={selectedAgentForTraining?.joiningDate || ''}
                onChange={async (e) => {
                  if (!selectedAgentForTraining?.id) return;
                  const val = e.target.value;
                  setSelectedAgentForTraining({ ...selectedAgentForTraining, joiningDate: val });
                  try {
                    await updateDoc(doc(db, 'employees', selectedAgentForTraining.id), { joiningDate: val });
                  } catch (error) {
                    handleFirestoreError(error, OperationType.UPDATE, 'employees');
                  }
                }}
              />
            </div>
            <div className="text-right">
              <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Tasks Done</p>
              <p className="text-2xl font-bold text-orange-400">
                {(() => {
                  const activeTMs = trainingMaterials.filter(tm => tm.isActive);
                  const moduleTMs = activeTMs.filter(tm => {
                    const tmMods = Array.isArray(tm.module) ? tm.module : [tm.module].filter(Boolean);
                    const expandedMods = tmMods.flatMap(m => m === 'Both' ? ['Sales', 'Operations'] : [m]);
                    return expandedMods.includes(selectedModuleForTraining as any);
                  });
                  const pointsMap = selectedAgentForTraining?.trainingPoints || {};
                  const tasksDone = moduleTMs.filter(tm => (pointsMap[tm.id!] || 0) > 0).length;
                  return `${tasksDone} / ${moduleTMs.length}`;
                })()}
              </p>
            </div>
          </div>

          <div className="max-h-[50vh] overflow-y-auto custom-scrollbar border border-zinc-200 rounded-xl overflow-hidden">
            <table className="w-full text-left border-collapse">
              <thead>
                <tr className="bg-zinc-100 border-b border-zinc-200">
                  <th className="px-3 py-2 text-[10px] font-bold text-zinc-500 uppercase">Pri</th>
                  <th className="px-3 py-2 text-[10px] font-bold text-zinc-500 uppercase">Task</th>
                  <th className="px-3 py-2 text-[10px] font-bold text-zinc-500 uppercase text-center">Max</th>
                  <th className="px-3 py-2 text-[10px] font-bold text-zinc-500 uppercase w-24">Achieved</th>
                  <th className="px-3 py-2 text-[10px] font-bold text-zinc-500 uppercase text-center">Action</th>
                </tr>
              </thead>
              <tbody>
                {trainingMaterials
                  .filter(tm => tm.isActive)
                  .filter(tm => {
                    const tmMods = Array.isArray(tm.module) ? tm.module : [tm.module].filter(Boolean);
                    const expandedMods = tmMods.flatMap(m => m === 'Both' ? ['Sales', 'Operations'] : [m]);
                    return expandedMods.includes(selectedModuleForTraining as any);
                  })
                  .sort((a, b) => (a.priority || 0) - (b.priority || 0))
                  .map(tm => {
                    const currentPoints = selectedAgentForTraining?.trainingPoints?.[tm.id!] || 0;
                    const isCompleted = currentPoints >= tm.points;

                    return (
                      <tr key={tm.id} className={cn(
                        "border-b border-zinc-100 transition-colors",
                        isCompleted ? "bg-green-50/50" : "hover:bg-zinc-50"
                      )}>
                        <td className="px-3 py-2 text-xs font-bold text-zinc-400">{tm.priority || 0}</td>
                        <td className="px-3 py-2">
                          <p className="text-sm font-medium text-zinc-900">{tm.remarks}</p>
                        </td>
                        <td className="px-3 py-2 text-center text-xs font-bold text-zinc-500">{tm.points}</td>
                        <td className="px-3 py-2">
                          <select 
                            className="w-full px-2 py-1 bg-white border border-zinc-200 rounded text-xs outline-none focus:ring-1 focus:ring-orange-500"
                            value={currentPoints}
                            onChange={(e) => {
                              if (!selectedAgentForTraining?.id) return;
                              const val = Number(e.target.value);
                              const newPoints = { ...(selectedAgentForTraining.trainingPoints || {}), [tm.id!]: val };
                              setSelectedAgentForTraining({ ...selectedAgentForTraining, trainingPoints: newPoints });
                            }}
                          >
                            {Array.from({ length: tm.points + 1 }, (_, i) => (
                              <option key={i} value={i}>{i}</option>
                            ))}
                          </select>
                        </td>
                        <td className="px-3 py-2 text-center">
                          <Button 
                            size="sm"
                            variant="ghost"
                            className="h-7 px-2 text-[10px] font-bold uppercase hover:bg-zinc-900 hover:text-white transition-all"
                            onClick={async () => {
                              if (!selectedAgentForTraining?.id) return;
                              try {
                                await updateDoc(doc(db, 'employees', selectedAgentForTraining.id), { 
                                  trainingPoints: selectedAgentForTraining.trainingPoints 
                                });
                              } catch (error) {
                                handleFirestoreError(error, OperationType.UPDATE, 'employees');
                              }
                            }}
                          >
                            Submit
                          </Button>
                        </td>
                      </tr>
                    );
                  })}
              </tbody>
            </table>
          </div>
          <div className="flex justify-end pt-4 border-t border-zinc-100">
            <Button onClick={() => setIsTrainingUpdateModalOpen(false)} className="bg-orange-600 w-full md:w-auto">
              Close
            </Button>
          </div>
        </div>
      </Modal>

      {/* Training Material Modal */}
      <Modal
        isOpen={isTrainingModalOpen}
        onClose={() => setIsTrainingModalOpen(false)}
        title={trainingForm.id ? 'Edit Training Material' : 'Add Training Material'}
      >
        <form onSubmit={async (e) => {
          e.preventDefault();
          try {
            console.log("Submitting training form:", trainingForm);
            if (trainingForm.id) {
              const { id, ...updateData } = trainingForm;
              await updateDoc(doc(db, 'trainingMaterials', id), {
                ...updateData,
                updatedAt: serverTimestamp()
              });
            } else {
              await addDoc(collection(db, 'trainingMaterials'), {
                ...trainingForm,
                createdAt: serverTimestamp()
              });
            }
            setIsTrainingModalOpen(false);
          } catch (err) {
            console.error(err);
            handleFirestoreError(err, OperationType.WRITE, 'trainingMaterials');
          }
        }} className="space-y-4">
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Remarks</label>
            <input type="text" required className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg" value={trainingForm.remarks || ''} onChange={e => setTrainingForm({...trainingForm, remarks: e.target.value})} />
          </div>
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Link</label>
            <input type="url" className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg" value={trainingForm.link || ''} onChange={e => setTrainingForm({...trainingForm, link: e.target.value})} />
          </div>
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Module</label>
            <div className="flex gap-4">
              {['Sales', 'Operations', 'Lead'].map(mod => (
                <label key={mod} className="flex items-center gap-2">
                  <input 
                    type="checkbox" 
                    checked={Array.isArray(trainingForm.module) ? trainingForm.module.includes(mod as any) : trainingForm.module === mod}
                    onChange={(e) => {
                      const currentModules = Array.isArray(trainingForm.module) ? trainingForm.module : [trainingForm.module].filter(Boolean) as any[];
                      if (e.target.checked) {
                        setTrainingForm({...trainingForm, module: [...currentModules, mod] as any});
                      } else {
                        setTrainingForm({...trainingForm, module: currentModules.filter(m => m !== mod) as any});
                      }
                    }}
                  />
                  <span className="text-sm">{mod}</span>
                </label>
              ))}
            </div>
          </div>
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Priority</label>
            <input type="number" className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg" value={trainingForm.priority || 0} onChange={e => setTrainingForm({...trainingForm, priority: Number(e.target.value)})} />
          </div>
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Points</label>
            <select className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg" value={trainingForm.points} onChange={e => setTrainingForm({...trainingForm, points: Number(e.target.value)})}>
              {[...Array(10)].map((_, i) => <option key={i+1} value={i+1}>{i+1}</option>)}
            </select>
          </div>
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Day</label>
            <select className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg" value={trainingForm.day} onChange={e => setTrainingForm({...trainingForm, day: Number(e.target.value)})}>
              {[...Array(15)].map((_, i) => <option key={i+1} value={i+1}>{i+1}</option>)}
            </select>
          </div>
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Time (Min)</label>
            <select className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg" value={trainingForm.time} onChange={e => setTrainingForm({...trainingForm, time: Number(e.target.value)})}>
              {[15, 30, 45, 60, 90, 120].map(t => <option key={t} value={t}>{t}</option>)}
            </select>
          </div>
          <div className="flex items-center gap-2">
            <input type="checkbox" checked={trainingForm.isActive} onChange={e => setTrainingForm({...trainingForm, isActive: e.target.checked})} />
            <label className="text-sm">Active</label>
          </div>
          <Button type="submit" className="w-full bg-orange-600">Save</Button>
        </form>
      </Modal>

      {/* Delete Confirmation Modal */}
      <Modal
        isOpen={!!deleteConfirmation}
        onClose={() => setDeleteConfirmation(null)}
        title="Confirm Deletion"
      >
        <div className="space-y-6">
          <div className="flex items-center gap-4 p-4 bg-red-50 border border-red-100 rounded-2xl">
            <div className="w-12 h-12 bg-red-100 rounded-full flex items-center justify-center text-red-600 shrink-0">
              <AlertCircle className="w-6 h-6" />
            </div>
            <div>
              <h4 className="font-bold text-red-900">Are you absolutely sure?</h4>
              <p className="text-sm text-red-700">
                You are about to delete the {deleteConfirmation?.type} <span className="font-bold">"{deleteConfirmation?.name}"</span>. This action cannot be undone.
              </p>
            </div>
          </div>

          {!(pagePasswords['globalDelete']?.isOpen ?? false) && (
            <div className="space-y-2">
              <label className="text-xs font-bold text-zinc-500 uppercase">Enter Delete Password to Confirm</label>
              <input
                type="password"
                className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-red-500 outline-none"
                placeholder="Password"
                value={deleteConfirmation?.passwordInput || ''}
                onChange={(e) => setDeleteConfirmation(prev => prev ? { ...prev, passwordInput: e.target.value } : null)}
                onKeyDown={(e) => e.key === 'Enter' && confirmDelete()}
                autoFocus
              />
              {deleteConfirmation?.error && <p className="text-xs text-red-600 font-medium">{deleteConfirmation.error}</p>}
            </div>
          )}

          <div className="flex gap-3 pt-4">
            <Button 
              onClick={confirmDelete}
              variant="danger" 
              className="flex-1"
            >
              Yes, Delete Permanently
            </Button>
            <Button 
              variant="ghost" 
              onClick={() => setDeleteConfirmation(null)}
              className="flex-1"
            >
              Cancel
            </Button>
          </div>
        </div>
      </Modal>
      </main>

      {/* Footer */}
      <footer className="max-w-7xl mx-auto px-4 py-12 border-t border-zinc-200 text-center">
        <p className="text-sm text-zinc-400">© 2026 Nitsa Bravo Points Dashboard. Built with ❤️ for the team.</p>
      </footer>
        {/* Password Prompt Modal */}
      <Modal
        isOpen={isPasswordPromptOpen}
        onClose={() => {
          setIsPasswordPromptOpen(false);
          setPendingTab(null);
          setPasswordInput('');
          setPasswordError(null);
        }}
        title={`Enter Password for ${allPermissions.find(p => p.id === pendingTab)?.label || 'Page'}`}
      >
        <div className="space-y-4">
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Password</label>
            <input
              type="password"
              className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg"
              value={passwordInput}
              onChange={(e) => setPasswordInput(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && verifyPagePassword()}
              autoFocus
            />
            {passwordError && <p className="text-xs text-red-600 font-medium">{passwordError}</p>}
          </div>
          <Button onClick={verifyPagePassword} className="w-full bg-orange-600">
            Unlock Page
          </Button>
        </div>
      </Modal>

      {/* Master Password Prompt Modal */}
      <Modal
        isOpen={isMasterPasswordPromptOpen}
        onClose={() => {
          setIsMasterPasswordPromptOpen(false);
          setMasterPasswordInput('');
          setPasswordError(null);
        }}
        title="Enter Master Password"
      >
        <div className="space-y-4">
          <p className="text-sm text-zinc-600">Please enter the master password to access the Password Manager.</p>
          <div className="space-y-2">
            <label className="text-xs font-bold text-zinc-500 uppercase">Master Password</label>
            <input
              type="password"
              className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg"
              value={masterPasswordInput}
              onChange={(e) => setMasterPasswordInput(e.target.value)}
              onKeyDown={(e) => e.key === 'Enter' && verifyMasterPassword()}
              autoFocus
            />
            {passwordError && <p className="text-xs text-red-600 font-medium">{passwordError}</p>}
          </div>
          <Button onClick={verifyMasterPassword} className="w-full bg-orange-600">
            Access Manager
          </Button>
        </div>
      </Modal>

      {/* Agent Ledger Pop-up Modal */}
      <Modal
        isOpen={agentLedgerModal.isOpen}
        onClose={() => setAgentLedgerModal({ isOpen: false, employeeId: null })}
        title={`Incentive Ledger - ${employees.find(e => e.id === agentLedgerModal.employeeId)?.name || 'Employee'}`}
        maxWidth="max-w-6xl"
      >
        <div className="space-y-6 max-h-[80vh] overflow-y-auto pr-2">
          {agentLedgerModal.employeeId && (
            <>
              {/* Summary Table */}
              <Card className="overflow-hidden border-zinc-200 shadow-sm">
                <div className="bg-zinc-50 px-4 py-2 border-b border-zinc-200 flex justify-between items-center">
                  <h3 className="text-[10px] font-bold text-zinc-500 uppercase tracking-wider">Employee Summary (Incentive + Quaterly) (Approx)</h3>
                  <div className="flex items-center gap-4 text-[10px] font-bold uppercase tracking-wider">
                    <div className="flex items-center gap-1.5">
                      <div className="w-3 h-3 bg-white border border-zinc-300 rounded-sm"></div>
                      <span className="text-zinc-500">White - Due</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <div className="w-3 h-3 bg-purple-100 border border-purple-200 rounded-sm"></div>
                      <span className="text-purple-600">Purple - Eligible to be paid</span>
                    </div>
                    <div className="flex items-center gap-1.5">
                      <div className="w-3 h-3 bg-green-100 border border-green-200 rounded-sm"></div>
                      <span className="text-green-600">Green - Paid</span>
                    </div>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-white text-[10px] font-bold text-zinc-400 uppercase">
                        <th className="px-4 py-2 border-r">Month</th>
                        <th className="px-4 py-2 border-r">Daily</th>
                        <th className="px-4 py-2 border-r">Weekly</th>
                        <th className="px-4 py-2 border-r">Monthly</th>
                        <th className="px-4 py-2 border-r">Quarterly</th>
                        <th className="px-4 py-2 border-r">Annually</th>
                        <th className="px-4 py-2 border-r">Total Approx Due</th>
                        <th className="px-4 py-2 border-r">Total Approx Paid</th>
                        <th className="px-4 py-2 bg-orange-50 text-orange-600">Approx Balance</th>
                      </tr>
                    </thead>
                    <tbody>
                      {modalMonthlyIncentiveData.map((row: any) => (
                        <tr key={row.month} className="text-sm font-medium border-b border-zinc-100 hover:bg-zinc-50 transition-colors text-zinc-600">
                          <td className="px-4 py-2 border-r font-bold text-zinc-900">{row.month}</td>
                          <td className={cn("px-4 py-2 border-r", row.typeStatuses.Daily.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Daily.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Daily.toLocaleString()}</td>
                          <td className={cn("px-4 py-2 border-r", row.typeStatuses.Weekly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Weekly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Weekly.toLocaleString()}</td>
                          <td className={cn("px-4 py-2 border-r", row.typeStatuses.Monthly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Monthly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Monthly.toLocaleString()}</td>
                          <td className={cn("px-4 py-2 border-r", row.typeStatuses.Quarterly.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Quarterly.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Quarterly.toLocaleString()}</td>
                          <td className={cn("px-4 py-2 border-r", row.typeStatuses.Annually.isPaid ? "bg-green-100 text-green-700" : row.typeStatuses.Annually.isEligible ? "bg-purple-100 text-purple-700" : "")}>₹{row.Annually.toLocaleString()}</td>
                          <td className="px-4 py-2 border-r text-red-600 font-bold">₹{row.totalDue.toLocaleString()}</td>
                          <td className="px-4 py-2 border-r text-green-600 font-bold">₹{row.totalPaid.toLocaleString()}</td>
                          <td className="px-4 py-2 bg-orange-50 text-orange-600 font-bold">₹{(row.totalDue - row.totalPaid).toLocaleString()}</td>
                        </tr>
                      ))}
                      <tr className="text-sm font-black text-zinc-900 bg-zinc-50">
                        <td className="px-4 py-3 border-r uppercase">Grand Total</td>
                        <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === agentLedgerModal.employeeId && i.type === 'Daily').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                        <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === agentLedgerModal.employeeId && i.type === 'Weekly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                        <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === agentLedgerModal.employeeId && i.type === 'Monthly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                        <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === agentLedgerModal.employeeId && i.type === 'Quarterly').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                        <td className="px-4 py-3 border-r">₹{incentives.filter(i => i.employeeId === agentLedgerModal.employeeId && i.type === 'Annually').reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                        <td className="px-4 py-3 border-r text-red-600 font-bold">₹{incentives.filter(i => i.employeeId === agentLedgerModal.employeeId).reduce((sum, i) => sum + i.amount, 0).toLocaleString()}</td>
                        <td className="px-4 py-3 border-r text-green-600 font-bold">₹{incentivePayments.filter(p => p.employeeId === agentLedgerModal.employeeId).reduce((sum, p) => sum + p.amount, 0).toLocaleString()}</td>
                        <td className="px-4 py-3 bg-orange-100 text-orange-700 text-lg font-bold">
                          ₹{(
                            incentives.filter(i => i.employeeId === agentLedgerModal.employeeId).reduce((sum, i) => sum + i.amount, 0) - 
                            incentivePayments.filter(p => p.employeeId === agentLedgerModal.employeeId).reduce((sum, p) => sum + p.amount, 0)
                          ).toLocaleString()}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </Card>

              {/* Ledger Table */}
              <Card className="p-6">
                <div className="flex items-center gap-4 mb-4 text-[10px] font-bold uppercase tracking-wider">
                  <div className="flex items-center gap-1.5">
                    <div className="w-3 h-3 rounded border border-zinc-200 bg-white"></div>
                    <span className="text-zinc-500">Due</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <div className="w-3 h-3 rounded bg-purple-100 border border-purple-200"></div>
                    <span className="text-purple-600">Eligible to be paid</span>
                  </div>
                  <div className="flex items-center gap-1.5">
                    <div className="w-3 h-3 rounded bg-green-100 border border-green-200"></div>
                    <span className="text-green-600">Paid</span>
                  </div>
                </div>
                <div className="overflow-x-auto">
                  <table className="w-full text-left border-collapse">
                    <thead>
                      <tr className="bg-zinc-50 text-xs font-bold text-zinc-500 uppercase">
                        <th className="px-4 py-3 border-b">Date</th>
                        <th className="px-4 py-3 border-b">Description</th>
                        <th className="px-4 py-3 border-b text-right">Daily</th>
                        <th className="px-4 py-3 border-b text-right">Weekly</th>
                        <th className="px-4 py-3 border-b text-right">Monthly</th>
                        <th className="px-4 py-3 border-b text-right">Quaterly</th>
                        <th className="px-4 py-3 border-b text-right">Annually</th>
                        <th className="px-4 py-3 border-b text-right">Approx Due (₹)</th>
                        <th className="px-4 py-3 border-b text-right">Approx Paid (₹)</th>
                        <th className="px-4 py-3 border-b text-right">Approx Balance (₹)</th>
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const employee = employees.find(e => e.id === agentLedgerModal.employeeId);
                        const empIncentives = incentives
                          .filter(i => i.employeeId === agentLedgerModal.employeeId)
                          .map(i => ({ ...i, entryType: 'due' as const }));
                        const empPayments = incentivePayments
                          .filter(p => p.employeeId === agentLedgerModal.employeeId)
                          .map(p => ({ ...p, entryType: 'paid' as const }));
                        
                        // Find sales for this employee
                        const empSales = !employee ? [] : sales.filter(s => {
                          const isAgent = (s.agentEmail?.toLowerCase() === employee.email?.toLowerCase()) || 
                                          (s.agent?.toLowerCase().trim() === employee.name.toLowerCase().trim());
                          const isAssocBDE = (s.associateBdeEmail?.toLowerCase() === employee.email?.toLowerCase()) || 
                                             (s.associateBde?.toLowerCase().trim() === employee.name.toLowerCase().trim());
                          return isAgent || isAssocBDE;
                        }).map(s => ({
                          id: s.id,
                          date: s.date,
                          type: `Sale: ${s.guestName} (${s.tripId})`,
                          amount: getAgentCreditForSale(s, employee, 'packageValue'),
                          remarks: `Destination: ${s.destination}, Status: ${s.advanceCN}`,
                          recordedBy: 'System',
                          recordedAt: s.createdAt ? (s.createdAt as any).toDate?.().toLocaleString() : '',
                          entryType: 'due' as const,
                          isPaid: false,
                          isEligible: true // Sales are generally eligible for incentive tracking
                        }));

                        const allEntries = [...empIncentives, ...empPayments, ...empSales].sort((a, b) => 
                          new Date(a.date).getTime() - new Date(b.date).getTime()
                        );

                        let runningBalance = 0;
                        return allEntries.map((entry, idx) => {
                          const due = entry.entryType === 'due' ? (entry as Incentive).amount : 0;
                          const paid = entry.entryType === 'paid' ? (entry as IncentivePayment).amount : 0;
                          runningBalance += due - paid;

                          const incentive = entry.entryType === 'due' ? (entry as Incentive) : null;

                          return (
                            <tr key={entry.id || idx} className="hover:bg-zinc-50 transition-colors border-b last:border-0">
                              <td className="px-4 py-3 text-sm text-zinc-600">
                                {entry.date}
                                {entry.recordedAt && (
                                  <p className="text-[10px] text-zinc-400 mt-0.5">
                                    {entry.recordedAt}
                                  </p>
                                )}
                              </td>
                              <td className="px-4 py-3 text-sm">
                                <div className="font-medium text-zinc-900">
                                  {entry.entryType === 'due' ? (entry as Incentive).type : 'Payment'}
                                </div>
                                <div className="flex flex-col gap-0.5">
                                  {entry.remarks && <p className="text-xs text-zinc-400 italic">{entry.remarks}</p>}
                                  {entry.recordedBy && (
                                    <p className="text-[10px] text-zinc-500 font-medium">
                                      By: {entry.recordedBy}
                                    </p>
                                  )}
                                </div>
                              </td>
                              <td className="px-4 py-3 text-sm text-right text-zinc-600">
                                {incentive?.type === 'Daily' ? `₹${incentive.amount.toLocaleString()}` : '-'}
                              </td>
                              <td className="px-4 py-3 text-sm text-right text-zinc-600">
                                {incentive?.type === 'Weekly' ? `₹${incentive.amount.toLocaleString()}` : '-'}
                              </td>
                              <td className="px-4 py-3 text-sm text-right text-zinc-600">
                                {incentive?.type === 'Monthly' ? `₹${incentive.amount.toLocaleString()}` : '-'}
                              </td>
                              <td className="px-4 py-3 text-sm text-right text-zinc-600">
                                {incentive?.type === 'Quarterly' ? `₹${incentive.amount.toLocaleString()}` : '-'}
                              </td>
                              <td className="px-4 py-3 text-sm text-right text-zinc-600">
                                {incentive?.type === 'Annually' ? `₹${incentive.amount.toLocaleString()}` : '-'}
                              </td>
                              <td className={cn(
                                "px-4 py-3 text-sm text-right font-medium transition-colors",
                                due > 0 && incentive?.isPaid ? "bg-green-100 text-green-700" : 
                                due > 0 && incentive?.isEligible ? "bg-purple-100 text-purple-700" : 
                                due > 0 ? "text-red-600" : "text-zinc-400"
                              )}>
                                {due > 0 ? (
                                  <div className="flex flex-col items-end gap-1">
                                    <span>₹{due.toLocaleString()}</span>
                                    <div className="flex gap-1">
                                      {incentive?.isEligible && (
                                        <span className="px-1.5 py-0.5 bg-purple-50 text-purple-600 text-[9px] font-bold rounded border border-purple-100 uppercase tracking-tighter">
                                          Eligible
                                        </span>
                                      )}
                                      {incentive?.isPaid && (
                                        <span className="px-1.5 py-0.5 bg-green-50 text-green-600 text-[9px] font-bold rounded border border-green-100 uppercase tracking-tighter">
                                          Paid
                                        </span>
                                      )}
                                    </div>
                                  </div>
                                ) : '-'}
                              </td>
                              <td className="px-4 py-3 text-sm text-right font-medium text-green-600">
                                {paid > 0 ? `₹${paid.toLocaleString()}` : '-'}
                              </td>
                              <td className="px-4 py-3 text-sm text-right font-bold text-zinc-900">
                                ₹{runningBalance.toLocaleString()}
                              </td>
                            </tr>
                          );
                        });
                      })()}
                      {incentives.filter(i => i.employeeId === agentLedgerModal.employeeId).length === 0 && 
                       incentivePayments.filter(p => p.employeeId === agentLedgerModal.employeeId).length === 0 && (
                        <tr>
                          <td colSpan={10} className="px-4 py-12 text-center text-zinc-400 italic">
                            No transactions found for this employee.
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                </div>
              </Card>
            </>
          )}
          <div className="flex justify-end">
            <Button variant="outline" onClick={() => setAgentLedgerModal({ isOpen: false, employeeId: null })}>
              Close
            </Button>
          </div>
        </div>
      </Modal>

      {remarksModalSale && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-md flex items-center justify-center z-[100] p-4">
          <motion.div 
            initial={{ opacity: 0, y: 20, scale: 0.95 }}
            animate={{ opacity: 1, y: 0, scale: 1 }}
            className="bg-white rounded-2xl shadow-2xl w-full max-w-4xl flex flex-col max-h-[85vh] overflow-hidden border border-zinc-200"
          >
            {/* Header */}
            <div className="px-6 py-4 border-b border-zinc-100 flex items-center justify-between bg-zinc-50/50">
              <div className="flex items-center gap-4">
                <div className="p-2 bg-orange-100 rounded-lg">
                  <MessageSquare className="w-5 h-5 text-orange-600" />
                </div>
                <div>
                  <h3 className="text-lg font-black text-zinc-900 tracking-tight">Remarks Log</h3>
                  <div className="flex items-center gap-2 text-xs font-bold text-zinc-500">
                    <span className="bg-zinc-200/50 px-1.5 py-0.5 rounded text-zinc-700">#{remarksModalSale.tripId}</span>
                    <span className="text-zinc-400">•</span>
                    <span className="text-zinc-600">{remarksModalSale.guestName}</span>
                  </div>
                </div>
              </div>
              <button 
                onClick={() => setRemarksModalSale(null)}
                className="p-2 hover:bg-zinc-200/50 rounded-lg transition-all group"
              >
                <X className="w-5 h-5 text-zinc-400 group-hover:text-zinc-900" />
              </button>
            </div>

            {/* Content - Excel Style Table */}
            <div className="flex-1 overflow-y-auto custom-scrollbar bg-white">
              <table className="w-full border-collapse text-left">
                <thead className="sticky top-0 bg-zinc-100 z-10">
                  <tr className="border-b border-zinc-200">
                    <th className="px-4 py-2 text-[10px] font-black text-zinc-500 uppercase tracking-wider border-r border-zinc-200 w-28">Date</th>
                    <th className="px-4 py-2 text-[10px] font-black text-zinc-500 uppercase tracking-wider border-r border-zinc-200 w-24">Time</th>
                    <th className="px-4 py-2 text-[10px] font-black text-zinc-500 uppercase tracking-wider border-r border-zinc-200 w-40">User</th>
                    <th className="px-4 py-2 text-[10px] font-black text-zinc-500 uppercase tracking-wider">Remark</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-zinc-100">
                  {remarksLog.map((remark, idx) => (
                    <tr key={remark.id} className="hover:bg-zinc-50 transition-colors group">
                      <td className="px-4 py-2 text-xs font-bold text-zinc-500 border-r border-zinc-100 whitespace-nowrap">
                        <div className="flex items-center gap-2">
                          <CalendarIcon className="w-3 h-3 text-zinc-300" />
                          {remark.date}
                        </div>
                      </td>
                      <td className="px-4 py-2 text-xs font-bold text-zinc-500 border-r border-zinc-100 whitespace-nowrap">
                        <div className="flex items-center gap-2">
                          <Clock className="w-3 h-3 text-zinc-300" />
                          {remark.time}
                        </div>
                      </td>
                      <td className="px-4 py-2 border-r border-zinc-100">
                        <div className="flex items-center gap-2">
                          <div className="w-5 h-5 rounded bg-orange-50 flex items-center justify-center text-[10px] font-black text-orange-600 border border-orange-100">
                            {remark.userName.charAt(0).toUpperCase()}
                          </div>
                          <span className="text-xs font-bold text-zinc-700 truncate max-w-[120px]">{remark.userName}</span>
                        </div>
                      </td>
                      <td className="px-4 py-2 text-xs font-medium text-zinc-600 leading-relaxed">
                        {remark.text}
                      </td>
                    </tr>
                  ))}
                  {remarksLog.length === 0 && (
                    <tr>
                      <td colSpan={4} className="px-6 py-20 text-center">
                        <div className="flex flex-col items-center gap-2">
                          <MessageSquare className="w-8 h-8 text-zinc-200" />
                          <p className="text-zinc-400 text-sm font-bold">No remarks yet</p>
                        </div>
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            {/* Footer / Input */}
            {(getPermissionLevel('sales') === 'Complete' || 
              getPermissionLevel('issueOverview') === 'Complete' || 
              activeTab === 'sales' || 
              activeTab === 'agentOverview' || 
              activeTab === 'bdeOverview') && (
              <div className="p-4 border-t border-zinc-100 bg-zinc-50/50">
                <div className="flex gap-3">
                  <div className="flex-1 relative">
                    <textarea
                      value={newRemarkText}
                      onChange={(e) => setNewRemarkText(e.target.value)}
                      placeholder="Type your remark here..."
                      className="w-full bg-white border border-zinc-200 rounded-xl px-4 py-2.5 text-sm focus:ring-2 focus:ring-orange-500/20 focus:border-orange-500 outline-none transition-all resize-none h-[42px] custom-scrollbar font-medium"
                    />
                  </div>
                  <button
                    onClick={handleAddRemark}
                    disabled={isAddingRemark || !newRemarkText.trim()}
                    className="bg-orange-600 text-white px-5 py-2.5 rounded-xl font-black text-xs hover:bg-orange-700 transition-all disabled:opacity-50 disabled:cursor-not-allowed shadow-lg shadow-orange-600/10 flex items-center gap-2 active:scale-95 whitespace-nowrap"
                  >
                    {isAddingRemark ? (
                      <Loader2 className="w-4 h-4 animate-spin" />
                    ) : (
                      <>
                        <Plus className="w-4 h-4" />
                        <span>ADD LOG</span>
                      </>
                    )}
                  </button>
                </div>
              </div>
            )}
          </motion.div>
        </div>
      )}
      {user && (
        <AIChatBot 
          data={{ employees, sales, incentives, incentivePayments, bdes, weeks }} 
          currentUser={user}
        />
      )}
    </div>
  );
}
