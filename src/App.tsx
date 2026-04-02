/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, Component, useRef } from 'react';
import React from 'react';
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
  handleFirestoreError,
  OperationType,
  or
} from './firebase';
import { 
  Trophy, 
  Send, 
  History, 
  LogOut, 
  Search, 
  Award, 
  Star, 
  TrendingUp,
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
  Calculator,
  Menu,
  Clock,
  ExternalLink
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { formatDistanceToNow } from 'date-fns';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

// --- Utilities ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Types ---
interface Role {
  id: string;
  name: string;
  permissions: string[];
  createdAt: any;
}

interface UserProfile {
  uid: string;
  email: string;
  displayName: string;
  photoURL: string;
  pointsBalance: number;
  role?: string;
  permissions?: string[];
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
  createdAt?: any;
}

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
];

interface Sales {
  id?: string;
  week: string;
  date: string;
  guestName: string;
  agent: string;
  agentEmail?: string;
  agentPercentage?: number;
  bde: string;
  associateBde: string;
  assocBdePercentage?: number;
  destination: string;
  tripId: string;
  tripDate: string;
  packageValue?: number;
  lessThan10Percent20k: string;
  advanceCN?: string;
  salesBy?: 'Inhouse' | 'Branch';
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
  ...props 
}: React.ButtonHTMLAttributes<HTMLButtonElement> & { 
  variant?: 'primary' | 'secondary' | 'outline' | 'ghost' | 'danger';
  isLoading?: boolean;
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

const Card = ({ children, className, id }: { children: React.ReactNode; className?: string; id?: string }) => (
  <div id={id} className={cn('bg-white border border-zinc-200 rounded-2xl shadow-sm overflow-hidden', className)}>
    {children}
  </div>
);

const Badge = ({ children, className }: { children: React.ReactNode; className?: string; key?: React.Key }) => (
  <span className={cn('px-2 py-0.5 rounded-full text-[10px] font-bold uppercase tracking-wider', className)}>
    {children}
  </span>
);

const Modal = ({ isOpen, onClose, title, children }: { isOpen: boolean; onClose: () => void; title: string; children: React.ReactNode }) => (
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
          className="relative w-full max-w-2xl bg-white rounded-3xl shadow-2xl overflow-hidden flex flex-col max-h-[90vh]"
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
  const [user, setUser] = useState<any>(null);
  const [profile, setProfile] = useState<UserProfile | null>(null);
  const [allUsers, setAllUsers] = useState<UserProfile[]>([]);
  const [roles, setRoles] = useState<Role[]>([]);
  const [isRolesModalOpen, setIsRolesModalOpen] = useState(false);
  const [selectedRole, setSelectedRole] = useState<Role | null>(null);
  const [isRoleEditModalOpen, setIsRoleEditModalOpen] = useState(false);
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [weeks, setWeeks] = useState<Week[]>([]);
  const [bdes, setBDES] = useState<BDE[]>([]);
  const [sales, setSales] = useState<Sales[]>([]);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'employees' | 'weeks' | 'bdes' | 'sales' | 'weeklyOverview' | 'agentOverview' | 'userManagement' | 'training' | 'trainingOverview'>('dashboard');
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
  const [isMenuOpen, setIsMenuOpen] = useState(false);
  const menuRef = useRef<HTMLDivElement>(null);
  const lastUserRef = useRef<any>(null);

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
  const [hasSetDefaultWeek, setHasSetDefaultWeek] = useState(false);
  const [hasSetDefaultMonth, setHasSetDefaultMonth] = useState(false);
  const [selectedUser, setSelectedUser] = useState<UserProfile | null>(null);
  const [pointsToAward, setPointsToAward] = useState(10);
  const [reason, setReason] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [fatalError, setFatalError] = useState<any>(null);
  const [incentiveCalcSales, setIncentiveCalcSales] = useState<number>(0);
  const [incentiveCalcMonthlySales, setIncentiveCalcMonthlySales] = useState<number>(0);
  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    id: string;
    type: 'employee' | 'week' | 'bde' | 'sales' | 'user' | 'trainingMaterials';
    name: string;
  } | null>(null);
  const unsubsRef = useRef<(() => void)[]>([]);

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

  const isAdmin = profile?.role === 'Admin' || (profile?.email === 'nsingla09@gmail.com' && auth.currentUser?.emailVerified);
  const isManager = profile?.role === 'Manager';
  const isTrainer = profile?.role?.toLowerCase() === 'trainer';
  const hasTrainingOverviewPermission = Array.isArray(profile?.permissions) && profile.permissions.includes('trainingOverview');
  const isPrivileged = isAdmin || isManager || isTrainer || hasTrainingOverviewPermission;

  const isEmployee = useMemo(() => {
    return employees.some(emp => emp.email && emp.email.toLowerCase() === profile?.email?.toLowerCase());
  }, [employees, profile]);

  const hasPermission = (permission: string) => {
    if (isAdmin) return true;
    if (permission === 'trainingOverview' && (isManager || isTrainer)) return true;
    if (permission === 'agentOverview' || permission === 'sales') {
      if (isEmployee) return true;
    }
    return (Array.isArray(profile?.permissions) && profile.permissions.includes(permission));
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
  }, [isEmployee, employees, profile, selectedAgentOverview, salesAgentFilter]);

  const filteredEmployees = useMemo(() => {
    return employees.filter(emp => {
      const matchesSearch = 
        emp.name.toLowerCase().includes(employeeSearchQuery.toLowerCase()) || 
        emp.employeeCode.toLowerCase().includes(employeeSearchQuery.toLowerCase()) ||
        (emp.email && emp.email.toLowerCase().includes(employeeSearchQuery.toLowerCase()));
      
      const matchesStatus = 
        employeeStatusFilter === 'All' || 
        emp.status === employeeStatusFilter;
      
      const canSeeEmployee = 
        isPrivileged || 
        !emp.email || 
        emp.email === profile?.email;
      
      const matchesAssocBde = 
        employeeAssocBdeFilter === 'All' || 
        (employeeAssocBdeFilter === 'Yes' ? emp.isAssocBDE === true : emp.isAssocBDE !== true);
      
      return matchesSearch && matchesStatus && canSeeEmployee && matchesAssocBde;
    });
  }, [employees, employeeSearchQuery, employeeStatusFilter, employeeAssocBdeFilter, profile, isPrivileged]);

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
            userSnap = await Promise.race([
              getDoc(userRef),
              new Promise((_, reject) => setTimeout(() => reject(new Error('Timeout')), 10000))
            ]);
          } catch (err) {
            console.error("Error fetching user profile:", err);
            throw err;
          }
          
          let currentProfile: UserProfile;
          if (!userSnap.exists()) {
            console.log("User profile not found by UID, checking by email...");
            const q = query(collection(db, 'users'), where('email', '==', currentUser.email), limit(1));
            const querySnap = await getDocs(q);
            
            if (!querySnap.empty) {
              console.log("Found pre-provisioned profile by email, migrating to UID...");
              const preDoc = querySnap.docs[0];
              const preData = preDoc.data() as UserProfile;
              
              currentProfile = {
                ...preData,
                uid: currentUser.uid,
                displayName: currentUser.displayName || preData.displayName,
                photoURL: currentUser.photoURL || preData.photoURL,
                pointsBalance: preData.pointsBalance ?? 100,
                role: preData.role || 'Agent',
                permissions: preData.permissions || ['dashboard', 'sales', 'training'],
              };
              
              await setDoc(userRef, currentProfile);
              if (preDoc.id !== currentUser.uid) {
                await deleteDoc(preDoc.ref);
              }
              setProfile(currentProfile);
            } else {
              console.log("No profile found, creating new profile...");
              const isFirstUser = currentUser.email === 'nsingla09@gmail.com' && currentUser.emailVerified;
              currentProfile = {
                uid: currentUser.uid,
                email: currentUser.email || '',
                displayName: currentUser.displayName || 'Anonymous User',
                photoURL: currentUser.photoURL || '',
                pointsBalance: 100, // Starting points
                role: isFirstUser ? 'Admin' : 'Agent',
                permissions: isFirstUser 
                  ? allPermissions.map(p => p.id)
                  : ['dashboard', 'sales', 'training'],
              };
              await setDoc(userRef, currentProfile);
              setProfile(currentProfile);
            }
            
            if (!(Array.isArray(currentProfile.permissions) && currentProfile.permissions.includes('dashboard')) && (Array.isArray(currentProfile.permissions) && currentProfile.permissions.includes('employees'))) {
              setActiveTab('employees');
            }
          } else {
            console.log("User profile found:", userSnap.data());
            currentProfile = userSnap.data() as UserProfile;
            
            // Ensure role and permissions exist
            if (!currentProfile.role) {
              const role = (currentProfile.email === 'nsingla09@gmail.com' && currentUser.emailVerified) ? 'Admin' : 'Agent';
              const permissions = role === 'Admin' 
                ? allPermissions.map(p => p.id)
                : ['dashboard', 'sales', 'training'];
              currentProfile.role = role;
              currentProfile.permissions = permissions;
              try {
                await updateDoc(userRef, { role, permissions });
              } catch (err) {
                console.warn("Could not update role/permissions in DB, using local defaults", err);
              }
            }
            setProfile(currentProfile);
            if (!(Array.isArray(currentProfile.permissions) && currentProfile.permissions.includes('dashboard')) && (Array.isArray(currentProfile.permissions) && currentProfile.permissions.includes('employees'))) {
              setActiveTab('employees');
            }
          }

          const isUserAdmin = currentProfile.role?.toLowerCase() === 'admin' || (currentProfile.email === 'nsingla09@gmail.com' && currentUser.emailVerified);
          const isUserManager = currentProfile.role?.toLowerCase() === 'manager';
          const isUserTrainer = currentProfile.role?.toLowerCase() === 'trainer';
          const hasTrainingOverviewPermission = Array.isArray(currentProfile.permissions) && currentProfile.permissions.includes('trainingOverview');
          const isUserPrivileged = isUserAdmin || isUserManager || isUserTrainer || hasTrainingOverviewPermission;

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

          // Listen for roles
          const unsubRoles = onSnapshot(collection(db, 'roles'), (snap) => {
            const rolesList = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Role));
            setRoles(rolesList);
            
            // Initialize default roles if none exist
            if (rolesList.length === 0 && isUserAdmin) {
              console.log("Initializing default roles...");
              const defaultRoles = [
                { name: 'Admin', permissions: allPermissions.map(p => p.id) },
                { name: 'Manager', permissions: ['dashboard', 'employees', 'weeks', 'bdes', 'sales', 'weeklyOverview', 'agentOverview', 'training', 'trainingOverview'] },
                { name: 'Trainer', permissions: ['dashboard', 'training', 'trainingOverview'] },
                { name: 'Agent', permissions: ['dashboard', 'sales', 'training', 'trainingOverview'] },
                { name: 'Extra Team', permissions: ['dashboard', 'training', 'trainingOverview'] },
                { name: 'Trainee', permissions: ['dashboard', 'training', 'trainingOverview'] },
              ];
              
              defaultRoles.forEach(async (role) => {
                try {
                  await addDoc(collection(db, 'roles'), {
                    ...role,
                    createdAt: serverTimestamp(),
                  });
                } catch (err) {
                  console.error("Error creating default role:", err);
                }
              });
            }
          }, (err) => {
            console.error("Roles listener error:", err);
            if (auth.currentUser) handleFirestoreError(err, OperationType.GET, 'roles');
          });
          unsubsRef.current.push(unsubRoles);

          // Listen for all users
          const unsubUsers = onSnapshot(collection(db, 'users'), (snap) => {
            const users = snap.docs.map(doc => doc.data() as UserProfile);
            setAllUsers(users);
          }, (err) => {
            console.error("Users listener error:", err);
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'users');
            }
          });
          unsubsRef.current.push(unsubUsers);

          // Listen for transactions
          const transQuery = isUserPrivileged 
            ? query(collection(db, 'transactions'), orderBy('createdAt', 'desc'), limit(20))
            : query(collection(db, 'transactions'), 
                or(
                  where('fromUid', '==', currentUser.uid),
                  where('toUid', '==', currentUser.uid)
                ));

          const unsubTrans = onSnapshot(transQuery, (snap) => {
            const trans = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Transaction));
            // Sort client-side for non-privileged users since we removed orderBy to avoid index requirement
            if (!isUserPrivileged) {
              trans.sort((a, b) => {
                const dateA = a.createdAt?.toMillis ? a.createdAt.toMillis() : 0;
                const dateB = b.createdAt?.toMillis ? b.createdAt.toMillis() : 0;
                return dateB - dateA;
              });
              setTransactions(trans.slice(0, 50));
            } else {
              setTransactions(trans);
            }
          }, (err) => {
            console.error("Transactions listener error:", err);
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'transactions');
            }
          });
          unsubsRef.current.push(unsubTrans);

          // Listen for Training Materials
          const unsubTraining = onSnapshot(collection(db, 'trainingMaterials'), (snap) => {
            const training = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as TrainingMaterial));
            setTrainingMaterials(training);
          }, (err) => {
            console.error("Training materials listener error:", err);
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'trainingMaterials');
            }
          });
          unsubsRef.current.push(unsubTraining);

          // Listen for weeks
          const unsubWeeks = onSnapshot(collection(db, 'weeks'), (snap) => {
            const wks = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Week));
            setWeeks(wks);
          }, (err) => {
            console.error("Weeks listener error:", err);
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'weeks');
            }
          });
          unsubsRef.current.push(unsubWeeks);

          // Listen for BDEs
          const unsubBDEs = onSnapshot(collection(db, 'bdes'), (snap) => {
            const bdeList = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as BDE));
            setBDES(bdeList);
          }, (err) => {
            console.error("BDEs listener error:", err);
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'bdes');
            }
          });
          unsubsRef.current.push(unsubBDEs);

          // Listen for Sales
          const salesQuery = isUserPrivileged
            ? collection(db, 'sales')
            : query(collection(db, 'sales'), 
                or(
                  where('agentEmail', '==', currentUser.email || ''),
                  where('agent', '==', currentProfile.displayName)
                ));

          const unsubSales = onSnapshot(salesQuery, (snap) => {
            const salesList = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Sales));
            setSales(salesList);
          }, (err) => {
            console.error("Sales listener error:", err);
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'sales');
            }
          });
          unsubsRef.current.push(unsubSales);

          // Listen for employees
          const employeesQuery = isUserPrivileged
            ? collection(db, 'employees')
            : query(collection(db, 'employees'), where('email', '==', currentUser.email || ''));

          const unsubEmployees = onSnapshot(employeesQuery, (snap) => {
            const emps = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Employee));
            setEmployees(emps);
          }, (err) => {
            console.error("Employees listener error:", err);
            if (auth.currentUser) {
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'employees');
            }
          });
          unsubsRef.current.push(unsubEmployees);

          if (isUserAdmin) {
            console.log("User is admin, starting admin listeners...");
            // Admin-only listeners (none currently, but kept for structure)
          }
        } else {
          setProfile(null);
        }
      } catch (err) {
        console.error("Auth sync error:", err);
        setFatalError(err);
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

  const handleLogin = async () => {
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (err) {
      console.error('Login error:', err);
    }
  };

  const handleLogout = async () => {
    await signOut(auth);
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
    const { id, type } = deleteConfirmation;
    
    if (type === 'trainingMaterials' && !hasPermission('training')) return;
    if (type !== 'trainingMaterials' && !isAdmin) return;
    
    try {
      const collectionName = type === 'employee' ? 'employees' : type === 'week' ? 'weeks' : type === 'bde' ? 'bdes' : type === 'trainingMaterials' ? 'trainingMaterials' : 'sales';
      await deleteDoc(doc(db, collectionName, id));
      setSuccess(`${type.charAt(0).toUpperCase() + type.slice(1)} deleted successfully`);
      setDeleteConfirmation(null);
    } catch (err: any) {
      const collectionName = type === 'employee' ? 'employees' : type === 'week' ? 'weeks' : type === 'bde' ? 'bdes' : type === 'trainingMaterials' ? 'trainingMaterials' : 'sales';
      handleFirestoreError(err, OperationType.DELETE, `${collectionName}/${id}`);
      setError(`Failed to delete ${type}`);
    }
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

  const handleCreateSales = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!hasPermission('sales')) return;
    
    setIsSavingSales(true);
    setError(null);
    setSuccess(null);

    try {
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
    setSalesForm(sale);
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
    if (isPrivileged) {
      return [...allUsers].sort((a, b) => b.pointsBalance - a.pointsBalance).slice(0, 5);
    }
    // Non-privileged users only see themselves in the leaderboard
    return allUsers.filter(u => u.uid === profile?.uid);
  }, [allUsers, profile, isPrivileged]);

  const allMonths = useMemo(() => {
    const months = Array.from(new Set(weeks.map(w => w.month))) as string[];
    const monthOrder = [
      'January', 'February', 'March', 'April', 'May', 'June',
      'July', 'August', 'September', 'October', 'November', 'December'
    ];
    return months.sort((a, b) => monthOrder.indexOf(a) - monthOrder.indexOf(b));
  }, [weeks]);

  const allSortedWeeks = useMemo(() => {
    return [...weeks].sort((a, b) => a.startDate.localeCompare(b.startDate));
  }, [weeks]);

  const sortedWeeks = useMemo(() => {
    return allSortedWeeks.slice(0, 12);
  }, [allSortedWeeks]);

  useEffect(() => {
    if (!hasSetDefaultMonth) {
      const lastMonthDate = new Date();
      lastMonthDate.setMonth(lastMonthDate.getMonth() - 1);
      const lastMonthName = lastMonthDate.toLocaleString('default', { month: 'long' });
      setSalesMonthFilter(lastMonthName);
      setHasSetDefaultMonth(true);
    }
  }, [hasSetDefaultMonth]);

  useEffect(() => {
    if (allSortedWeeks.length > 0 && !hasSetDefaultWeek) {
      // Only set default week if it matches the current month filter? 
      // Or maybe we don't need default week if we have default month.
      // The user specifically asked for default month.
      // Let's keep the default week logic but maybe it should be secondary.
      setSalesWeekFilter(allSortedWeeks[allSortedWeeks.length - 1].weekName);
      setHasSetDefaultWeek(true);
    }
  }, [allSortedWeeks, hasSetDefaultWeek]);

  const filteredSales = useMemo(() => {
    return sales.filter(sale => {
      const matchesAgentFilter = !salesAgentFilter || sale.agent === salesAgentFilter;
      const matchesBde = !salesBdeFilter || sale.bde === salesBdeFilter;
      const matchesWeek = !salesWeekFilter || sale.week === salesWeekFilter;
      
      // Find the month for the sale's week
      const saleWeekObj = weeks.find(w => w.weekName === sale.week);
      const matchesMonth = !salesMonthFilter || (saleWeekObj && saleWeekObj.month === salesMonthFilter);
      
      // Privilege restriction: only see own sales if not privileged
      const canSeeSale = isPrivileged || 
        (!sale.agentEmail && !sale.agent) || // If no owner, visible to all with permission
        sale.agentEmail === profile?.email || 
        sale.agent === profile?.displayName;

      return matchesAgentFilter && matchesBde && matchesWeek && matchesMonth && canSeeSale;
    });
  }, [sales, salesAgentFilter, salesBdeFilter, salesWeekFilter, salesMonthFilter, weeks, profile, isPrivileged]);

  const totalPkgValue = useMemo(() => {
    return filteredSales.reduce((sum, sale) => sum + (sale.packageValue || 0), 0);
  }, [filteredSales]);

  const inhouseSalesCount = useMemo(() => filteredSales.filter(s => !s.salesBy || s.salesBy === 'Inhouse').length, [filteredSales]);
  const branchSalesCount = useMemo(() => filteredSales.filter(s => s.salesBy === 'Branch').length, [filteredSales]);
  
  const inhousePkgValue = useMemo(() => 
    filteredSales.filter(s => !s.salesBy || s.salesBy === 'Inhouse').reduce((sum, s) => sum + (s.packageValue || 0), 0), 
    [filteredSales]
  );
  const branchPkgValue = useMemo(() => 
    filteredSales.filter(s => s.salesBy === 'Branch').reduce((sum, s) => sum + (s.packageValue || 0), 0), 
    [filteredSales]
  );

  if (fatalError) {
    return (
      <div className="min-h-screen bg-zinc-50 flex flex-col items-center justify-center p-4 text-center">
        <AlertCircle className="w-12 h-12 text-red-600 mb-4" />
        <h1 className="text-2xl font-bold text-zinc-900 mb-2">Application Error</h1>
        <p className="text-zinc-600 mb-6 max-w-md">{fatalError.message || "An unexpected error occurred. Please try refreshing the page."}</p>
        <Button onClick={() => window.location.reload()}>Refresh Page</Button>
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
    const data = sales.map(s => ({
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

  const downloadWeeklyOverview = () => {
    const activeEmployees = employees.filter(e => e.status === 'Active');
    const data = activeEmployees.map(emp => {
      const agentSales = sales.filter(s => s.agent === emp.name);
      
      const getBravoAchievement = (numWeeks: number) => {
        const targetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks);
        let totalSales = 0;
        let totalTarget = 0;
        const weeklyBaseTarget = emp.target / 4;

        targetWeeks.forEach(wk => {
          const wkSales = agentSales.filter(s => s.week === wk.weekName).reduce((sum, s) => sum + (s.packageValue || 0), 0);
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
        const weekSales = agentSales.filter(s => s.week === w.weekName).reduce((sum, s) => sum + (s.packageValue || 0), 0);
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
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 bg-orange-600 rounded-lg flex items-center justify-center">
              <Award className="w-5 h-5 text-white" />
            </div>
            <span className="font-bold text-lg tracking-tight">Nitsa Bravo</span>
          </div>

          {profile && (
            <div className="relative" ref={menuRef}>
              <button className="p-2 bg-zinc-100 rounded-lg" onClick={() => setIsMenuOpen(!isMenuOpen)}>
                <Menu className="w-6 h-6 text-zinc-700" />
              </button>
              {isMenuOpen && (
                <div className="absolute right-0 top-full mt-2 w-max bg-white border border-zinc-200 rounded-xl shadow-xl p-6 flex gap-8 z-50">
                  <div className="flex flex-col gap-2">
                    <h4 className="font-bold text-zinc-900 border-b pb-2 mb-2">Masters</h4>
                    {hasPermission('employees') && (
                      <button onClick={() => { setActiveTab('employees'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Employee master</button>
                    )}
                    {hasPermission('weeks') && (
                      <button onClick={() => { setActiveTab('weeks'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Weeks master</button>
                    )}
                    {hasPermission('bdes') && (
                      <button onClick={() => { setActiveTab('bdes'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">BDE master</button>
                    )}
                    {hasPermission('sales') && (
                      <button onClick={() => { setActiveTab('sales'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Sales Master</button>
                    )}
                    {hasPermission('training') && (
                      <button onClick={() => { setActiveTab('training'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Training Master</button>
                    )}
                  </div>
                  <div className="flex flex-col gap-2">
                    <h4 className="font-bold text-zinc-900 border-b pb-2 mb-2">Review</h4>
                    {hasPermission('weeklyOverview') && (
                      <button onClick={() => { setActiveTab('weeklyOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Weekly overview</button>
                    )}
                    {hasPermission('agentOverview') && (
                      <button onClick={() => { setActiveTab('agentOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Agent Overview</button>
                    )}
                    {hasPermission('trainingOverview') && (
                      <button onClick={() => { setActiveTab('trainingOverview'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">Training Overview</button>
                    )}
                  </div>
                  <div className="flex flex-col gap-2">
                    <h4 className="font-bold text-zinc-900 border-b pb-2 mb-2">Management</h4>
                    {hasPermission('userManagement') && (
                      <button onClick={() => { setActiveTab('userManagement'); setIsMenuOpen(false); }} className="text-sm text-zinc-600 hover:text-orange-600 text-left">User management</button>
                    )}
                  </div>
                </div>
              )}
            </div>
          )}
          
          <div className="flex items-center gap-4">
            <div className="hidden md:flex items-center gap-2 bg-zinc-100 px-3 py-1.5 rounded-full">
              <Star className="w-4 h-4 text-orange-500 fill-orange-500" />
              <span className="text-sm font-bold">{profile?.pointsBalance || 0} Points</span>
            </div>
            <div className="flex items-center gap-3">
              <img 
                src={profile?.photoURL} 
                className="w-8 h-8 rounded-full border border-zinc-200" 
                alt={profile?.displayName} 
                referrerPolicy="no-referrer"
              />
              <Button onClick={handleLogout} variant="ghost" className="p-2">
                <LogOut className="w-5 h-5" />
              </Button>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 py-8">
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
                {isPrivileged && (
                  <Button 
                    onClick={downloadEmployeeMaster}
                    variant="outline"
                    className="border-zinc-200"
                  >
                    <Download className="w-4 h-4" />
                    Download Excel
                  </Button>
                )}
                {isPrivileged && (
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
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={employeeForm.bde}
                      onChange={(e) => setEmployeeForm({...employeeForm, bde: e.target.value})}
                    >
                      <option value="None">None</option>
                      {bdes.map(bde => (
                        <option key={bde.id} value={bde.name}>{bde.name}</option>
                      ))}
                    </select>
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
                              {hasPermission('employees') && profile?.role !== 'Agent' && (
                                <button 
                                  onClick={() => handleEditEmployee(emp)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {isAdmin && (
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
              {isPrivileged && (
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
                              {hasPermission('weeks') && (
                                <button 
                                  onClick={() => handleEditWeek(wk)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {isAdmin && (
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
              {isPrivileged && (
                <Button 
                  onClick={() => {
                    setIsEditingBDE(false);
                    setBDEForm({ name: '' });
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
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {bdes.length === 0 ? (
                      <tr>
                        <td colSpan={2} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No BDE records found.
                        </td>
                      </tr>
                    ) : (
                      bdes.map(bde => (
                        <tr key={bde.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-bold text-orange-600 border border-zinc-200">{bde.name}</td>
                          <td className="px-4 py-3 border border-zinc-200">
                            <div className="flex items-center gap-2">
                              {hasPermission('bdes') && (
                                <button 
                                  onClick={() => handleEditBDE(bde)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {isAdmin && (
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
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Weekly Overview</h2>
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
                  {employees.filter(e => e.status === 'Active').map(emp => {
                    const agentSales = sales.filter(s => s.agent === emp.name);
                    
                    const getBravoAchievement = (numWeeks: number) => {
                      const targetWeeks = [...allSortedWeeks].reverse().slice(0, numWeeks);
                      let totalSales = 0;
                      let totalTarget = 0;
                      const weeklyBaseTarget = emp.target / 4;

                      targetWeeks.forEach(wk => {
                        const wkSales = agentSales.filter(s => s.week === wk.weekName).reduce((sum, s) => sum + (s.packageValue || 0), 0);
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
                        <td className="px-4 py-3 text-sm font-bold text-blue-600 border border-zinc-200 bg-blue-50/10 text-center">{bravo4}%</td>
                        <td className="px-4 py-3 text-sm font-bold text-purple-600 border border-zinc-200 bg-purple-50/10 text-center">{bravo8}%</td>
                        {[...sortedWeeks].reverse().map((w) => {
                          const weekSales = agentSales.filter(s => s.week === w.weekName).reduce((sum, s) => sum + s.packageValue, 0);
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
                              <td className="px-4 py-3 text-sm border border-zinc-200 text-right">
                                {weekSales > 0 ? weekSales.toLocaleString() : '-'}
                              </td>
                            </React.Fragment>
                          );
                        })}
                      </tr>
                    );
                  })}
                  {employees.filter(e => e.status === 'Active').length === 0 && (
                    <tr>
                      <td colSpan={6 + sortedWeeks.length * 2} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                        No active agents found.
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </Card>
          </div>
        )}
        {activeTab === 'sales' && hasPermission('sales') && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Sales Master</h2>
              </div>
              <div className="flex items-center gap-4">
                {isPrivileged && (
                  <Button 
                    onClick={downloadSalesMaster}
                    variant="outline"
                    className="flex items-center gap-2"
                  >
                    <Download className="w-4 h-4" />
                    Download Excel
                  </Button>
                )}
                {isPrivileged && (
                  <Button 
                    onClick={() => {
                      setIsEditingSales(false);
                      setSalesForm({
                        week: '',
                        date: new Date().toISOString().split('T')[0],
                        guestName: '',
                        agent: profile?.role === 'Agent' ? profile?.displayName || '' : '',
                        agentEmail: profile?.role === 'Agent' ? profile?.email || '' : '',
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
                    <select 
                      required
                      disabled={profile?.role === 'Agent'}
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none disabled:opacity-50"
                      value={salesForm.agent || ''}
                      onChange={(e) => setSalesForm({...salesForm, agent: e.target.value})}
                    >
                      <option value="">Select Agent</option>
                      {employees.map(emp => (
                        <option key={emp.id} value={emp.name}>{emp.name}</option>
                      ))}
                    </select>
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
                      <option value="70">70</option>
                      <option value="100">100</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">BDE</label>
                    <select 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.bde || ''}
                      onChange={(e) => setSalesForm({...salesForm, bde: e.target.value})}
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
                      onChange={(e) => setSalesForm({...salesForm, associateBde: e.target.value})}
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
                      <option value="Done">Done</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Sales By</label>
                    <select 
                      required
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.salesBy || 'Inhouse'}
                      onChange={(e) => setSalesForm({...salesForm, salesBy: e.target.value as 'Inhouse' | 'Branch'})}
                    >
                      <option value="Inhouse">Inhouse</option>
                      <option value="Branch">Branch</option>
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
                  <select 
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={salesAgentFilter}
                    onChange={(e) => setSalesAgentFilter(e.target.value)}
                  >
                    {isPrivileged && <option value="">All Agents</option>}
                    {employees
                      .filter(e => isPrivileged || (e.email && e.email.toLowerCase() === profile?.email?.toLowerCase()))
                      .map(emp => (
                      <option key={emp.id} value={emp.name}>{emp.name}</option>
                    ))}
                  </select>
                </div>
                <div className="flex-1 space-y-1">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Filter by BDE</label>
                  <select 
                    className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                    value={salesBdeFilter}
                    onChange={(e) => setSalesBdeFilter(e.target.value)}
                  >
                    <option value="">All BDEs</option>
                    {bdes.map(bde => (
                      <option key={bde.id} value={bde.name}>{bde.name}</option>
                    ))}
                  </select>
                </div>
                <Button 
                  variant="ghost" 
                  onClick={() => { setSalesAgentFilter(''); setSalesBdeFilter(''); setSalesWeekFilter(''); }}
                  className="text-zinc-500"
                >
                  <Filter className="w-4 h-4 mr-2" />
                  Clear Filters
                </Button>
              </div>

              {/* Summary Cards */}
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
                <Card className="p-4 bg-orange-50 border-orange-100">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-orange-100 rounded-lg">
                      <TrendingUp className="w-5 h-5 text-orange-600" />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Inhouse Sales</p>
                      <p className="text-2xl font-bold text-zinc-900">{inhouseSalesCount}</p>
                    </div>
                  </div>
                </Card>
                <Card className="p-4 bg-orange-50 border-orange-100">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-orange-100 rounded-lg">
                      <TrendingUp className="w-5 h-5 text-orange-600" />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-orange-600 uppercase tracking-wider">Branch Sales</p>
                      <p className="text-2xl font-bold text-zinc-900">{branchSalesCount}</p>
                    </div>
                  </div>
                </Card>
                <Card className="p-4 bg-zinc-900 text-white">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-zinc-800 rounded-lg">
                      <IndianRupee className="w-5 h-5 text-orange-400" />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Inhouse Value</p>
                      <p className="text-2xl font-bold text-white">₹{inhousePkgValue.toLocaleString()}</p>
                    </div>
                  </div>
                </Card>
                <Card className="p-4 bg-zinc-900 text-white">
                  <div className="flex items-center gap-3">
                    <div className="p-2 bg-zinc-800 rounded-lg">
                      <IndianRupee className="w-5 h-5 text-orange-400" />
                    </div>
                    <div>
                      <p className="text-[10px] font-bold text-zinc-400 uppercase tracking-wider">Branch Value</p>
                      <p className="text-2xl font-bold text-white">₹{branchPkgValue.toLocaleString()}</p>
                    </div>
                  </div>
                </Card>
              </div>

              <h3 className="text-lg font-bold">Sales Records</h3>
              <Card className="overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[2000px] border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50">
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase sticky left-0 bg-zinc-50 z-10 border border-zinc-200">Guest Name</th>
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
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredSales.length === 0 ? (
                      <tr>
                        <td colSpan={27} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No sales records found.
                        </td>
                      </tr>
                    ) : (
                      filteredSales.map(sale => (
                        <tr key={sale.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-bold text-orange-600 sticky left-0 bg-white z-10 border border-zinc-200">{sale.guestName}</td>
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
                              sale.salesBy === 'Branch' ? "bg-purple-100 text-purple-700" : "bg-blue-100 text-blue-700"
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
                          <td className="px-4 py-3 border border-zinc-200">
                            <div className="flex items-center gap-2">
                              {hasPermission('sales') && (
                                <button 
                                  onClick={() => handleEditSales(sale)}
                                  className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                                >
                                  <Edit2 className="w-4 h-4" />
                                </button>
                              )}
                              {isAdmin && (
                                <button 
                                  onClick={() => sale.id && handleDeleteSales(sale.id, sale.guestName)}
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
        {activeTab === 'agentOverview' && hasPermission('agentOverview') && (
          <div className="space-y-8">
            <div className="flex items-center gap-2">
              <UserIcon className="w-6 h-6 text-orange-600" />
              <h2 className="text-3xl font-bold tracking-tight">Agent Overview</h2>
            </div>
            
            <div className="bg-white p-6 rounded-xl border border-zinc-200 shadow-sm space-y-4">
              <label className="text-sm font-bold text-zinc-700">Select Agent</label>
              <select 
                className="w-full md:w-1/3 px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                value={selectedAgentOverview}
                onChange={(e) => setSelectedAgentOverview(e.target.value)}
              >
                <option value="">-- Select an Agent --</option>
                {employees
                  .filter(e => e.status === 'Active')
                  .filter(e => isPrivileged || (e.email && e.email.toLowerCase() === profile?.email?.toLowerCase()))
                  .map(emp => (
                  <option key={emp.id} value={emp.id}>{emp.name} ({emp.employeeCode})</option>
                ))}
              </select>
            </div>

            {selectedAgentOverview && employees.find(e => e.id === selectedAgentOverview) && (() => {
              const agent = employees.find(e => e.id === selectedAgentOverview)!;
              const agentSales = sales.filter(s => s.agent === agent.name);
              
              // Helper to get sales in the last N weeks from master
              const reversedWeeks = [...allSortedWeeks].reverse();
              const getSalesByWeeks = (numWeeks: number) => {
                const targetWeeks = reversedWeeks.slice(0, numWeeks).map(w => w.weekName);
                const filteredSales = agentSales.filter(s => targetWeeks.includes(s.week || ''));
                return {
                  amount: filteredSales.reduce((sum, s) => sum + (s.packageValue || 0), 0),
                  count: filteredSales.length
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
                const weekSales = agentSales
                  .filter(s => s.week === wk.weekName)
                  .reduce((sum, s) => sum + (s.packageValue || 0), 0);
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
                
                const weekSales = agentSales
                  .filter(s => s.week === wk.weekName)
                  .reduce((sum, s) => sum + (s.packageValue || 0), 0);
                
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
                monthsMap[month].sales += agentSales
                  .filter(s => s.week === wk.weekName)
                  .reduce((sum, s) => sum + (s.packageValue || 0), 0);
                
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
                  {/* 8 Weeks Performance Highlight */}
                  <Card className={cn("p-8 border-2 shadow-lg transition-all hover:scale-[1.01]", status.bg, status.border)}>
                    <div className="grid grid-cols-1 lg:grid-cols-3 gap-8 items-center">
                      <div className="space-y-2 text-center lg:text-left lg:border-r lg:border-zinc-200 pr-0 lg:pr-8">
                        <p className={cn("text-sm font-bold uppercase tracking-widest", status.color)}>8 Weeks Bravo Point Achievement</p>
                        <div className="flex items-center justify-center lg:justify-start gap-4">
                          <h1 className={cn("text-7xl font-black tracking-tighter", status.color)}>
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
                        <h2 className={cn("text-5xl font-black tracking-tighter", eightWeekTotals.contribution >= 0 ? "text-green-600" : "text-red-600")}>
                          ₹{eightWeekTotals.contribution.toLocaleString()}
                        </h2>
                        <p className="text-zinc-500 text-[10px] font-medium">Net contribution to company after workstation costs</p>
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
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4">
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
                            <span className="text-lg font-bold text-orange-600">{weeklyTarget > 0 ? Math.round((weeklySalesData.amount / weeklyTarget) * 100) : 0}%</span>
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
                            <span className="text-lg font-bold text-blue-600">{fourWeekTarget > 0 ? Math.round((fourWeekSalesData.amount / fourWeekTarget) * 100) : 0}%</span>
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
                            <span className="text-lg font-bold text-purple-600">{eightWeekTarget > 0 ? Math.round((eightWeekSalesData.amount / eightWeekTarget) * 100) : 0}%</span>
                          </div>
                          <div className="w-full bg-purple-100 rounded-full h-2">
                            <div className="bg-purple-500 h-2 rounded-full" style={{ width: `${Math.min(100, eightWeekTarget > 0 ? (eightWeekSalesData.amount / eightWeekTarget) * 100 : 0)}%` }}></div>
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
                            <span className="text-lg font-bold text-emerald-600">{twentyFourWeekTarget > 0 ? Math.round((twentyFourWeekSalesData.amount / twentyFourWeekTarget) * 100) : 0}%</span>
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
                            const weekSales = agentSales
                              .filter(s => s.week === wk.weekName)
                              .reduce((sum, s) => sum + (s.packageValue || 0), 0);
                            const shortfall = Math.max(0, finalTarget - weekSales);
                            
                            const workstationCost = isJoined ? (agent.salary * 3) / 4 : 0;
                            const contribution = isJoined ? (weekSales * 0.12) - workstationCost : 0;

                            return (
                              <tr key={wk.id} className="hover:bg-zinc-50 transition-colors">
                                <td className="px-4 py-3 text-sm font-medium border border-zinc-200">{wk.weekName}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200">{wk.month}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200">₹{currentWeeklyBaseTarget.toLocaleString()}</td>
                                <td className="px-4 py-3 text-lg font-bold text-orange-600 border border-zinc-200 bg-orange-50">{percentageStr}</td>
                                <td className="px-4 py-3 text-sm font-bold border border-zinc-200">₹{finalTarget.toLocaleString()}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200 text-green-600 font-medium">₹{weekSales.toLocaleString()}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200 text-red-600 font-medium">₹{shortfall.toLocaleString()}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200">₹{workstationCost.toLocaleString()}</td>
                                <td className="px-4 py-3 text-sm border border-zinc-200 font-medium" style={{ color: contribution >= 0 ? '#16a34a' : '#dc2626' }}>₹{contribution.toLocaleString()}</td>
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

                        <div className={cn("p-4 rounded-lg border flex justify-between items-center", isThreeMonthConsecutiveEligible ? "bg-green-50 border-green-200" : "bg-zinc-50 border-zinc-200")}>
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
                        <div className="space-y-2">
                          <label className="text-xs font-bold text-zinc-400 uppercase">Hypothetical Weekly Sales (₹)</label>
                          <input 
                            type="number"
                            className="w-full px-4 py-2 bg-zinc-800 border border-zinc-700 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none text-white"
                            placeholder="Enter weekly sales..."
                            value={incentiveCalcSales || ''}
                            onChange={(e) => setIncentiveCalcSales(parseFloat(e.target.value) || 0)}
                          />
                        </div>
                        <div className="space-y-2">
                          <label className="text-xs font-bold text-zinc-400 uppercase">Hypothetical Monthly Sales (₹)</label>
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
                            <span className="text-lg font-bold text-orange-500">
                              ₹{(isWeeklyEligible ? incentiveCalcSales * 0.001 : 0).toLocaleString()} approx
                            </span>
                          </div>
                          <div className="flex justify-between items-center">
                            <span className="text-sm text-zinc-400">Potential Monthly Incentive:</span>
                            <div className="text-right">
                              <p className="text-lg font-bold text-orange-500">
                                ₹{(incentiveCalcMonthlySales >= fourWeekTarget ? ((incentiveCalcMonthlySales - fourWeekTarget) * 0.03 + 2000) : 0).toLocaleString()} approx
                              </p>
                              {incentiveCalcMonthlySales >= fourWeekTarget && (
                                <p className="text-[10px] text-zinc-500">
                                  Breakdown: ₹{((incentiveCalcMonthlySales - fourWeekTarget) * 0.03).toLocaleString()} (3%) + ₹2,000
                                </p>
                              )}
                            </div>
                          </div>
                          <div className="flex justify-between items-center pt-2 border-t border-zinc-700">
                            <span className="font-bold">Total Potential Earnings:</span>
                            <span className="text-2xl font-black text-green-400">
                              ₹{(
                                (isWeeklyEligible ? incentiveCalcSales * 0.001 : 0) + 
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
                </div>
              );
            })()}
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

            <Card className="overflow-x-auto">
              <table className="w-full text-left border-collapse border border-zinc-200">
                <thead>
                  <tr className="bg-zinc-50">
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">User</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Email</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Role</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Permissions</th>
                    <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                  </tr>
                </thead>
                <tbody>
                  {allUsers.map(u => (
                    <tr key={u.uid} className="hover:bg-zinc-50 transition-colors">
                      <td className="px-4 py-3 border border-zinc-200">
                        <div className="flex items-center gap-3">
                          <img src={u.photoURL} className="w-8 h-8 rounded-full" alt="" referrerPolicy="no-referrer" />
                          <span className="text-sm font-bold">{u.displayName}</span>
                        </div>
                      </td>
                      <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{u.email}</td>
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
                          {Array.isArray(u.permissions) && u.permissions.length > 0 ? (
                            u.permissions.map(pId => {
                              const p = allPermissions.find(ap => ap.id === pId);
                              return (
                                <Badge key={pId} className="bg-zinc-50 text-zinc-400 lowercase text-[9px]">
                                  {p?.label || pId}
                                </Badge>
                              );
                            })
                          ) : null}
                        </div>
                      </td>
                      <td className="px-4 py-3 border border-zinc-200">
                        <button 
                          onClick={() => handleEditUser(u)}
                          className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                        >
                          <Edit2 className="w-4 h-4" />
                        </button>
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
                          permissions: roleObj ? roleObj.permissions : []
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
                    <div className="grid grid-cols-2 gap-2 p-4 bg-zinc-50 rounded-2xl border border-zinc-200">
                      {allPermissions.map(p => (
                        <div key={p.id} className="flex items-center gap-2">
                          <div className={cn(
                            "w-2 h-2 rounded-full",
                            (Array.isArray(userForm.permissions) && userForm.permissions.includes(p.id)) ? "bg-green-500" : "bg-zinc-300"
                          )} />
                          <span className="text-xs text-zinc-600">{p.label}</span>
                        </div>
                      ))}
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
                        <p className="text-xs text-zinc-500">{role.permissions.length} permissions</p>
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
                          onClick={async () => {
                            if (window.confirm(`Are you sure you want to delete the ${role.name} role?`)) {
                              try {
                                await deleteDoc(doc(db, 'roles', role.id));
                              } catch (err) {
                                console.error("Error deleting role:", err);
                              }
                            }
                          }}
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
                  const selectedPerms = allPermissions
                    .filter(p => formData.get(`perm-${p.id}`) === 'on')
                    .map(p => p.id);

                  try {
                    if (selectedRole) {
                      await updateDoc(doc(db, 'roles', selectedRole.id), {
                        name,
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
                    <label className="text-xs font-bold text-zinc-500 uppercase">Permissions</label>
                    <div className="grid grid-cols-2 gap-3 p-4 bg-zinc-50 rounded-2xl border border-zinc-200 max-h-60 overflow-y-auto">
                      {allPermissions.map(p => (
                        <label key={p.id} className="flex items-center gap-3 cursor-pointer group">
                          <input 
                            name={`perm-${p.id}`}
                            type="checkbox"
                            defaultChecked={selectedRole?.permissions.includes(p.id)}
                            className="w-4 h-4 rounded border-zinc-300 text-orange-600 focus:ring-orange-500"
                          />
                          <span className="text-sm text-zinc-600 group-hover:text-zinc-900 transition-colors">{p.label}</span>
                        </label>
                      ))}
                    </div>
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
                <div className="flex items-center gap-2">
                  <label className="text-xs font-bold text-zinc-500 uppercase">Agent</label>
                  <select 
                    className="px-3 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none focus:ring-2 focus:ring-orange-500"
                    value={trainingAgentFilter}
                    onChange={(e) => setTrainingAgentFilter(e.target.value)}
                  >
                    <option value="">All Agents</option>
                    {employees.filter(e => e.status === 'Active').map(e => (
                      <option key={e.id} value={e.name}>{e.name}</option>
                    ))}
                  </select>
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
              const filteredEmployees = employees.filter(e => trainingEmployeeStatusFilter === 'All' ? true : e.status === trainingEmployeeStatusFilter);
              
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
              {hasPermission('training') && (
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
                          {hasPermission('training') && (
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
      </div>
  );
}
