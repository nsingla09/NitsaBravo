/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import { useState, useEffect, useMemo, Component } from 'react';
import React from 'react';
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
  orderBy, 
  limit, 
  onSnapshot, 
  addDoc, 
  serverTimestamp,
  runTransaction,
  handleFirestoreError,
  OperationType
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
  Trash2
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
interface UserProfile {
  uid: string;
  email: string;
  displayName: string;
  photoURL: string;
  pointsBalance: number;
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

interface Employee {
  id?: string;
  employeeCode: string;
  name: string;
  target: number;
  position: string;
  joiningDate: string;
  joiningWeek: string;
  bde: string;
  status: 'Active' | 'Deactive';
  week1Target: number;
  week2Target: number;
  week3Target: number;
  week4Target: number;
  week5Target: number;
  week6Target: number;
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

interface Sales {
  id?: string;
  week: string;
  date: string;
  guestName: string;
  agent: string;
  bde: string;
  associateBde: string;
  percentage: string;
  destination: string;
  tripId: string;
  tripDate: string;
  packageValue: number;
  lessThan10Percent20k: string;
  ppMargin: number;
  noOfPax: number;
  totalMargin: number;
  flight: number;
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

const Badge = ({ children, className }: { children: React.ReactNode; className?: string }) => (
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
  const [transactions, setTransactions] = useState<Transaction[]>([]);
  const [employees, setEmployees] = useState<Employee[]>([]);
  const [weeks, setWeeks] = useState<Week[]>([]);
  const [bdes, setBDES] = useState<BDE[]>([]);
  const [sales, setSales] = useState<Sales[]>([]);
  const [activeTab, setActiveTab] = useState<'dashboard' | 'employees' | 'weeks' | 'bdes' | 'sales'>('dashboard');
  const [isLoading, setIsLoading] = useState(true);
  const [isAwarding, setIsAwarding] = useState(false);
  const [isSavingEmployee, setIsSavingEmployee] = useState(false);
  const [isSavingWeek, setIsSavingWeek] = useState(false);
  const [isSavingBDE, setIsSavingBDE] = useState(false);
  const [isSavingSales, setIsSavingSales] = useState(false);
  const [isEditingEmployee, setIsEditingEmployee] = useState(false);
  const [isEditingWeek, setIsEditingWeek] = useState(false);
  const [isEditingBDE, setIsEditingBDE] = useState(false);
  const [isEditingSales, setIsEditingSales] = useState(false);
  const [isEmployeeModalOpen, setIsEmployeeModalOpen] = useState(false);
  const [isWeekModalOpen, setIsWeekModalOpen] = useState(false);
  const [isBDEModalOpen, setIsBDEModalOpen] = useState(false);
  const [isSalesModalOpen, setIsSalesModalOpen] = useState(false);
  const [searchQuery, setSearchQuery] = useState('');
  const [selectedUser, setSelectedUser] = useState<UserProfile | null>(null);
  const [pointsToAward, setPointsToAward] = useState(10);
  const [reason, setReason] = useState('');
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [fatalError, setFatalError] = useState<any>(null);
  const [deleteConfirmation, setDeleteConfirmation] = useState<{
    id: string;
    type: 'employee' | 'week' | 'bde' | 'sales';
    name: string;
  } | null>(null);

  if (fatalError) throw fatalError;

  // Employee Form State
  const [employeeForm, setEmployeeForm] = useState<Partial<Employee>>({
    status: 'Active',
    position: 'Sales Agent',
    joiningWeek: 'Week 1',
    bde: 'None',
    week1Target: 25,
    week2Target: 25,
    week3Target: 50,
    week4Target: 50,
    week5Target: 50,
    week6Target: 50,
  });

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
    bde: '',
    associateBde: '',
    percentage: '',
    destination: '',
    tripId: '',
    tripDate: '',
    packageValue: 0,
    lessThan10Percent20k: '',
    ppMargin: 0,
    noOfPax: 0,
    totalMargin: 0,
    flight: 0,
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

  const isAdmin = profile?.email === 'nsingla09@gmail.com';

  // --- Auth & Profile Sync ---
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      try {
        setUser(currentUser);
        console.log("Auth state changed:", currentUser?.email, currentUser?.uid);
        if (currentUser) {
          const userRef = doc(db, 'users', currentUser.uid);
          console.log("Fetching user profile for:", currentUser.uid);
          const userSnap = await getDoc(userRef);
          
          let currentProfile: UserProfile;
          if (!userSnap.exists()) {
            console.log("User profile not found, creating new profile...");
            currentProfile = {
              uid: currentUser.uid,
              email: currentUser.email || '',
              displayName: currentUser.displayName || 'Anonymous User',
              photoURL: currentUser.photoURL || '',
              pointsBalance: 100, // Starting points
            };
            await setDoc(userRef, currentProfile);
            setProfile(currentProfile);
          } else {
            console.log("User profile found:", userSnap.data());
            currentProfile = userSnap.data() as UserProfile;
            setProfile(currentProfile);
          }

          const isUserAdmin = currentProfile.email === 'nsingla09@gmail.com';

          console.log("Starting Firestore listeners...");
          const unsubProfile = onSnapshot(userRef, (doc) => {
            setProfile(doc.data() as UserProfile);
          }, (err) => {
            console.error("Profile listener error:", err);
            setFatalError(err);
            handleFirestoreError(err, OperationType.GET, `users/${currentUser.uid}`);
          });

          // Listen for all users
          const unsubUsers = onSnapshot(collection(db, 'users'), (snap) => {
            const users = snap.docs.map(doc => doc.data() as UserProfile);
            setAllUsers(users);
          }, (err) => {
            console.error("Users listener error:", err);
            setFatalError(err);
            handleFirestoreError(err, OperationType.GET, 'users');
          });

          // Listen for transactions
          const q = query(collection(db, 'transactions'), orderBy('createdAt', 'desc'), limit(20));
          const unsubTrans = onSnapshot(q, (snap) => {
            const trans = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Transaction));
            setTransactions(trans);
          }, (err) => {
            console.error("Transactions listener error:", err);
            setFatalError(err);
            handleFirestoreError(err, OperationType.GET, 'transactions');
          });

          let unsubEmployees = () => {};
          let unsubWeeks = () => {};
          let unsubBDEs = () => {};
          let unsubSales = () => {};

          if (isUserAdmin) {
            console.log("User is admin, starting admin listeners...");
            // Listen for employees
            unsubEmployees = onSnapshot(collection(db, 'employees'), (snap) => {
              const emps = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Employee));
              setEmployees(emps);
            }, (err) => {
              console.error("Employees listener error:", err);
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'employees');
            });

            // Listen for weeks
            unsubWeeks = onSnapshot(collection(db, 'weeks'), (snap) => {
              const wks = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Week));
              setWeeks(wks);
            }, (err) => {
              console.error("Weeks listener error:", err);
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'weeks');
            });

            // Listen for BDEs
            unsubBDEs = onSnapshot(collection(db, 'bdes'), (snap) => {
              const bdeList = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as BDE));
              setBDES(bdeList);
            }, (err) => {
              console.error("BDEs listener error:", err);
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'bdes');
            });

            // Listen for Sales
            unsubSales = onSnapshot(collection(db, 'sales'), (snap) => {
              const salesList = snap.docs.map(doc => ({ id: doc.id, ...doc.data() } as Sales));
              setSales(salesList);
            }, (err) => {
              console.error("Sales listener error:", err);
              setFatalError(err);
              handleFirestoreError(err, OperationType.GET, 'sales');
            });
          }

          return () => {
            unsubProfile();
            unsubUsers();
            unsubTrans();
            unsubEmployees();
            unsubWeeks();
            unsubBDEs();
            unsubSales();
          };
        } else {
          setProfile(null);
        }
      } catch (err) {
        console.error("Auth sync error:", err);
        setFatalError(err);
      } finally {
        setIsLoading(false);
      }
    });

    return () => unsubscribe();
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
    if (!isAdmin) return;
    
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
        week1Target: 25,
        week2Target: 25,
        week3Target: 50,
        week4Target: 50,
        week5Target: 50,
        week6Target: 50,
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
    if (!deleteConfirmation || !isAdmin) return;
    const { id, type } = deleteConfirmation;
    
    try {
      const collectionName = type === 'employee' ? 'employees' : type === 'week' ? 'weeks' : type === 'bde' ? 'bdes' : 'sales';
      await deleteDoc(doc(db, collectionName, id));
      setSuccess(`${type.charAt(0).toUpperCase() + type.slice(1)} deleted successfully`);
      setDeleteConfirmation(null);
    } catch (err: any) {
      const collectionName = type === 'employee' ? 'employees' : type === 'week' ? 'weeks' : type === 'bde' ? 'bdes' : 'sales';
      handleFirestoreError(err, OperationType.DELETE, `${collectionName}/${id}`);
      setError(`Failed to delete ${type}`);
    }
  };

  const handleCreateWeek = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!isAdmin) return;
    
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
    if (!isAdmin) return;
    
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
    if (!isAdmin) return;
    
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
        bde: '',
        associateBde: '',
        percentage: '',
        destination: '',
        tripId: '',
        tripDate: '',
        packageValue: 0,
        lessThan10Percent20k: '',
        ppMargin: 0,
        noOfPax: 0,
        totalMargin: 0,
        flight: 0,
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

  const leaderboard = useMemo(() => {
    return [...allUsers].sort((a, b) => b.pointsBalance - a.pointsBalance).slice(0, 5);
  }, [allUsers]);

  if (isLoading) {
    return (
      <div className="min-h-screen bg-zinc-50 flex items-center justify-center">
        <Loader2 className="w-8 h-8 text-orange-600 animate-spin" />
      </div>
    );
  }

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

          {isAdmin && (
            <nav className="flex items-center gap-1 bg-zinc-100 p-1 rounded-xl">
              <button
                onClick={() => setActiveTab('dashboard')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'dashboard' ? "bg-white shadow-sm text-orange-600" : "text-zinc-500 hover:text-zinc-700"
                )}
              >
                Dashboard
              </button>
              <button
                onClick={() => setActiveTab('employees')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'employees' ? "bg-white shadow-sm text-orange-600" : "text-zinc-500 hover:text-zinc-700"
                )}
              >
                Employee Master
              </button>
              <button
                onClick={() => setActiveTab('weeks')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'weeks' ? "bg-white shadow-sm text-orange-600" : "text-zinc-500 hover:text-zinc-700"
                )}
              >
                Weeks Master
              </button>
              <button
                onClick={() => setActiveTab('bdes')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'bdes' ? "bg-white shadow-sm text-orange-600" : "text-zinc-500 hover:text-zinc-700"
                )}
              >
                BDE Master
              </button>
              <button
                onClick={() => setActiveTab('sales')}
                className={cn(
                  "px-4 py-1.5 rounded-lg text-sm font-bold transition-all",
                  activeTab === 'sales' ? "bg-white shadow-sm text-orange-600" : "text-zinc-500 hover:text-zinc-700"
                )}
              >
                Sales Master
              </button>
            </nav>
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
                <h2 className="text-3xl font-bold tracking-tight">Hello, {profile?.displayName.split(' ')[0]}! 👋</h2>
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
                  {transactions.length === 0 ? (
                    <Card className="p-12 text-center text-zinc-400">
                      <p>No transactions yet. Be the first to give a Bravo!</p>
                    </Card>
                  ) : (
                    transactions.map((t, idx) => (
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

        {activeTab === 'employees' && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Users className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Employee Master</h2>
              </div>
              <Button 
                onClick={() => {
                  setIsEditingEmployee(false);
                  setEmployeeForm({
                    status: 'Active',
                    position: 'Sales Agent',
                    joiningWeek: 'Week 1',
                    bde: 'None',
                    week1Target: 25,
                    week2Target: 25,
                    week3Target: 50,
                    week4Target: 50,
                    week5Target: 50,
                    week6Target: 50,
                  });
                  setIsEmployeeModalOpen(true);
                }} 
                className="bg-orange-600"
              >
                <Plus className="w-4 h-4" />
                Add Employee
              </Button>
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
                      {weeks.map(wk => (
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
                  <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 w-4 h-4 text-zinc-400" />
                    <input 
                      type="text" 
                      placeholder="Filter employees..."
                      className="pl-9 pr-4 py-1.5 bg-white border border-zinc-200 rounded-lg text-sm outline-none"
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
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Position</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Target</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Joined</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Status</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {employees.length === 0 ? (
                      <tr>
                        <td colSpan={7} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No employee records found.
                        </td>
                      </tr>
                    ) : (
                      employees.map(emp => (
                        <tr key={emp.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-mono text-orange-600 border border-zinc-200">{emp.employeeCode}</td>
                          <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{emp.name}</td>
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
                              <button 
                                onClick={() => handleEditEmployee(emp)}
                                className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                              >
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => emp.id && handleDeleteEmployee(emp.id, emp.name)}
                                className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
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

        {activeTab === 'weeks' && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <CalendarIcon className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Weeks Master</h2>
              </div>
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
                    {weeks.length === 0 ? (
                      <tr>
                        <td colSpan={5} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No week records found.
                        </td>
                      </tr>
                    ) : (
                      weeks.map(wk => (
                        <tr key={wk.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-bold text-orange-600 border border-zinc-200">{wk.weekName}</td>
                          <td className="px-4 py-3 text-sm font-bold border border-zinc-200">{wk.month}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{wk.startDate}</td>
                          <td className="px-4 py-3 text-sm text-zinc-600 border border-zinc-200">{wk.endDate}</td>
                          <td className="px-4 py-3 border border-zinc-200">
                            <div className="flex items-center gap-2">
                              <button 
                                onClick={() => handleEditWeek(wk)}
                                className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                              >
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => wk.id && handleDeleteWeek(wk.id, wk.weekName)}
                                className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
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

        {activeTab === 'bdes' && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <Users className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">BDE Master</h2>
              </div>
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
                              <button 
                                onClick={() => handleEditBDE(bde)}
                                className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                              >
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => bde.id && handleDeleteBDE(bde.id, bde.name)}
                                className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
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
        {activeTab === 'sales' && (
          <div className="space-y-8">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2">
                <TrendingUp className="w-6 h-6 text-orange-600" />
                <h2 className="text-3xl font-bold tracking-tight">Sales Master</h2>
              </div>
              <Button 
                onClick={() => {
                  setIsEditingSales(false);
                  setSalesForm({
                    week: '',
                    date: new Date().toISOString().split('T')[0],
                    guestName: '',
                    agent: '',
                    bde: '',
                    associateBde: '',
                    percentage: '',
                    destination: '',
                    tripId: '',
                    tripDate: '',
                    packageValue: 0,
                    lessThan10Percent20k: '',
                    ppMargin: 0,
                    noOfPax: 0,
                    totalMargin: 0,
                    flight: 0,
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
                      {weeks.map(wk => (
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
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
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
                    <input 
                      type="text" 
                      placeholder="Enter Associate BDE"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.associateBde || ''}
                      onChange={(e) => setSalesForm({...salesForm, associateBde: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Percentage</label>
                    <input 
                      type="text" 
                      placeholder="e.g. 70%"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.percentage || ''}
                      onChange={(e) => setSalesForm({...salesForm, percentage: e.target.value})}
                    />
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
                      value={salesForm.packageValue || 0}
                      onChange={(e) => setSalesForm({...salesForm, packageValue: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Less Than 10% / 20k</label>
                    <input 
                      type="text" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.lessThan10Percent20k || ''}
                      onChange={(e) => setSalesForm({...salesForm, lessThan10Percent20k: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">PP Margin</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.ppMargin || 0}
                      onChange={(e) => setSalesForm({...salesForm, ppMargin: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">No. of Pax</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.noOfPax || 0}
                      onChange={(e) => setSalesForm({...salesForm, noOfPax: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Total Margin</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.totalMargin || 0}
                      onChange={(e) => setSalesForm({...salesForm, totalMargin: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Flight</label>
                    <input 
                      type="number" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.flight || 0}
                      onChange={(e) => setSalesForm({...salesForm, flight: Number(e.target.value)})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Source</label>
                    <input 
                      type="text" 
                      placeholder="e.g. social"
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.source || ''}
                      onChange={(e) => setSalesForm({...salesForm, source: e.target.value})}
                    />
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
                      <option value="Pending">Pending</option>
                    </select>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Tasks Pending</label>
                    <input 
                      type="text" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.tasksPending || ''}
                      onChange={(e) => setSalesForm({...salesForm, tasksPending: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">P - Hotel</label>
                    <input 
                      type="text" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.pHotel || ''}
                      onChange={(e) => setSalesForm({...salesForm, pHotel: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">P Flight</label>
                    <input 
                      type="text" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.pFlight || ''}
                      onChange={(e) => setSalesForm({...salesForm, pFlight: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Land</label>
                    <input 
                      type="text" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.land || ''}
                      onChange={(e) => setSalesForm({...salesForm, land: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">H+F+L issue</label>
                    <input 
                      type="text" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.hflIssue || ''}
                      onChange={(e) => setSalesForm({...salesForm, hflIssue: e.target.value})}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-bold text-zinc-500 uppercase">Work Pending</label>
                    <input 
                      type="text" 
                      className="w-full px-4 py-2 bg-zinc-50 border border-zinc-200 rounded-lg focus:ring-2 focus:ring-orange-500 outline-none"
                      value={salesForm.workPending || ''}
                      onChange={(e) => setSalesForm({...salesForm, workPending: e.target.value})}
                    />
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
              <h3 className="text-lg font-bold">Sales Records</h3>
              <Card className="overflow-x-auto">
                <table className="w-full text-left border-collapse min-w-[2000px] border border-zinc-200">
                  <thead>
                    <tr className="bg-zinc-50">
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase sticky left-0 bg-zinc-50 z-10 border border-zinc-200">Guest Name</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Week</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Date</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Agent</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">BDE</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Assoc BDE</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">%</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Destination</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Trip ID</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Trip Date</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Pkg Value</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">PP Margin</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Pax</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Total Margin</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Flight</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Source</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Converted</th>
                      <th className="px-4 py-3 text-xs font-bold text-zinc-500 uppercase border border-zinc-200">Actions</th>
                    </tr>
                  </thead>
                  <tbody>
                    {sales.length === 0 ? (
                      <tr>
                        <td colSpan={18} className="px-4 py-12 text-center text-zinc-400 italic border border-zinc-200">
                          No sales records found.
                        </td>
                      </tr>
                    ) : (
                      sales.map(sale => (
                        <tr key={sale.id} className="hover:bg-zinc-50 transition-colors">
                          <td className="px-4 py-3 text-sm font-bold text-orange-600 sticky left-0 bg-white z-10 border border-zinc-200">{sale.guestName}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.week}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.date}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.agent}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.bde}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.associateBde}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.percentage}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.destination}</td>
                          <td className="px-4 py-3 text-sm font-mono border border-zinc-200">{sale.tripId}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.tripDate}</td>
                          <td className="px-4 py-3 text-sm font-bold border border-zinc-200">₹{sale.packageValue.toLocaleString()}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">₹{sale.ppMargin.toLocaleString()}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">{sale.noOfPax}</td>
                          <td className="px-4 py-3 text-sm font-bold text-green-600 border border-zinc-200">₹{sale.totalMargin.toLocaleString()}</td>
                          <td className="px-4 py-3 text-sm border border-zinc-200">₹{sale.flight.toLocaleString()}</td>
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
                          <td className="px-4 py-3 border border-zinc-200">
                            <div className="flex items-center gap-2">
                              <button 
                                onClick={() => handleEditSales(sale)}
                                className="p-1.5 hover:bg-orange-50 text-zinc-400 hover:text-orange-600 rounded-lg transition-colors"
                              >
                                <Edit2 className="w-4 h-4" />
                              </button>
                              <button 
                                onClick={() => sale.id && handleDeleteSales(sale.id, sale.guestName)}
                                className="p-1.5 hover:bg-red-50 text-zinc-400 hover:text-red-600 rounded-lg transition-colors"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
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
      </main>

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

      {/* Footer */}
      <footer className="max-w-7xl mx-auto px-4 py-12 border-t border-zinc-200 text-center">
        <p className="text-sm text-zinc-400">© 2026 Nitsa Bravo Points Dashboard. Built with ❤️ for the team.</p>
      </footer>
      </div>
  );
}
