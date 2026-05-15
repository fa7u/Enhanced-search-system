/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React from 'react';
import { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx-js-style';
import { 
  Upload, 
  Search, 
  FileSpreadsheet, 
  X, 
  User, 
  Phone, 
  MessageCircle,
  MapPin, 
  SearchSlash, 
  Clock, 
  Calendar,
  ShieldCheck,
  Infinity,
  BarChart3, 
  ChevronRight,
  Database,
  Download,
  Info,
  RotateCcw,
  Check,
  CheckCircle,
  FileEdit,
  LayoutGrid,
  Circle,
  Wallet,
  Coins,
  History,
  TrendingDown,
  Lock,
  LogOut,
  ShieldAlert,
  Mail,
  Trash2,
  RefreshCw
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { initializeApp } from 'firebase/app';
import { getAuth, signInWithPopup, GoogleAuthProvider, onAuthStateChanged, signOut, User as FirebaseUser, browserPopupRedirectResolver } from 'firebase/auth';
import { 
  getFirestore, 
  doc, 
  getDocFromServer, 
  getDoc, 
  onSnapshot, 
  collection, 
  getDocs, 
  updateDoc, 
  setDoc, 
  deleteDoc, 
  serverTimestamp 
} from 'firebase/firestore';
import firebaseConfig from '@/firebase-applet-config.json';

// Initialize Firebase
const app = initializeApp(firebaseConfig);
export const db = getFirestore(app, firebaseConfig.firestoreDatabaseId);
export const auth = getAuth();

// Provider factory to ensure a fresh instance or correctly configured one
const getGoogleProvider = () => {
  const provider = new GoogleAuthProvider();
  provider.addScope('email');
  provider.addScope('profile');
  provider.setCustomParameters({
    prompt: 'select_account'
  });
  return provider;
};

enum OperationType {
  CREATE = 'create',
  UPDATE = 'update',
  DELETE = 'delete',
  LIST = 'list',
  GET = 'get',
  WRITE = 'write',
}

interface FirestoreErrorInfo {
  error: string;
  operationType: OperationType;
  path: string | null;
  authInfo: {
    userId?: string | null;
    email?: string | null;
    emailVerified?: boolean | null;
    isAnonymous?: boolean | null;
    tenantId?: string | null;
    providerInfo?: {
      providerId?: string | null;
      email?: string | null;
    }[];
  }
}

function handleFirestoreError(error: unknown, operationType: OperationType, path: string | null) {
  const errInfo: FirestoreErrorInfo = {
    error: error instanceof Error ? error.message : String(error),
    authInfo: {
      userId: auth.currentUser?.uid,
      email: auth.currentUser?.email,
      emailVerified: auth.currentUser?.emailVerified,
      isAnonymous: auth.currentUser?.isAnonymous,
      tenantId: auth.currentUser?.tenantId,
      providerInfo: auth.currentUser?.providerData?.map(provider => ({
        providerId: provider.providerId,
        email: provider.email,
      })) || []
    },
    operationType,
    path
  }
  console.error('Firestore Error: ', JSON.stringify(errInfo));
  throw new Error(JSON.stringify(errInfo));
}

interface DataRow {
  [key: string]: string | number | null | undefined;
}

type FilterType = 'all' | 'paid' | 'modified' | 'normal';

export default function App() {
  const [user, setUser] = useState<FirebaseUser | null>(null);
  const [isAuthorized, setIsAuthorized] = useState<boolean | null>(null);
  const [isFrozen, setIsFrozen] = useState(false);
  const [isAdmin, setIsAdmin] = useState(false);
  const [isCheckingAuth, setIsCheckingAuth] = useState(true);
  const [isLoadingAuth, setIsLoadingAuth] = useState(false);
  const [showLogoutConfirm, setShowLogoutConfirm] = useState(false);
  const [showAdminPanel, setShowAdminPanel] = useState(false);
  const [frozenReason, setFrozenReason] = useState<'manual' | 'expired' | null>(null);
  const [whitelist, setWhitelist] = useState<{email: string, addedAt: any, subscriptionExpiry: any, role: 'admin' | 'visitor', isFrozen?: boolean, name?: string}[]>([]);
  const [newEmail, setNewEmail] = useState('');
  const [newName, setNewName] = useState('');
  const [newSubscriptionPeriod, setNewSubscriptionPeriod] = useState<'1m' | '6m' | '1y' | 'lifetime'>('1m');
  const [newRole, setNewRole] = useState<'admin' | 'visitor'>('visitor');
  const [isAddingEmail, setIsAddingEmail] = useState(false);
  const [emailToDelete, setEmailToDelete] = useState<string | null>(null);
  const [isDeleting, setIsDeleting] = useState(false);
  
  const [data, setData] = useState<DataRow[]>([]);
  const [originalData, setOriginalData] = useState<DataRow[]>([]);
  const [headers, setHeaders] = useState<string[]>([]);
  const [searchQuery, setSearchQuery] = useState('');
  const [filterType, setFilterType] = useState<FilterType>('all');
  const [isLoading, setIsLoading] = useState(false);
  const [fileName, setFileName] = useState<string | null>(null);
  const [paidRows, setPaidRows] = useState<Set<number>>(new Set());
  const [modifiedRows, setModifiedRows] = useState<Set<number>>(new Set());
  const [editingCell, setEditingCell] = useState<{ row: number; col: string } | null>(null);
  const [editingValue, setEditingValue] = useState<string>('');
  const [isModified, setIsModified] = useState(false);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const searchResults = useMemo(() => {
    let results = data.map((row, index) => ({ row, index }));
    
    if (searchQuery.trim()) {
      const query = searchQuery.toLowerCase();
      results = results.filter(({ row }) => {
        return Object.values(row).some(value => 
          String(value || '').toLowerCase().includes(query)
        );
      });
    }
    return results;
  }, [data, searchQuery]);

  const filteredItems = useMemo(() => {
    if (filterType === 'all') {
      return searchResults;
    }

    return searchResults.filter(({ index }) => {
      const isPaid = paidRows.has(index);
      const isModified = modifiedRows.has(index);
      
      if (filterType === 'paid') return isPaid;
      if (filterType === 'modified') return isModified;
      return true;
    });
  }, [searchResults, filterType, paidRows, modifiedRows]);

  const filteredData = useMemo(() => filteredItems.map(item => item.row), [filteredItems]);

  // Financial Summary Calculation
  const totals = useMemo(() => {
    const isSearching = searchQuery.trim() !== '' || filterType !== 'all';
    // Use indexed items to avoid repeated findIndex calls
    const currentItems = isSearching ? filteredItems : data.map((row, index) => ({ row, index }));
    
    // Attempt to find columns for Amount and Remaining
    const amountKeywords = ['مبلغ', 'المبلغ', 'قيمة', 'القيمة', 'amount', 'total', 'price', 'السعر'];
    const remainingKeywords = ['متبقي', 'المتبقي', 'باقي', 'الباقي', 'remaining', 'balance', 'left'];

    const amountCol = headers.find(h => amountKeywords.some(k => h.toLowerCase().includes(k)));
    const remainingCol = headers.find(h => remainingKeywords.some(k => h.toLowerCase().includes(k)));

    let totalAmount = 0;
    let totalRemaining = 0;
    let totalSettledFiltered = 0;
    let totalSettledGlobal = 0;
    let totalLiquidation = 0;

    currentItems.forEach(({ row, index }) => {
      const isPaid = paidRows.has(index);
      
      // Check for liquidation keyword "تصفية" in all columns of this row
      const isLiquidation = Object.values(row).some(val => 
        String(val).toLowerCase().includes('تصفية')
      );

      if (amountCol) {
        const val = parseFloat(String(row[amountCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalAmount += val;
      }
      if (remainingCol) {
        const val = parseFloat(String(row[remainingCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) {
          totalRemaining += val;
          
          // Only add to summary if this row is paid
          if (isPaid) {
            totalSettledFiltered += val;
          }
          
          // Add to liquidation total if keyword found
          if (isLiquidation) {
            totalLiquidation += val;
          }
        }
      }
    });

    // Calculate total settled GLOBAL (Entire file) for Dashboard
    paidRows.forEach(idx => {
      const row = data[idx];
      if (row && remainingCol) {
        const val = parseFloat(String(row[remainingCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) totalSettledGlobal += val;
      }
    });

    // Calculate counts for filters based on search results
    const counts = {
      all: searchResults.length,
      paid: searchResults.filter(item => paidRows.has(item.index)).length,
      modified: searchResults.filter(item => modifiedRows.has(item.index)).length
    };

    return { 
      totalAmount, 
      totalRemaining, 
      totalSettledGlobal, 
      totalSettledFiltered, 
      totalLiquidation,
      amountCol, 
      remainingCol,
      counts
    };
  }, [data, searchResults, filteredItems, headers, searchQuery, filterType, paidRows, modifiedRows]);

  // Initialize IndexedDB and load saved data
  useEffect(() => {
    const loadSavedData = async () => {
      if (!user) {
        // Clear local states on logout
        setData([]);
        setHeaders([]);
        setFileName(null);
        setPaidRows(new Set());
        setModifiedRows(new Set());
        setIsModified(false);
        setSearchQuery('');
        return;
      }

      setIsLoading(true);
      
      // Restore search query from localStorage - per user
      const savedQuery = localStorage.getItem(`last_search_query_${user.uid}`);
      if (savedQuery) setSearchQuery(savedQuery);
      
      try {
        const dbRequest = indexedDB.open('FileSearchDB', 1);
        
        dbRequest.onupgradeneeded = (e: any) => {
          const dbView = e.target.result;
          if (!dbView.objectStoreNames.contains('files')) {
            dbView.createObjectStore('files', { keyPath: 'id' });
          }
        };

        dbRequest.onsuccess = (e: any) => {
          const dbView = e.target.result;
          const transaction = dbView.transaction(['files'], 'readonly');
          const store = transaction.objectStore('files');
          const getRequest = store.get(`user_file_${user.uid}`);

          getRequest.onsuccess = () => {
            if (getRequest.result) {
              const res = getRequest.result as { 
              data: DataRow[], 
              originalData?: DataRow[],
              headers: string[], 
              fileName: string, 
              paidIndices?: number[], 
              modifiedIndices?: number[],
              isModified?: boolean 
            };
            const { data: savedData, originalData: savedOriginal, headers: savedHeaders, fileName: savedName, paidIndices, modifiedIndices, isModified: savedModified } = res;
            setData(savedData);
            if (savedOriginal && savedOriginal.length > 0) {
              setOriginalData(savedOriginal);
            } else if (savedData && savedData.length > 0) {
              setOriginalData(JSON.parse(JSON.stringify(savedData)));
            }
            setHeaders(savedHeaders);
            setFileName(savedName);
            if (paidIndices) setPaidRows(new Set(paidIndices));
            if (modifiedIndices) setModifiedRows(new Set(modifiedIndices));
            if (savedModified) setIsModified(savedModified);
            }
            setIsLoading(false);
          };
          
          getRequest.onerror = () => {
            console.error('Failed to get data from store');
            setIsLoading(false);
          };
        };

        dbRequest.onerror = (e) => {
          console.error('DB Open Error:', e);
          setIsLoading(false);
        };
      } catch (error) {
        console.error('IndexedDB Error:', error);
        setIsLoading(false);
      }
    };

    loadSavedData();
  }, [user]);

  // Firebase Auth Observer
  useEffect(() => {
    const findEmail = (u: any) => {
      if (!u) return null;
      console.log("Detecting email for user:", u.uid);
      if (u.email) {
        console.log("Primary email found:", u.email);
        return u.email;
      }
      if (u.providerData && u.providerData.length > 0) {
        for (const p of u.providerData) {
          if (p.email) {
            console.log("Provider email found:", p.email);
            return p.email;
          }
        }
      }
      console.warn("No email found in any provider data for UID:", u.uid);
      return null;
    };

    const unsubscribe = onAuthStateChanged(auth, async (currentUser) => {
      setUser(currentUser);
      if (currentUser) {
        const detectedEmail = findEmail(currentUser);
        console.log("Auth event triggered. Email detected:", detectedEmail);
        await checkAuthorization(detectedEmail || '');
      } else {
        setIsAuthorized(null);
        setIsAdmin(false);
        setIsCheckingAuth(false);
      }
    });
    return () => unsubscribe();
  }, []);

  // Add Real-time listener for current user's status
  useEffect(() => {
    if (!user?.email) return;
    
    const email = user.email.toLowerCase();
    
    // Skip real-time listener for owners
    if (email === 'langmix2@gmail.com' || email === 'lnagmix2@gmail.com' || user.uid === 'acCG3siZciQkWN7jRj5FXwGtDCf2') {
      return;
    }

    const docRef = doc(db, 'whitelist', email);
    
    const unsubscribe = onSnapshot(docRef, (docSnap) => {
      if (docSnap.exists()) {
        const data = docSnap.data();
        const expiry = data.subscriptionExpiry;
        const now = new Date();
        const hasExpired = expiry && expiry.toDate() < now;

        if (data.isFrozen || hasExpired) {
          setIsFrozen(true);
          setFrozenReason(hasExpired ? 'expired' : 'manual');
          clearLocalData();
        } else {
          setIsFrozen(false);
          setFrozenReason(null);
        }
        setIsAuthorized(true);
      } else {
        // Doc removed
        setIsAuthorized(false);
      }
    });

    return () => unsubscribe();
  }, [user]);

  const checkAuthorization = async (rawEmail: string) => {
    setIsCheckingAuth(true);
    setIsAdmin(false); 
    const email = rawEmail?.trim()?.toLowerCase() || '';
    console.log("--- AUTHORIZATION START ---");
    console.log("Target Email:", email || "NO_EMAIL");
    
    if (!email) {
      console.error("Auth Fail: Email is missing");
      setIsAuthorized(false);
      setIsCheckingAuth(false);
      return;
    }

    try {
      const currentUser = auth.currentUser;
      if (!currentUser) {
        console.error("Auth Fail: No authenticated user object found");
        setIsAuthorized(false);
        setIsCheckingAuth(false);
        return;
      }

      // 1. Owner Hardcoded Check (Admin privileges ONLY for owner)
      if (email === 'langmix2@gmail.com' || email === 'lnagmix2@gmail.com' || currentUser.uid === 'acCG3siZciQkWN7jRj5FXwGtDCf2') {
        console.log("Owner admin detected");
        setIsAdmin(true);
        setIsAuthorized(true);
        setIsFrozen(false); // Owner is never frozen
        setIsCheckingAuth(false);
        return;
      }
      
      // 2. Just Whitelist Check (Access only)
      const whitelistRef = doc(db, 'whitelist', email);
      const whitelistSnap = await getDoc(whitelistRef);

      if (whitelistSnap?.exists()) {
        const whitelistData = whitelistSnap.data();
        console.log("Access granted: Email found in whitelist");
        
        const expiry = whitelistData.subscriptionExpiry;
        const now = new Date();
        const hasExpired = expiry && expiry.toDate() < now;

        if (whitelistData.isFrozen || hasExpired) {
          console.log("User is frozen or expired. Clearing local data and blocking access to upload.");
          setIsFrozen(true);
          setFrozenReason(hasExpired ? 'expired' : 'manual');
          clearLocalData(); 
        } else {
          setIsFrozen(false);
          setFrozenReason(null);
        }
        
        setIsAuthorized(true);
      } else {
        console.warn(`Authorization Result: REJECTED (Email ${email} not found in whitelist)`);
        setIsAuthorized(false);
      }
    } catch (globalError: any) {
      console.error('Fatal Authorization Error:', globalError);
      setIsAuthorized(false);
    } finally {
      setIsCheckingAuth(false);
      console.log("--- AUTHORIZATION END ---");
    }
  };

  const handleRefreshAuth = () => {
    if (user?.email) {
      checkAuthorization(user.email);
    }
  };

  const login = async () => {
    if (isLoadingAuth) {
      console.warn("Login already in progress...");
      return;
    }
    setIsLoadingAuth(true);
    console.log("Initiating login sequence...");
    try {
      const provider = getGoogleProvider();
      const result = await signInWithPopup(auth, provider, browserPopupRedirectResolver);
      console.log("Login successful for:", result.user.email);
    } catch (error: any) {
      if (error.code === 'auth/popup-closed-by-user' || error.code === 'auth/cancelled-by-user') {
        console.warn('Login flow cancelled by user');
      } else {
        console.error('Firebase Auth Error:', error);
        // Special handling for the internal assertion error
        if (error.message?.includes('INTERNAL ASSERTION FAILED')) {
          console.error("Detected Firebase Internal state corruption. Advising refresh.");
        }
        alert(`حدث خطأ أثناء تسجيل الدخول: ${error.message}\nالكود: ${error.code}`);
      }
    } finally {
      setIsLoadingAuth(false);
      console.log("Login sequence concluded.");
    }
  };

  const logout = async () => {
    try {
      await signOut(auth);
      setShowLogoutConfirm(false);
    } catch (error) {
      console.error('Logout Error:', error);
    }
  };

  const clearLocalData = () => {
    setData([]);
    setHeaders([]);
    setFileName(null);
    setPaidRows(new Set());
    setModifiedRows(new Set());
    setIsModified(false);
    removeFromDB();
  };

  // Whitelist Management Functions
  const fetchWhitelist = async () => {
    if (!isAdmin) return;
    const path = 'whitelist';
    try {
      const querySnapshot = await getDocs(collection(db, path));
      const list = querySnapshot.docs.map(doc => ({
        ...doc.data(),
        email: doc.id
      })) as {email: string, addedAt: any, subscriptionExpiry: any, role: 'admin' | 'visitor', isFrozen?: boolean, name?: string}[];
      setWhitelist(list);
    } catch (error) {
      console.error('Fetch Whitelist Error:', error);
      handleFirestoreError(error, OperationType.LIST, path);
    }
  };

  const toggleFreezeStatus = async (targetEmail: string, currentStatus: boolean) => {
    if (!isAdmin) return;
    const path = `whitelist/${targetEmail.toLowerCase()}`;
    try {
      await updateDoc(doc(db, 'whitelist', targetEmail.toLowerCase()), {
        isFrozen: !currentStatus
      });
      await fetchWhitelist();
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, path);
    }
  };

  const addToWhitelist = async () => {
    if (!newEmail.trim() || !newName.trim() || !isAdmin) return;
    setIsAddingEmail(true);
    const email = newEmail.trim().toLowerCase();
    const path = `whitelist/${email}`;
    try {
      // Calculate expiry date
      let expiryDate: Date | null = new Date();
      if (newSubscriptionPeriod === '1m') expiryDate.setDate(expiryDate.getDate() + 30);
      else if (newSubscriptionPeriod === '6m') expiryDate.setDate(expiryDate.getDate() + 180);
      else if (newSubscriptionPeriod === '1y') expiryDate.setDate(expiryDate.getDate() + 365);
      else if (newSubscriptionPeriod === 'lifetime') expiryDate = null;

      const docRef = doc(db, 'whitelist', email);
      await setDoc(docRef, {
        email,
        name: newName.trim(),
        addedAt: serverTimestamp(),
        subscriptionExpiry: expiryDate,
        role: 'visitor',
        isFrozen: false
      });
      console.log("Successfully added to whitelist in Firestore:", email);
      setNewEmail('');
      setNewName('');
      setNewSubscriptionPeriod('1m');
      await fetchWhitelist();
      alert(`تمت إضافة ${email} بنجاح إلى قاعدة البيانات`);
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, path);
    } finally {
      setIsAddingEmail(false);
    }
  };

  const updateSubscription = async (targetEmail: string, period: '1m' | '6m' | '1y' | 'lifetime') => {
    if (!isAdmin) return;
    const path = `whitelist/${targetEmail.toLowerCase()}`;
    try {
      const docRef = doc(db, 'whitelist', targetEmail.toLowerCase());
      const docSnap = await getDoc(docRef);
      
      if (!docSnap.exists()) return;
      
      const currentData = docSnap.data();
      let currentExpiry = currentData.subscriptionExpiry?.toDate() || new Date();
      
      // If subscription already expired, start from today
      if (currentExpiry < new Date()) {
        currentExpiry = new Date();
      }

      let newExpiry: Date | null = new Date(currentExpiry);

      if (period === '1m') {
        newExpiry.setDate(newExpiry.getDate() + 30);
      } else if (period === '6m') {
        newExpiry.setDate(newExpiry.getDate() + 180);
      } else if (period === '1y') {
        newExpiry.setDate(newExpiry.getDate() + 365);
      } else if (period === 'lifetime') {
        newExpiry = null;
      }

      await updateDoc(docRef, {
        subscriptionExpiry: newExpiry,
        isFrozen: false 
      });
      
      await fetchWhitelist();
      alert('تم تحديث مدة الاشتراك بنجاح');
    } catch (error) {
      handleFirestoreError(error, OperationType.UPDATE, path);
    }
  };

  const removeFromWhitelist = async (email: string) => {
    if (!isAdmin) {
      alert("ليس لديك صلاحية للقيام بهذا الإجراء");
      return;
    }

    const lowerEmail = email.trim().toLowerCase();
    const path = `whitelist/${lowerEmail}`;
    
    setIsDeleting(true);
    try {
      await deleteDoc(doc(db, 'whitelist', lowerEmail));
      
      // Update local state immediately
      setWhitelist(prev => prev.filter(item => item.email.toLowerCase() !== lowerEmail));
      setEmailToDelete(null);
      
      // Sync from server
      await fetchWhitelist();
    } catch (error) {
      handleFirestoreError(error, OperationType.DELETE, path);
    } finally {
      setIsDeleting(false);
    }
  };

  useEffect(() => {
    if (showAdminPanel && isAdmin) {
      fetchWhitelist();
    }
  }, [showAdminPanel, isAdmin]);

  // Sync search query to localStorage - per user
  useEffect(() => {
    const currentUser = auth.currentUser;
    if (currentUser) {
      if (searchQuery) {
        localStorage.setItem(`last_search_query_${currentUser.uid}`, searchQuery);
      } else {
        localStorage.removeItem(`last_search_query_${currentUser.uid}`);
      }
    }
  }, [searchQuery]);

  // Save to IndexedDB helper
  const saveToDB = (
    fileData: DataRow[], 
    fileHeaders: string[], 
    name: string, 
    paidIndices: number[] = [], 
    modifiedIndices: number[] = [], 
    modified = false, 
    original: DataRow[] | null = null
  ) => {
    const currentUser = auth.currentUser;
    if (!currentUser) return;

    try {
      const dbRequest = indexedDB.open('FileSearchDB', 1);
      dbRequest.onsuccess = (e: any) => {
        const dbView = e.target.result;
        const transaction = dbView.transaction(['files'], 'readwrite');
        const store = transaction.objectStore('files');
        
        // Ensure we always have some original data to save
        const dataToSaveAsOriginal = (original && original.length > 0) 
          ? original 
          : (originalData && originalData.length > 0 ? originalData : fileData);

        store.put({
          id: `user_file_${currentUser.uid}`,
          data: fileData,
          originalData: dataToSaveAsOriginal,
          headers: fileHeaders,
          fileName: name,
          paidIndices: paidIndices,
          modifiedIndices: modifiedIndices,
          isModified: modified,
          updatedAt: new Date().toISOString()
        });
      };
    } catch (e) {
      console.error('Failed to save to DB:', e);
    }
  };

  // Clear IndexedDB helper
  const removeFromDB = () => {
    const user = auth.currentUser;
    if (user) {
      localStorage.removeItem(`last_search_query_${user.uid}`);
    }
    try {
      const dbRequest = indexedDB.open('FileSearchDB', 1);
      dbRequest.onsuccess = (e: any) => {
        const dbView = e.target.result;
        const transaction = dbView.transaction(['files'], 'readwrite');
        const store = transaction.objectStore('files');
        if (user) {
          store.delete(`user_file_${user.uid}`);
        }
      };
    } catch (e) {
      console.error('Failed to delete from DB:', e);
    }
  };

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    processFile(file);
  };

  const togglePaid = (originalIdx: number) => {
    if (!isAuthorized) return;
    setPaidRows(prev => {
      const next = new Set(prev);
      if (next.has(originalIdx)) {
        next.delete(originalIdx);
      } else {
        next.add(originalIdx);
      }
      setIsModified(true);
      // Persist immediately
      saveToDB(data, headers, fileName || '', Array.from(next) as number[], Array.from(modifiedRows) as number[], true, originalData);
      return next;
    });
  };

  const updateNote = (originalIdx: number, header: string, newValue: string) => {
    if (!isAuthorized) return;
    
    // Use functional updates to ensure we have the latest state
    setData(prevData => {
      const nextData = [...prevData];
      nextData[originalIdx] = { ...nextData[originalIdx], [header]: newValue };
      
      // Track modification
      setModifiedRows(prevModified => {
        const nextModified = new Set(prevModified);
        nextModified.add(originalIdx);
        
        // Persist to DB after both states are updated
        // Note: we use values derived from the closure of the functional update to be safe
        saveToDB(
          nextData, 
          headers, 
          fileName || '', 
          Array.from(paidRows) as number[], 
          Array.from(nextModified) as number[], 
          true, 
          originalData
        );
        
        return nextModified;
      });
      
      return nextData;
    });
    
    setIsModified(true);
  };

  const undoRowChanges = (originalIdx: number) => {
    if (originalData[originalIdx]) {
      // 1. Restore from original
      const nextData = [...data];
      nextData[originalIdx] = { ...originalData[originalIdx] };
      setData(nextData);
      
      // 2. Remove from modified
      const newModifiedRows = new Set(modifiedRows);
      newModifiedRows.delete(originalIdx);
      setModifiedRows(newModifiedRows);
      
      const stillModified = newModifiedRows.size > 0 || paidRows.size > 0;
      setIsModified(stillModified);

      // 3. Persist back
      saveToDB(
        nextData, 
        headers, 
        fileName || '', 
        Array.from(paidRows) as number[], 
        Array.from(newModifiedRows) as number[], 
        stillModified, 
        originalData
      );
    }
  };

  const processFile = (file: File) => {
    setIsLoading(true);
    setFileName(file.name);
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const wb = XLSX.read(bstr, { type: 'binary' });
        const wsname = wb.SheetNames[0];
        const ws = wb.Sheets[wsname];
        const jsonData = XLSX.utils.sheet_to_json<DataRow>(ws);
        
        if (jsonData.length > 0) {
          const newHeaders = Object.keys(jsonData[0]);
          setHeaders(newHeaders);
          setData(jsonData);
          setOriginalData(JSON.parse(JSON.stringify(jsonData)));
          setPaidRows(new Set());
          setModifiedRows(new Set());
          setIsModified(false);
          saveToDB(jsonData, newHeaders, file.name, [], [], false, JSON.parse(JSON.stringify(jsonData)));
        } else {
          alert('الملف فارغ أو لا يحتوي على بيانات صحيحة.');
        }
      } catch (error) {
        console.error('Error parsing file:', error);
        alert('حدث خطأ أثناء قراءة الملف. تأكد من أنه ملف إكسل صالح.');
      } finally {
        setIsLoading(false);
      }
    };
    reader.readAsBinaryString(file);
  };

  const clearFile = () => {
    setData([]);
    setHeaders([]);
    setSearchQuery('');
    setFilterType('all');
    setFileName(null);
    setPaidRows(new Set());
    setModifiedRows(new Set());
    setIsModified(false);
    removeFromDB();
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const getHeaderTooltip = (header: string) => {
    const h = header.toLowerCase();
    if (h.includes('اسم') || h.includes('name')) return 'الاسم الكامل للعميل أو المستأجر';
    if (h.includes('عقار') || h.includes('وحدة') || h.includes('address')) return 'وصف العقار أو الوحدة السكنية';
    if (h.includes('مبلغ') || h.includes('إجمالي') || h.includes('amount')) return 'إجمالي القيمة المالية المستحقة';
    if (h.includes('متبقي') || h.includes('باقي') || h.includes('remaining')) return 'المبلغ المتبقي الذي لم يتم سداده بعد';
    if (h.includes('دفع') || h.includes('مدفوع') || h.includes('payment')) return 'المبالغ التي تم تحصيلها بالفعل';
    if (h.includes('تاريخ') || h.includes('date')) return 'تاريخ العملية أو موعد الاستحقاق';
    if (h.includes('ملاحظات') || h.includes('notes')) return 'أي تفاصيل إضافية أو تنبيهات متعلقة بالسجل';
    if (h.includes('هاتف') || h.includes('جوال') || h.includes('phone')) return 'رقم التواصل الخاص بالسجل';
    return 'بيانات إضافية متعلقة بهذا العمود';
  };

  const exportToExcel = (forceAll = false) => {
    if (!isAuthorized) return;
    // Deciding formatting locale (ar-EG for Gregorian)
    const locale = 'ar-EG';
    
    // 1. Decide which data set to export
    const hasActiveFilter = searchQuery.trim() !== '' || filterType !== 'all';
    const isExportAll = forceAll || !hasActiveFilter;
    
    // Use indexed items for efficiency and correctness
    const itemsToExport = isExportAll 
      ? data.map((row, index) => ({ row, index })) 
      : filteredItems;
    
    if (itemsToExport.length === 0) return;

    const dataToExport = itemsToExport.map(i => i.row);

    // 1.1 Calculate specific totals for the exported data
    let exportAmount = 0;
    let exportRemaining = 0;
    let exportSettled = 0;
    let exportLiquidation = 0;

    itemsToExport.forEach(({ row, index }) => {
      const isPaid = paidRows.has(index);
      const isLiquidation = Object.values(row).some(v => String(v).includes('تصفية'));

      if (totals.amountCol) {
        const val = parseFloat(String(row[totals.amountCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) exportAmount += val;
      }
      if (totals.remainingCol) {
        const val = parseFloat(String(row[totals.remainingCol]).replace(/[^0-9.-]+/g, ''));
        if (!isNaN(val)) {
          exportRemaining += val;
          // Only add to summary if this row is marked as settled/paid
          if (isPaid) {
            exportSettled += val;
          }
          // Liquidation total
          if (isLiquidation) {
            exportLiquidation += val;
          }
        }
      }
    });

    // 2. Create worksheet from data
    const worksheet = XLSX.utils.json_to_sheet(dataToExport);

    // 3. Define Styles
    const headerStyle = {
      fill: { fgColor: { rgb: "4F46E5" } }, // Indigo 600
      font: { color: { rgb: "FFFFFF" }, bold: true, sz: 12 },
      alignment: { horizontal: "center", vertical: "center", wrapText: true },
      border: {
        top: { style: "thin", color: { rgb: "E2E8F0" } },
        bottom: { style: "thin", color: { rgb: "E2E8F0" } },
        left: { style: "thin", color: { rgb: "E2E8F0" } },
        right: { style: "thin", color: { rgb: "E2E8F0" } }
      }
    };

    const dataStyle = {
      font: { sz: 11, color: { rgb: "334155" } },
      alignment: { horizontal: "right", vertical: "center", wrapText: true },
      border: {
        bottom: { style: "thin", color: { rgb: "F1F5F9" } }
      }
    };

    const summaryStyle = {
      fill: { fgColor: { rgb: "F1F5F9" } }, 
      font: { bold: true, sz: 12, color: { rgb: "1E293B" } },
      alignment: { horizontal: "right", vertical: "center" },
      border: {
        top: { style: "medium", color: { rgb: "4F46E5" } }
      }
    };

    const highlightStyle = (colorCode: string) => ({
      ...summaryStyle,
      font: { ...summaryStyle.font, color: { rgb: colorCode } }
    });

    // 4. Apply Styles to Header
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_col(C) + "1";
      if (worksheet[address]) {
        worksheet[address].s = headerStyle;
      }
    }

    // 5. Apply Styles to Data Rows
    for (let R = range.s.r + 1; R <= range.e.r; ++R) {
      const { row: rowData, index: originalIdx } = itemsToExport[R - 1];
      const isPaid = paidRows.has(originalIdx);
      const isModifiedRow = modifiedRows.has(originalIdx);
      
      let rowBgColor = "FFFFFF";
      if (isPaid) {
        rowBgColor = "D1FAE5"; // Emerald 100
      } else if (isModifiedRow) {
        // Check if notes column contains "تصفية"
        const rowString = JSON.stringify(rowData);
        if (rowString.includes('تصفية')) {
          rowBgColor = "FECACA"; // Red 200 for Tasfya
        } else {
          rowBgColor = "DBEAFE"; // Blue 100 for general edit
        }
      }

      for (let C = range.s.c; C <= range.e.c; ++C) {
        const address = XLSX.utils.encode_cell({ r: R, c: C });
        if (worksheet[address]) {
          worksheet[address].s = {
            ...dataStyle,
            fill: { fgColor: { rgb: rowBgColor } }
          };
        }
      }
    }

    // 6. Add Summary Row
    const summaryRowData = {};
    const settledColIndex = headers.indexOf(totals.remainingCol || '') + 1;
    const hasNextCol = settledColIndex < headers.length;
    const nextColHeader = hasNextCol ? headers[settledColIndex] : null;
    
    const liquidationColIndex = settledColIndex + 1;
    const hasNextCol2 = liquidationColIndex < headers.length;
    const nextColHeader2 = hasNextCol2 ? headers[liquidationColIndex] : null;

    const settledLabel = isExportAll ? "المسدد الإجمالي" : "المسدد للنتائج الحالية";

    headers.forEach((h) => {
      if (amountCol) {
        summaryRowData[h] = `إجمالي المبلغ: ${exportAmount.toLocaleString(locale)} ر.س`;
      } else if (h === totals.remainingCol) {
        summaryRowData[h] = `إجمالي المتبقي: ${exportRemaining.toLocaleString(locale)} ر.س`;
        // If we can't use next column, append it here
        if (!nextColHeader) {
          if (exportSettled > 0) summaryRowData[h] += ` | ${settledLabel}: ${exportSettled.toLocaleString(locale)} ر.س`;
          if (exportLiquidation > 0) summaryRowData[h] += ` | إجمالي التصفية: ${exportLiquidation.toLocaleString(locale)} ر.س`;
        }
      } else if (nextColHeader && h === nextColHeader) {
        if (exportSettled > 0) {
          summaryRowData[h] = `${settledLabel}: ${exportSettled.toLocaleString(locale)} ر.س`;
          if (!nextColHeader2 && exportLiquidation > 0) {
            summaryRowData[h] += ` | إجمالي التصفية: ${exportLiquidation.toLocaleString(locale)} ر.س`;
          }
        }
      } else if (nextColHeader2 && h === nextColHeader2 && exportLiquidation > 0) {
        summaryRowData[h] = `إجمالي التصفية: ${exportLiquidation.toLocaleString(locale)} ر.س`;
      } else if (h === headers[0]) {
        summaryRowData[h] = '--- الملخص الإجمالي ---';
      } else {
        summaryRowData[h] = '';
      }
    });

    XLSX.utils.sheet_add_json(worksheet, [summaryRowData], {
      skipHeader: true,
      origin: -1
    });

    // Style the summary row
    const lastRowIndex = range.e.r + 2; 
    for (let C = range.s.c; C <= range.e.c; ++C) {
      const address = XLSX.utils.encode_cell({ r: lastRowIndex - 1, c: C });
      if (worksheet[address]) {
        const header = headers[C];
        if (header === totals.amountCol) {
          worksheet[address].s = highlightStyle("4F46E5"); 
        } else if (header === totals.remainingCol) {
          worksheet[address].s = highlightStyle("EA580C"); 
        } else if (nextColHeader && header === nextColHeader && exportSettled > 0) {
          worksheet[address].s = highlightStyle("10B981"); // Emerald 500 for settled
        } else if (nextColHeader2 && header === nextColHeader2 && exportLiquidation > 0) {
          worksheet[address].s = highlightStyle("EF4444"); // Red 500 for liquidation
        } else {
          worksheet[address].s = summaryStyle;
        }
      }
    }

    // 7. Config widths & RTL
    const wscols = headers.map(h => {
      const hLower = h.toLowerCase();
      let w = 18;
      if (hLower.includes('ملاحظات')) w = 50;
      else if (hLower.includes('اسم') || hLower.includes('عقار')) w = 35;
      else if (hLower.includes('دفع') || hLower.includes('تاريخ')) w = 25;
      
      // Ensure summary row text fits
      const cellValue = summaryRowData[h] || '';
      if (cellValue.length > w) {
        w = Math.min(cellValue.length + 5, 80); // Max 80 width
      }
      
      return { wch: w };
    });
    worksheet['!cols'] = wscols;
    worksheet['!views'] = [{ RTL: true }];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "البيانات المعدلة");
    const dateStr = new Date().toISOString().split('T')[0];
    const fileNameSafe = (fileName || 'بيانات').replace(/[\\/:*?"<>|]/g, '_');
    const suffix = isExportAll ? 'الكل_معدل' : 'نتائج_البحث';
    XLSX.writeFile(workbook, `سجلات_${fileNameSafe}_${suffix}_${dateStr}.xlsx`);
  };

  if (isCheckingAuth) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center font-sans" dir="rtl">
        <div className="w-16 h-16 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mb-6"></div>
        <p className="text-slate-600 font-bold animate-pulse">جاري التحقق من الهوية...</p>
      </div>
    );
  }

  if (!user) {
    return (
      <div className="min-h-screen bg-slate-50 flex flex-col items-center justify-center font-sans p-6" dir="rtl">
        <motion.div 
          initial={{ opacity: 0, y: 20 }}
          animate={{ opacity: 1, y: 0 }}
          className="bg-white p-10 rounded-[2.5rem] shadow-2xl border border-slate-200 max-w-md w-full text-center"
        >
          <div className="w-20 h-20 bg-indigo-600 rounded-3xl flex items-center justify-center text-white mx-auto mb-8 shadow-xl shadow-indigo-100">
            <Lock size={40} />
          </div>
          <h1 className="text-3xl font-black text-slate-800 mb-4">نظام البحث الخاص</h1>
          <p className="text-slate-500 mb-8 leading-relaxed">
            هذا النظام مخصص للأشخاص المصرح لهم فقط. يرجى تسجيل الدخول للوصول إلى بياناتك.
          </p>

          <div className="bg-slate-50 rounded-2xl p-6 mb-8 border border-slate-100/80 group transition-all hover:bg-slate-100/50 duration-500 text-center">
            <p className="text-[10px] text-slate-400 mb-4 font-bold uppercase tracking-widest">للتواصل مع المالك للتفعيل</p>
            <div className="flex flex-col gap-3">
              <a 
                href="mailto:fahussein79@gmail.com" 
                className="flex items-center justify-center gap-3 text-indigo-600 font-bold hover:underline break-all bg-white p-3 rounded-xl shadow-sm border border-slate-50 text-sm"
              >
                <Mail size={16} />
                <span className="font-mono">fahussein79@gmail.com</span>
              </a>
              <a 
                href="https://wa.me/966550665495"
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center justify-center gap-3 text-emerald-600 font-bold hover:underline bg-white p-3 rounded-xl shadow-sm border border-slate-50 text-sm"
              >
                <MessageCircle size={16} />
                <span className="font-mono">0550665495</span>
              </a>
            </div>
          </div>

          <button 
            onClick={login}
            disabled={isLoadingAuth}
            className="w-full py-4 bg-indigo-600 text-white rounded-2xl font-bold flex items-center justify-center gap-3 hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 active:scale-[0.98] disabled:opacity-70 disabled:cursor-not-allowed"
          >
            {isLoadingAuth ? (
              <div className="w-6 h-6 border-2 border-white border-t-transparent rounded-full animate-spin"></div>
            ) : (
              <>
                <img src="https://www.gstatic.com/firebasejs/ui/2.0.0/images/auth/google.svg" alt="Google" className="w-6 h-6 bg-white rounded-full p-1" />
                تسجيل الدخول باستخدام جوجل
              </>
            )}
          </button>
        </motion.div>
      </div>
    );
  }

  if (isAuthorized === false) {
    return (
      <div className="min-h-screen bg-slate-50 flex items-center justify-center p-6" dir="rtl">
        <motion.div 
          initial={{ opacity: 0, scale: 0.95 }}
          animate={{ opacity: 1, scale: 1 }}
          className="max-w-md w-full bg-white rounded-3xl shadow-xl shadow-slate-200/50 p-10 text-center"
        >
          <div className="w-20 h-20 bg-indigo-50 text-indigo-600 rounded-3xl flex items-center justify-center mx-auto mb-8 ring-8 ring-indigo-50/50">
            <Lock size={40} />
          </div>
          
          <h1 className="text-2xl font-bold text-slate-800 mb-4 font-sans tracking-tight">الدخول غير مصرح به</h1>
          
          <p className="text-slate-500 text-sm mb-8 leading-relaxed px-4">
            عذراً، هذا البريد الإلكتروني غير مضاف في قائمة المصرح لهم بالدخول.
            يرجى التواصل مع مسؤول النظام لطلب صلاحية الوصول.
          </p>

          <div className="bg-slate-50 rounded-2xl p-6 mb-8 border border-slate-100/80 group transition-all hover:bg-slate-100/50 duration-500 text-center">
            <p className="text-[10px] text-slate-400 mb-4 font-bold uppercase tracking-widest">للتواصل مع المالك للتفعيل</p>
            <div className="space-y-4">
              <a 
                href="mailto:fahussein79@gmail.com" 
                className="flex items-center justify-center gap-3 text-indigo-600 font-bold hover:underline break-all bg-white p-3 rounded-xl shadow-sm border border-slate-50"
              >
                <Mail size={18} />
                <span className="text-base font-mono">fahussein79@gmail.com</span>
              </a>
              <a 
                href="https://wa.me/966550665495"
                target="_blank"
                rel="noopener noreferrer"
                className="flex items-center justify-center gap-3 text-emerald-600 font-bold hover:underline bg-white p-3 rounded-xl shadow-sm border border-slate-50"
              >
                <MessageCircle size={18} />
                <span className="text-base font-mono">0550665495</span>
              </a>
            </div>
          </div>

          <div className="flex flex-col gap-4">
            <button 
              onClick={login}
              className="w-full h-14 bg-indigo-600 text-white rounded-2xl font-bold hover:bg-indigo-700 transition-all flex items-center justify-center gap-2 shadow-lg shadow-indigo-100"
            >
              <User size={20} />
              تبديل الحساب
            </button>

            <button 
              onClick={() => auth.signOut()}
              className="w-full h-14 bg-white text-red-500 border border-red-100 rounded-2xl font-bold hover:bg-red-50 transition-all flex items-center justify-center gap-2"
            >
              <LogOut size={20} />
              تسجيل الخروج
            </button>
          </div>
          
          <div className="mt-8 pt-8 border-t border-slate-50">
            <p className="text-[10px] text-slate-300 font-mono uppercase tracking-tighter uppercase">
              Current ID: {user?.uid?.slice(0, 8)}...
            </p>
          </div>
        </motion.div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-50 flex flex-col font-sans" dir="rtl">
      {/* Header Banner for Frozen Accounts */}
      {isFrozen && (
        <div className="bg-red-600 text-white px-4 py-3 text-center flex items-center justify-center gap-4 animate-in slide-in-from-top duration-500 sticky top-0 z-[60]">
          <div className="flex items-center gap-2 font-black text-sm">
            <ShieldAlert size={20} className="animate-pulse" />
            <span>{frozenReason === 'expired' ? 'تنبيه: انتهى اشتراكك!' : 'تنبيه: حسابك مجمد حالياً!'}</span>
          </div>
          <div className="h-4 w-[1px] bg-white/30 hidden sm:block"></div>
          <p className="text-xs font-bold hidden sm:block">
            {frozenReason === 'expired' 
              ? 'يرجى تجديد الاشتراك لاستعادة كامل الصلاحيات والوصول لملفاتك' 
              : 'يرجى التواصل مع المالك لتفعيل الحساب واستعادة كامل الصلاحيات'}
          </p>
          <div className="flex items-center gap-4 mr-auto sm:mr-0">
            <a href="mailto:fahussein79@gmail.com" className="flex items-center gap-1.5 bg-white/20 hover:bg-white/30 px-3 py-1.5 rounded-lg transition-all text-[10px] font-black border border-white/10 uppercase tracking-tighter">
              <Mail size={12} />
              الايميل
            </a>
            <a href="https://wa.me/966550665495" target="_blank" rel="noopener noreferrer" className="flex items-center gap-1.5 bg-white/20 hover:bg-white/30 px-3 py-1.5 rounded-lg transition-all text-[10px] font-black border border-white/10 uppercase tracking-tighter">
              <MessageCircle size={12} />
              واتساب
            </a>
          </div>
        </div>
      )}

      {/* Top Navigation Bar */}
      <nav className="h-16 bg-white border-b border-slate-200 px-6 md:px-8 flex items-center justify-between shadow-sm shrink-0 sticky top-0 z-50">
        <div className="flex items-center gap-3">
          <span className="text-xl font-bold text-slate-800 hidden sm:block">نظام البحث الذكي</span>
        </div>
        
        <div className="flex items-center gap-4">
          <AnimatePresence>
            {user && (
              <div key="user-info" className="text-left flex flex-col items-end">
                <p className="text-[10px] uppercase font-bold text-slate-400 leading-tight">المستخدم</p>
                <p className="text-sm font-medium text-slate-700 max-w-[150px] truncate">{user.displayName || user.email}</p>
              </div>
            )}
            {fileName && (
              <motion.div 
                key="file-info"
                initial={{ opacity: 0, x: -10 }}
                animate={{ opacity: 1, x: 0 }}
                className="text-left flex flex-col items-end"
              >
                <p className="text-[10px] uppercase font-bold text-slate-400 leading-tight">الملف النشط</p>
                <p className="text-sm font-medium text-slate-700 max-w-[150px] truncate">{fileName}</p>
              </motion.div>
            )}
          </AnimatePresence>
          
          {isAdmin && (
            <button 
              onClick={() => setShowAdminPanel(true)}
              className="w-10 h-10 rounded-xl bg-indigo-600 text-white flex items-center justify-center hover:bg-indigo-700 transition-all shadow-lg shadow-indigo-100 relative group"
              title="إدارة المستخدمين"
            >
              <Lock size={18} />
              <div className="absolute -top-1 -right-1 w-3 h-3 bg-red-500 border-2 border-white rounded-full"></div>
            </button>
          )}

          <button 
            onClick={() => setShowLogoutConfirm(true)}
            className="w-10 h-10 rounded-xl bg-slate-50 border border-slate-100 flex items-center justify-center text-slate-400 hover:text-red-500 hover:bg-red-50 transition-all"
            title="تسجيل الخروج"
          >
            <LogOut size={20} />
          </button>
        </div>
      </nav>

      {/* Logout Confirmation Modal */}
      <AnimatePresence>
        {showLogoutConfirm && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowLogoutConfirm(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, scale: 0.9, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.9, y: 20 }}
              className="relative bg-white rounded-[2rem] p-8 shadow-2xl border border-slate-200 max-w-sm w-full text-center"
            >
              <div className="w-16 h-16 bg-red-50 text-red-500 rounded-2xl flex items-center justify-center mx-auto mb-6">
                <LogOut size={32} />
              </div>
              <h3 className="text-xl font-bold text-slate-800 mb-2">تأكيد تسجيل الخروج</h3>
              <p className="text-slate-500 mb-8 font-medium">هل أنت متأكد أنك تريد تسجيل الخروج من النظام؟</p>
              <div className="flex gap-3">
                <button 
                  onClick={logout}
                  className="flex-1 py-3 bg-red-500 text-white rounded-xl font-bold hover:bg-red-600 transition-all shadow-lg shadow-red-100"
                >
                  نعم، خروج
                </button>
                <button 
                  onClick={() => setShowLogoutConfirm(false)}
                  className="flex-1 py-3 bg-slate-100 text-slate-600 rounded-xl font-bold hover:bg-slate-200 transition-all"
                >
                  إلغاء
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Admin Whitelist Panel Modal */}
      <AnimatePresence>
        {showAdminPanel && isAdmin && (
          <div className="fixed inset-0 z-[100] flex items-center justify-center p-4">
            <motion.div 
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowAdminPanel(false)}
              className="absolute inset-0 bg-slate-900/60 backdrop-blur-sm"
            />
            <motion.div 
              initial={{ opacity: 0, x: 100 }}
              animate={{ opacity: 1, x: 0 }}
              exit={{ opacity: 0, x: 100 }}
              className="relative bg-white h-full max-w-md w-full ml-auto shadow-2xl flex flex-col"
            >
              <div className="p-6 border-b border-slate-100 flex items-center justify-between">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 bg-indigo-600 rounded-xl flex items-center justify-center text-white">
                    <User size={20} />
                  </div>
                  <div>
                    <h3 className="font-bold text-slate-800">إدارة المصرح لهم</h3>
                    <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">لوحة التحكم الإدارية</p>
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <button onClick={() => setShowAdminPanel(false)} className="p-2 hover:bg-slate-100 rounded-xl text-slate-400">
                    <X size={20} />
                  </button>
                </div>
              </div>

              <div className="p-6 border-b border-slate-100 bg-slate-50/50">
                <p className="text-xs font-bold text-slate-500 mb-3">إضافة مستخدم جديد:</p>
                <div className="flex flex-col gap-3">
                  <div className="grid grid-cols-2 gap-2">
                    <div className="relative">
                      <input 
                        type="text" 
                        placeholder="اسم صاحب الحساب"
                        value={newName}
                        onChange={(e) => setNewName(e.target.value)}
                        className="w-full h-11 pl-4 pr-10 rounded-xl border border-slate-200 bg-white text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                      />
                      <User size={16} className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400" />
                    </div>
                    <div className="relative">
                      <select
                        value={newSubscriptionPeriod}
                        onChange={(e) => setNewSubscriptionPeriod(e.target.value as any)}
                        className="w-full h-11 px-4 rounded-xl border border-slate-200 bg-white text-sm focus:ring-2 focus:ring-indigo-500 outline-none appearance-none"
                      >
                        <option value="1m">شهر واحد</option>
                        <option value="6m">6 أشهر</option>
                        <option value="1y">سنة كاملة</option>
                        <option value="lifetime">مدى الحياة</option>
                      </select>
                      <Calendar size={16} className="absolute left-4 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                    </div>
                  </div>
                  <div className="flex gap-2">
                    <div className="flex-1 relative">
                      <input 
                        type="email" 
                        placeholder="example@gmail.com"
                        value={newEmail}
                        onChange={(e) => setNewEmail(e.target.value)}
                        className="w-full h-11 pl-4 pr-10 rounded-xl border border-slate-200 bg-white text-sm focus:ring-2 focus:ring-indigo-500 outline-none"
                      />
                      <Mail size={16} className="absolute right-4 top-1/2 -translate-y-1/2 text-slate-400" />
                    </div>
                    <button 
                      onClick={addToWhitelist}
                      disabled={isAddingEmail || !newEmail.trim() || !newName.trim()}
                      className="px-6 bg-indigo-600 text-white rounded-xl font-bold text-sm hover:bg-indigo-700 transition-all disabled:opacity-50 flex items-center justify-center min-w-[80px]"
                    >
                      {isAddingEmail ? <div className="w-4 h-4 border-2 border-white/30 border-t-white rounded-full animate-spin" /> : 'إضافة'}
                    </button>
                  </div>
                </div>
              </div>

              <div className="flex-1 overflow-auto p-6 space-y-4 custom-scrollbar">
                <p className="text-xs font-bold text-slate-500 mb-2">القائمة الحالية ({whitelist.length}):</p>
                {whitelist.length === 0 ? (
                  <div key="no-whitelist" className="text-center py-10 opacity-40">
                    <Mail size={40} className="mx-auto mb-2" />
                    <p className="text-xs font-bold">لا يوجد مستخدمين مضافين حالياً</p>
                  </div>
                ) : (
                  whitelist.map((item) => (
                    <div key={item.email} className="p-4 bg-white border border-slate-100 rounded-[2rem] group hover:border-indigo-100 transition-all shadow-sm">
                      <div className="flex flex-col gap-4">
                        {/* Top: User Info */}
                        <div className="flex items-center gap-3">
                          <div className="w-10 h-10 bg-indigo-50 text-indigo-600 rounded-2xl flex items-center justify-center flex-shrink-0 shadow-inner">
                            <Mail size={18} />
                          </div>
                          <div className="flex-1 min-w-0">
                            <div className="flex items-center flex-wrap gap-2">
                              <p className="text-sm font-black text-slate-800 truncate">{item.name || 'مستخدم بدون اسم'}</p>
                              {item.isFrozen && (
                                <span className="bg-red-50 text-red-600 text-[9px] px-2 py-0.5 rounded-full font-black border border-red-100">مجمد</span>
                              )}
                            </div>
                            <p className="text-[10px] text-slate-500 font-bold truncate break-all">{item.email}</p>
                            <div className="flex flex-wrap gap-x-3 gap-y-1 mt-1">
                              <p className="text-[10px] text-slate-400 font-bold">
                                مضاف: {item.addedAt?.seconds ? new Date(item.addedAt.seconds * 1000).toLocaleDateString('ar-EG') : 'غير معروف'}
                              </p>
                              {item.subscriptionExpiry ? (
                                <p className={`text-[10px] font-black ${new Date(item.subscriptionExpiry.seconds * 1000) < new Date() ? 'text-red-500' : 'text-emerald-600'}`}>
                                  {new Date(item.subscriptionExpiry.seconds * 1000) < new Date() ? 'منتهي: ' : 'ينتهي: '}
                                  {new Date(item.subscriptionExpiry.seconds * 1000).toLocaleDateString('ar-EG')}
                                </p>
                              ) : (
                                <p className="text-[10px] text-indigo-600 font-black">
                                  اشتراك مدى الحياة
                                </p>
                              )}
                            </div>
                          </div>
                        </div>

                        {/* Bottom: Actions */}
                        <div className="space-y-2 pt-3 border-t border-slate-50">
                          <div className="flex items-center gap-1">
                            <span className="text-[9px] font-bold text-slate-400 ml-1">تجديد:</span>
                            <button onClick={() => updateSubscription(item.email, '1m')} className="px-2 py-1 bg-indigo-50 text-indigo-600 rounded-lg text-[9px] font-black hover:bg-indigo-100 transition-all border border-indigo-100">شهر</button>
                            <button onClick={() => updateSubscription(item.email, '6m')} className="px-2 py-1 bg-indigo-50 text-indigo-600 rounded-lg text-[9px] font-black hover:bg-indigo-100 transition-all border border-indigo-100">6 أشهر</button>
                            <button onClick={() => updateSubscription(item.email, '1y')} className="px-2 py-1 bg-indigo-50 text-indigo-600 rounded-lg text-[9px] font-black hover:bg-indigo-100 transition-all border border-indigo-100">سنة</button>
                            <button onClick={() => updateSubscription(item.email, 'lifetime')} className="px-2 py-1 bg-slate-100 text-slate-600 rounded-lg text-[9px] font-black hover:bg-slate-200 transition-all border border-slate-200">دائم</button>
                          </div>
                          <div className="flex items-center gap-2">
                            <button 
                              onClick={() => toggleFreezeStatus(item.email, !!item.isFrozen)}
                              className={`flex-1 flex items-center justify-center gap-2 py-2 rounded-xl transition-all font-black text-[10px] border shadow-sm active:scale-95 ${
                                item.isFrozen 
                                  ? 'text-emerald-700 bg-emerald-50 border-emerald-100 hover:bg-emerald-100' 
                                  : 'text-orange-700 bg-orange-50 border-orange-100 hover:bg-orange-100'
                              }`}
                            >
                              {item.isFrozen ? <CheckCircle size={12} /> : <ShieldAlert size={12} />}
                              <span>{item.isFrozen ? 'تفعيل يدوي' : 'تجميد يدوي'}</span>
                            </button>
                            
                            <button 
                              type="button"
                              onClick={(e) => {
                                e.preventDefault();
                                e.stopPropagation();
                                setEmailToDelete(item.email);
                              }}
                              className="px-3 flex items-center justify-center gap-2 py-2 rounded-xl text-red-700 bg-red-50 border border-red-100 hover:bg-red-100 active:scale-95 transition-all font-black text-[10px] shadow-sm"
                            >
                              <Trash2 size={12} />
                            </button>
                          </div>
                        </div>
                      </div>
                    </div>
                  ))
                )}
              </div>
              
              <div className="p-6 bg-slate-50 text-[10px] text-slate-400 font-bold text-center border-t border-slate-100">
                ملاحظة: الصلاحيات تمنح فور الإضافة.
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>

      {/* Delete Confirmation Modal */}
      <AnimatePresence>
        {emailToDelete && (
          <motion.div 
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            className="fixed inset-0 z-[100] flex items-center justify-center p-4 bg-slate-900/60 backdrop-blur-sm"
            onClick={() => !isDeleting && setEmailToDelete(null)}
          >
            <motion.div 
              initial={{ scale: 0.9, opacity: 0, y: 20 }}
              animate={{ scale: 1, opacity: 1, y: 0 }}
              exit={{ scale: 0.9, opacity: 0, y: 20 }}
              className="bg-white rounded-[2rem] p-8 max-w-sm w-full shadow-2xl border border-slate-100 text-center relative overflow-hidden"
              onClick={(e) => e.stopPropagation()}
            >
              {/* Decorative Background */}
              <div className="absolute top-0 left-0 w-full h-2 bg-red-500"></div>
              
              <div className="w-20 h-20 bg-red-50 text-red-600 rounded-3xl flex items-center justify-center mx-auto mb-6 shadow-inner ring-4 ring-white">
                <Trash2 size={40} />
              </div>
              
              <h3 className="text-2xl font-black text-slate-800 mb-3 tracking-tight">تأكيد الحذف</h3>
              <p className="text-slate-500 text-sm mb-8 leading-relaxed px-4">
                هل أنت متأكد من حذف الحساب <br/>
                <span className="font-bold text-red-600 text-base block mt-1 dir-ltr">({emailToDelete})</span>
                <span className="block mt-2 text-xs opacity-70">سيتم منعه من الدخول للنظام نهائياً.</span>
              </p>
              
              <div className="flex flex-col gap-3">
                <button
                  disabled={isDeleting}
                  onClick={() => removeFromWhitelist(emailToDelete)}
                  className="w-full py-4 bg-red-600 hover:bg-red-700 text-white rounded-2xl font-black text-sm shadow-xl shadow-red-200 transition-all flex items-center justify-center gap-3 active:scale-[0.98] disabled:opacity-50 disabled:pointer-events-none"
                >
                  {isDeleting ? (
                    <>
                      <RefreshCw size={18} className="animate-spin" />
                      <span>جاري الحذف...</span>
                    </>
                  ) : (
                    <>
                      <Trash2 size={18} />
                      <span>نعم، أحذف</span>
                    </>
                  )}
                </button>
                <button
                  disabled={isDeleting}
                  onClick={() => setEmailToDelete(null)}
                  className="w-full py-4 bg-slate-100 hover:bg-slate-200 text-slate-600 rounded-2xl font-bold text-sm transition-all active:scale-[0.98] disabled:opacity-50"
                >
                  تراجع
                </button>
              </div>
            </motion.div>
          </motion.div>
        )}
      </AnimatePresence>

      <main className="flex-1 p-4 md:p-8 grid grid-cols-12 gap-8 max-w-[1400px] mx-auto w-full">
        {/* Sidebar / Action Panel */}
        <aside className="col-span-12 lg:col-span-4 flex flex-col gap-6">
          {/* Upload Area */}
          <section className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm">
            <h2 className="text-lg font-bold text-slate-800 mb-4 flex items-center gap-2">
              <Upload size={18} className="text-indigo-600" />
              رفع الملفات
            </h2>
            <div 
              onClick={() => {
                if (isFrozen) {
                  const msg = frozenReason === 'expired' 
                    ? 'عذراً، انتهى اشتراكك. يرجى التواصل مع مالك النظام لتجديد الاشتراك.' 
                    : 'عذراً، حسابك مجمد. يرجى التواصل مع مالك النظام لتفعيل الحساب.';
                  alert(`${msg}\n\nالإيميل: fahussein79@gmail.com\nواتساب: 0550665495`);
                  return;
                }
                if (isAuthorized) {
                  fileInputRef.current?.click();
                } else {
                  alert('ليس لديك صلاحية رفع الملفات');
                }
              }}
              className={`border-2 border-dashed rounded-2xl p-8 flex flex-col items-center justify-center text-center group transition-all ${isFrozen ? 'border-red-200 bg-red-50/30 grayscale opacity-70 cursor-not-allowed' : isAuthorized ? 'border-indigo-100 bg-indigo-50/30 cursor-pointer hover:bg-indigo-50/50 hover:border-indigo-300' : 'border-slate-100 bg-slate-50 opacity-50 grayscale cursor-not-allowed'}`}
            >
              <div className={`w-12 h-12 bg-white rounded-2xl shadow-sm flex items-center justify-center mb-3 group-hover:scale-110 transition-transform ${isFrozen ? 'text-red-500' : 'text-indigo-600'}`}>
                {isFrozen ? <ShieldAlert size={24} /> : <FileSpreadsheet size={24} />}
              </div>
              <p className="text-sm font-semibold text-slate-700">
                {isFrozen ? 'الحساب مجمد' : 'اضغط لرفع ملف Excel'}
              </p>
              {isFrozen && (
                <div className="mt-3 px-4">
                  <p className="text-[10px] text-red-600 font-black leading-relaxed mb-2">
                    {frozenReason === 'expired' 
                      ? 'لا يمكنك رفع ملفات لأن اشتراكك منتهي.' 
                      : 'لا يمكنك رفع ملفات لأن حسابك مجمد.'}
                  </p>
                  <div className="flex flex-col items-center gap-1 opacity-80">
                    <p className="text-[10px] font-bold text-slate-500 font-mono">fahussein79@gmail.com</p>
                    <p className="text-[10px] font-bold text-slate-500 font-mono">0550665495</p>
                  </div>
                </div>
              )}
              {!isFrozen && !isAuthorized && <p className="text-[10px] text-red-500 font-bold mt-2">غير مصرح لك بالرفع</p>}
              <p className="text-xs text-slate-400 mt-1">يدعم XLSX, XLS, CSV</p>
            </div>
            <input
              ref={fileInputRef}
              type="file"
              className="hidden"
              accept=".xlsx, .xls, .csv"
              disabled={!isAuthorized || isFrozen}
              onChange={handleFileUpload}
            />
            
            {fileName && (
              <button 
                onClick={clearFile}
                className="w-full mt-4 py-2 text-sm font-medium text-red-500 bg-red-50 rounded-xl hover:bg-red-100 transition-colors flex items-center justify-center gap-2"
              >
                <X size={16} />
                إلغاء الملف
              </button>
            )}
          </section>

          {/* Statistics Cards */}
          <div className="grid grid-cols-2 gap-4">
            <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm group hover:border-indigo-200 transition-colors">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-normal mb-1">إجمالي السجلات</p>
              <p className="text-2xl font-bold text-slate-800 tracking-normal">{data.length.toLocaleString()}</p>
            </div>
            <div className="bg-white rounded-2xl p-4 border border-slate-200 shadow-sm group hover:border-indigo-200 transition-colors">
              <p className="text-[10px] font-bold text-slate-400 uppercase tracking-normal mb-1">عدد الأعمدة</p>
              <p className="text-2xl font-bold text-slate-800 tracking-normal">{headers.length}</p>
            </div>
          </div>

          {/* Financial Summary */}
          <section className="bg-white rounded-3xl p-6 border border-slate-200 shadow-sm overflow-hidden relative">
            <div className="absolute top-0 left-0 w-1.5 h-full bg-indigo-600"></div>
            <div className="flex items-center justify-between mb-6">
              <h2 className="text-xs font-bold text-slate-500 uppercase tracking-wider flex items-center gap-2">
                <BarChart3 size={16} className="text-indigo-600" />
                الملخص المالي
              </h2>
              <div className="flex gap-2">
                <span className="text-[9px] font-bold bg-indigo-50 text-indigo-600 px-2 py-1 rounded-lg border border-indigo-100">
                  {searchQuery.trim() || filterType !== 'all' ? 'فلترة نشطة' : 'البيانات كاملة'}
                </span>
              </div>
            </div>
            
            <div className="space-y-4">
              {/* Total Amount Card */}
              <div className="bg-slate-50 rounded-2xl p-5 border border-slate-200/50 relative overflow-hidden group">
                <div className="absolute -right-4 -top-4 text-slate-200/40 group-hover:scale-110 transition-transform">
                  <Wallet size={96} strokeWidth={1} />
                </div>
                <div className="relative z-10 text-right">
                  <div className="flex items-center justify-end gap-2 mb-2">
                    <span className="text-[10px] font-bold text-slate-400 uppercase">إجمالي المبلغ في الملف</span>
                    <div className="p-1.5 bg-white rounded-lg text-indigo-600 shadow-sm border border-slate-100">
                      <Wallet size={16} />
                    </div>
                  </div>
                  <div className="flex items-baseline justify-end gap-2">
                    <p className="text-3xl font-black text-slate-800 leading-none tracking-tight">
                      {totals.totalAmount.toLocaleString('ar-EG')}
                    </p>
                    <span className="text-xs font-bold text-slate-400">ر.س</span>
                  </div>
                  <div className="flex items-center justify-end gap-2 mt-3">
                    <span className="text-[9px] bg-slate-200/50 text-slate-500 px-2 py-0.5 rounded-full font-bold">
                      العمود: {totals.amountCol || 'غير محدد'}
                    </span>
                  </div>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-4">
                {/* Remaining Card */}
                <div className="bg-orange-50/50 rounded-2xl p-4 border border-orange-100/50 relative overflow-hidden group">
                  <div className="absolute -right-2 -top-2 text-orange-200/30 group-hover:scale-110 transition-transform">
                    <History size={64} strokeWidth={1} />
                  </div>
                  <div className="relative z-10">
                    <div className="flex items-center gap-2 mb-2">
                      <div className="p-1.5 bg-orange-100 rounded-lg text-orange-600">
                        <History size={16} />
                      </div>
                      <span className="text-[10px] font-bold text-orange-700">إجمالي المتبقي</span>
                    </div>
                    <div className="flex items-baseline gap-1">
                      <p className="text-xl font-black text-orange-900 tracking-tight">
                        {totals.totalRemaining.toLocaleString('ar-EG')}
                      </p>
                      <span className="text-[10px] font-bold text-orange-400 leading-none">ر.س</span>
                    </div>
                  </div>
                </div>

                {/* Settled Card */}
                <div className="bg-emerald-50/50 rounded-2xl p-4 border border-emerald-100/50 relative overflow-hidden group">
                  <div className="absolute -right-2 -top-2 text-emerald-200/30 group-hover:scale-110 transition-transform">
                    <CheckCircle size={64} strokeWidth={1} />
                  </div>
                  <div className="relative z-10">
                    <div className="flex items-center gap-2 mb-2">
                      <div className="p-1.5 bg-emerald-100 rounded-lg text-emerald-600">
                        <CheckCircle size={16} />
                      </div>
                      <span className="text-[10px] font-bold text-emerald-700">المسدد حالياً</span>
                    </div>
                    <div className="flex items-baseline gap-1">
                      <p className="text-xl font-black text-emerald-900 tracking-tight">
                        {totals.totalSettledFiltered.toLocaleString('ar-EG')}
                      </p>
                      <span className="text-[10px] font-bold text-emerald-400 leading-none">ر.س</span>
                    </div>
                  </div>
                </div>

                {/* Liquidation Card (Restored and Integrated) */}
                {totals.totalLiquidation > 0 && (
                  <div className="bg-red-50/50 rounded-2xl p-4 border border-red-100/50 relative overflow-hidden group">
                    <div className="absolute -right-2 -top-2 text-red-200/30 group-hover:scale-110 transition-transform">
                      <TrendingDown size={64} strokeWidth={1} />
                    </div>
                    <div className="relative z-10">
                      <div className="flex items-center gap-2 mb-2">
                        <div className="p-1.5 bg-red-100 rounded-lg text-red-600">
                          <TrendingDown size={16} />
                        </div>
                        <span className="text-[10px] font-bold text-red-700">إجمالي التصفية</span>
                      </div>
                      <div className="flex items-baseline gap-1">
                        <p className="text-xl font-black text-red-900 tracking-tight">
                          {totals.totalLiquidation.toLocaleString('ar-EG')}
                        </p>
                        <span className="text-[10px] font-bold text-red-400 leading-none">ر.س</span>
                      </div>
                    </div>
                  </div>
                )}
              </div>

              <div className="bg-slate-50 p-4 rounded-2xl border border-slate-100">
                <div className="flex items-start gap-3">
                  <div className="p-1.5 bg-indigo-100 rounded-lg text-indigo-600 shrink-0">
                    <Info size={14} />
                  </div>
                  <div>
                    <h4 className="text-[10px] font-black text-slate-800 uppercase mb-1">دليل الملخص</h4>
                    <p className="text-[10px] text-slate-400 leading-relaxed font-medium">
                      المسدد حالياً يمثل إجمالي مبالغ الصفوف المختارة في البحث المتاح حالياً، بينما المتبقي يمثل مجموع المبالغ غير المسددة.
                    </p>
                  </div>
                </div>
              </div>
            </div>
          </section>

          {!totals.amountCol && !totals.remainingCol && fileName && (
              <p className="mt-6 text-[10px] text-slate-400 italic text-center leading-relaxed">
                * لم يتم العثور على أعمدة مخصصة للمبالغ تلقائياً. تأكد من تسمية الأعمدة بكلمات مثل (المبلغ) أو (المتبقي).
              </p>
            )}

            {/* Final Save/Download Button */}
            <AnimatePresence>
              {isModified && (
                <motion.div
                  initial={{ opacity: 0, height: 0, marginTop: 0 }}
                  animate={{ opacity: 1, height: 'auto', marginTop: 24 }}
                  exit={{ opacity: 0, height: 0, marginTop: 0 }}
                  className="overflow-hidden"
                >
                  <button 
                    onClick={() => exportToExcel(true)}
                    className="w-full py-4 bg-emerald-600 text-white rounded-2xl font-black text-sm hover:bg-emerald-700 transition-all shadow-xl shadow-emerald-100 flex items-center justify-center gap-3 group"
                  >
                    <Database size={20} className="group-hover:scale-110 transition-transform" />
                    <span>حفظ وتحميل الملف المعدّل (كامل)</span>
                  </button>
                  <p className="text-[10px] text-emerald-600 font-bold mt-2 text-center flex items-center justify-center gap-1">
                    <Info size={10} />
                    تم إجراء تعديلات على البيانات
                  </p>
                </motion.div>
              )}
            </AnimatePresence>
          </aside>

        {/* Main Search Area */}
        <div className="col-span-12 lg:col-span-8 flex flex-col gap-6">
          {/* Search Bar */}
          <div className="relative group">
            <div className="absolute inset-y-0 right-6 flex items-center pointer-events-none text-slate-400 group-focus-within:text-indigo-600 transition-colors">
              <Search size={24} />
            </div>
            <input 
              type="text" 
              placeholder="ابحث في كافة الحقول والأعمدة (مثلاً: اسم، رقم، عقار، أو أي قيمة أخرى)..." 
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              disabled={!fileName || isLoading}
              className="w-full h-16 pr-16 pl-6 rounded-2xl border-none shadow-xl focus:ring-2 focus:ring-indigo-500 text-lg text-slate-700 placeholder-slate-400 bg-white transition-all disabled:bg-slate-50 disabled:cursor-not-allowed"
            />
            {searchQuery && (
              <button 
                onClick={() => setSearchQuery('')}
                className="absolute left-4 top-1/2 -translate-y-1/2 p-2 hover:bg-slate-100 rounded-lg text-slate-400 transition-colors"
                title="مسح البحث"
              >
                <X size={20} />
              </button>
            )}
          </div>

          {/* Results Container */}
          <div className="flex-1 flex flex-col gap-6 overflow-hidden min-h-[500px]">
            {isLoading ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm transition-all animate-pulse">
                <div className="w-12 h-12 border-4 border-indigo-600 border-t-transparent rounded-full animate-spin mb-4"></div>
                <p className="text-slate-500 font-medium">جاري معالجة البيانات...</p>
              </div>
            ) : !fileName ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm text-center p-8">
                <div className="w-20 h-20 bg-slate-50 rounded-full flex items-center justify-center mb-6 text-slate-200">
                  <Database size={48} strokeWidth={1} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">ابدأ برفع ملف للبحث</h3>
                <p className="text-slate-500 max-w-sm">ارفع ملف إكسل يحتوي على بياناتك، وسنساعدك في البحث عنها بسرعة فائقة.</p>
              </div>
            ) : filteredData.length === 0 ? (
              <div className="flex-1 flex flex-col items-center justify-center bg-white rounded-3xl border border-slate-200 shadow-sm text-center p-8">
                <div className="w-20 h-20 bg-red-50 rounded-full flex items-center justify-center mb-6 text-red-300">
                  <SearchSlash size={48} strokeWidth={1} />
                </div>
                <h3 className="text-xl font-bold text-slate-800 mb-2">لا توجد نتائج مطابقة</h3>
                <p className="text-slate-500 max-w-sm">
                  {filterType !== 'all' 
                    ? `لا توجد نتائج مطابقة لهذا الفلتر في ${searchQuery ? 'نتائج البحث' : 'الملف'}.`
                    : `لم نجد أي سجل يحتوي على "${searchQuery}". حاول تجربة كلمات بحث أخرى.`
                  }
                </p>
                {filterType !== 'all' && (
                  <button 
                    onClick={() => setFilterType('all')}
                    className="mt-4 text-indigo-600 font-bold hover:underline"
                  >
                    عرض كافة السجلات
                  </button>
                )}
              </div>
            ) : (
              <div className="flex flex-col gap-6 flex-1 overflow-hidden">
                {/* Search Results Header */}
                <div className="flex flex-col md:flex-row md:items-end justify-between gap-4 px-2">
                  <div className="flex-1">
                    <h1 className="text-2xl font-bold text-slate-800">نتائج البحث</h1>
                    <div className="flex items-center gap-2 mt-1">
                      <p className="text-slate-500 font-medium">تم العثور على {filteredData.length} سجل مطابق</p>
                      {filterType !== 'all' && (
                        <span className="text-[10px] bg-indigo-100 text-indigo-600 px-2 py-0.5 rounded-full font-bold uppercase">
                          مصفى بالحالة
                        </span>
                      )}
                    </div>
                  </div>
                  
                  {/* Filter Controls */}
                  <div className="flex flex-wrap items-center gap-2 bg-white/80 backdrop-blur-sm p-1.5 rounded-2xl border border-slate-200 shadow-sm overflow-x-auto max-w-full">
                    <button 
                      onClick={() => setFilterType('all')}
                      className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 shrink-0 ${filterType === 'all' ? 'bg-indigo-600 text-white shadow-md shadow-indigo-100 scale-105' : 'text-slate-500 hover:bg-slate-100'}`}
                    >
                      <LayoutGrid size={14} />
                      <span>الكل</span>
                      <span className={`px-1.5 py-0.5 rounded-lg text-[10px] ${filterType === 'all' ? 'bg-indigo-500 text-white' : 'bg-slate-100 text-slate-500'}`}>
                        {totals.counts.all}
                      </span>
                    </button>

                    <button 
                      onClick={() => setFilterType('paid')}
                      className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 shrink-0 ${filterType === 'paid' ? 'bg-emerald-500 text-white shadow-md shadow-emerald-100 scale-105' : 'text-emerald-600 hover:bg-emerald-50'}`}
                    >
                      <CheckCircle size={14} />
                      <span>المسدد</span>
                      <span className={`px-1.5 py-0.5 rounded-lg text-[10px] ${filterType === 'paid' ? 'bg-emerald-400 text-white' : 'bg-emerald-100 text-emerald-600'}`}>
                        {totals.counts.paid}
                      </span>
                    </button>

                    <button 
                      onClick={() => setFilterType('modified')}
                      className={`px-4 py-2 rounded-xl text-xs font-bold transition-all flex items-center gap-2 shrink-0 ${filterType === 'modified' ? 'bg-indigo-500 text-white shadow-md shadow-indigo-100 scale-105' : 'text-indigo-600 hover:bg-indigo-50'}`}
                    >
                      <FileEdit size={14} />
                      <span>المعدلة</span>
                      <span className={`px-1.5 py-0.5 rounded-lg text-[10px] ${filterType === 'modified' ? 'bg-indigo-400 text-white' : 'bg-indigo-100 text-indigo-600'}`}>
                        {totals.counts.modified}
                      </span>
                    </button>
                  </div>

                  <div className="flex gap-2">
                    <button 
                      onClick={() => exportToExcel(false)}
                      className="px-6 py-3 bg-indigo-600 text-white rounded-xl hover:bg-indigo-700 transition-all shadow-md flex items-center gap-2 group"
                      title="تحميل نتائج البحث الحالية"
                    >
                      <Download size={18} className="group-hover:translate-y-0.5 transition-transform" />
                      <span className="text-sm font-bold">تحميل النتائج</span>
                    </button>
                  </div>
                </div>

                {/* Table View */}
                <div className="flex-1 bg-white rounded-3xl border border-slate-200 shadow-sm overflow-hidden flex flex-col">
                  <div className="flex-1 overflow-auto custom-scrollbar">
                    <table className="w-full text-right border-collapse">
                      <thead className="sticky top-0 bg-slate-50 z-20 border-b border-slate-200">
                        <tr>
                          <th className="px-6 py-4 text-[10px] font-bold text-slate-400 uppercase text-center w-12 hover:text-indigo-600 transition-colors">#</th>
                          <th className="px-6 py-4 text-[10px] font-bold text-slate-400 uppercase text-center w-24">إجراء</th>
                          {headers.map((header, hIdx) => (
                            <th 
                              key={`${header}-${hIdx}`} 
                              className={`px-6 py-4 text-[10px] font-black text-slate-500 uppercase tracking-wider hover:text-indigo-600 transition-colors whitespace-nowrap relative group/tooltip ${
                                header.includes('الملاحظات') || header.includes('ملاحظات') || header.includes('notes') 
                                  ? 'min-w-[300px]' 
                                  : header.includes('الدفع') || header.includes('payment')
                                  ? 'min-w-[150px]'
                                  : 'min-w-[140px]'
                              }`}
                            >
                              <div className="flex items-center justify-center gap-1.5">
                                {header}
                                <Info size={12} className="text-slate-300 group-hover/tooltip:text-indigo-400 transition-colors" />
                              </div>
                              
                              {/* Tooltip */}
                              <div className="absolute bottom-full left-1/2 -translate-x-1/2 mb-2 px-3 py-2 bg-slate-800 text-white text-[10px] rounded-lg opacity-0 invisible group-hover/tooltip:opacity-100 group-hover/tooltip:visible transition-all duration-200 z-50 w-48 pointer-events-none shadow-xl border border-slate-700">
                                <div className="relative z-10 text-center font-medium leading-relaxed">
                                  {getHeaderTooltip(header)}
                                </div>
                                <div className="absolute top-full left-1/2 -translate-x-1/2 -mt-1 border-l-4 border-r-4 border-t-4 border-transparent border-t-slate-800"></div>
                              </div>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {filteredItems.map(({ row, index: originalIdx }, idx) => {
                          const isPaid = paidRows.has(originalIdx);
                          const isModifiedRow = modifiedRows.has(originalIdx);
                          
                          // Styling classes based on status
                          let rowBgClass = "hover:bg-indigo-50/30";
                          let textClass = "text-slate-600";
                          let numBgClass = "text-slate-400 bg-slate-50/30";
                          
                          if (isPaid) {
                            rowBgClass = "bg-emerald-50 hover:bg-emerald-100/80";
                            textClass = "text-emerald-700 font-bold";
                            numBgClass = "text-emerald-500 bg-emerald-50/50";
                          } else if (isModifiedRow) {
                            // Find any note-like field and check for "تصفية"
                            const rowString = JSON.stringify(row);
                            if (rowString.includes('تصفية')) {
                              rowBgClass = "bg-red-50/70 hover:bg-red-100/70 border-r-4 border-r-red-400";
                              textClass = "text-red-700 font-bold";
                              numBgClass = "text-red-500 bg-red-100/50";
                            } else {
                              rowBgClass = "bg-blue-50/70 hover:bg-blue-100/70 border-r-4 border-r-blue-400";
                              textClass = "text-blue-700 font-bold";
                              numBgClass = "text-blue-500 bg-blue-100/50";
                            }
                          }
                          
                          return (
                            <motion.tr
                              key={originalIdx}
                              initial={{ opacity: 0 }}
                              animate={{ opacity: 1 }}
                              transition={{ delay: Math.min(idx * 0.01, 0.2) }}
                              className={`transition-colors group ${rowBgClass}`}
                            >
                              <td className={`px-6 py-4 text-xs font-medium text-center ${numBgClass}`}>
                                {idx + 1}
                              </td>
                               <td className="px-4 py-2 text-center">
                                <div className="flex flex-col items-center gap-2">
                                  <button
                                    onClick={() => togglePaid(originalIdx)}
                                    className={`px-3 py-1.5 rounded-lg text-[10px] font-bold transition-all w-full max-w-[60px] ${
                                      isPaid 
                                        ? 'bg-emerald-500 text-white shadow-sm' 
                                        : 'bg-slate-100 text-slate-500 hover:bg-emerald-500 hover:text-white'
                                    }`}
                                  >
                                    {isPaid ? 'تم السداد' : 'تسوية'}
                                  </button>
                                  {modifiedRows.has(originalIdx) && (
                                    <button 
                                      onClick={() => undoRowChanges(originalIdx)}
                                      className="flex items-center gap-1 text-[9px] text-orange-500 hover:text-orange-600 font-bold bg-orange-50 px-2 py-0.5 rounded transition-colors"
                                      title="تراجع عن التعديل"
                                    >
                                      <RotateCcw size={10} />
                                      <span>تراجع</span>
                                    </button>
                                  )}
                                </div>
                              </td>
                              {headers.map((header, hIdx) => {
                                const isNoteColumn = header.includes('الملاحظات') || header.includes('ملاحظات') || header.includes('notes');
                                const isEditing = editingCell?.row === originalIdx && editingCell?.col === header;

                                const isAmount = header === totals.amountCol;
                                const isRemaining = header === totals.remainingCol;

                                return (
                                  <td 
                                    key={`${header}-${hIdx}`} 
                                    onClick={() => {
                                      if (isNoteColumn) {
                                        if (isAuthorized) {
                                          setEditingCell({ row: originalIdx, col: header });
                                          setEditingValue(String(row[header] || ''));
                                        } else {
                                          alert('ليس لديك صلاحية التعديل');
                                        }
                                      }
                                    }}
                                    className={`px-6 py-4 text-sm font-medium transition-all ${textClass} ${
                                      isNoteColumn 
                                        ? `whitespace-normal min-w-[300px] break-words ${isAuthorized ? 'cursor-pointer' : 'cursor-not-allowed opacity-80'}` 
                                        : header.includes('الدفع') || header.includes('payment')
                                        ? 'whitespace-normal min-w-[150px]'
                                        : 'whitespace-nowrap max-w-[250px] overflow-hidden text-ellipsis'
                                    }`}
                                  >
                                    {isEditing ? (
                                      <div className="flex items-center gap-2 w-full bg-white border-2 border-indigo-500 rounded-xl p-1 shadow-lg shadow-indigo-100 z-10 animate-in zoom-in-95 duration-200">
                                        <input 
                                          autoFocus
                                          className="flex-1 bg-transparent px-2 py-1 outline-none text-sm font-medium text-slate-700 min-w-0"
                                          value={editingValue}
                                          onChange={(e) => setEditingValue(e.target.value)}
                                          onKeyDown={(e) => {
                                            if (e.key === 'Enter') {
                                              updateNote(originalIdx, header, editingValue);
                                              setEditingCell(null);
                                            }
                                            if (e.key === 'Escape') {
                                              setEditingCell(null);
                                            }
                                          }}
                                          onBlur={() => {
                                            // Delay closing to allow button clicks to register
                                            setTimeout(() => {
                                              setEditingCell(prev => {
                                                if (prev?.row === originalIdx && prev?.col === header) {
                                                  return null;
                                                }
                                                return prev;
                                              });
                                            }, 250);
                                          }}
                                        />
                                        <div className="flex items-center gap-1 border-r border-slate-100 pr-1">
                                          <button 
                                            onMouseDown={(e) => e.preventDefault()}
                                            onClick={() => {
                                              updateNote(originalIdx, header, editingValue);
                                              setEditingCell(null);
                                            }}
                                            className="p-1.5 bg-emerald-50 text-emerald-600 hover:bg-emerald-100 rounded-lg transition-all"
                                            title="حفظ (Enter)"
                                          >
                                            <Check size={14} strokeWidth={3} />
                                          </button>
                                          <button 
                                            onMouseDown={(e) => e.preventDefault()}
                                            onClick={() => setEditingCell(null)}
                                            className="p-1.5 bg-red-50 text-red-500 hover:bg-red-100 rounded-lg transition-all"
                                            title="إلغاء (Esc)"
                                          >
                                            <X size={14} strokeWidth={3} />
                                          </button>
                                        </div>
                                      </div>
                                    ) : (
                                      <div className="flex items-start justify-between gap-2">
                                        <span>{highlightText(String(row[header] || '-'), searchQuery)}</span>
                                        {isNoteColumn && (
                                          <Info size={12} className="text-slate-300 mt-1 flex-shrink-0 opacity-0 group-hover:opacity-100 transition-opacity" />
                                        )}
                                      </div>
                                    )}
                                  </td>
                                );
                              })}
                            </motion.tr>
                          );
                        })}
                      </tbody>

                      {/* PDF/Print Financial Summary Row */}
                      <tfoot className="bg-slate-50 font-bold border-t-2 border-slate-200">
                        <tr>
                          <td className="px-6 py-4 text-xs text-slate-400 text-center">Σ</td>
                          <td className="px-6 py-4 text-xs text-slate-400 text-center">-</td>
                          {headers.map((header, hIdx) => {
                            const isAmount = header === totals.amountCol;
                            const isRemaining = header === totals.remainingCol;
                            
                            return (
                              <td key={`${header}-${hIdx}`} className={`px-6 py-4 text-sm ${isAmount ? 'text-indigo-600' : isRemaining ? 'text-orange-600' : 'text-slate-500'}`}>
                                {isAmount ? (
                                  <div className="flex items-center gap-2 whitespace-nowrap">
                                    <span className="text-[10px] text-slate-400 font-normal">إجمالي المبلغ:</span>
                                    <span>{totals.totalAmount.toLocaleString('ar-EG')} ر.س</span>
                                  </div>
                                ) : isRemaining ? (
                                  <div className="flex items-center gap-4 whitespace-nowrap">
                                    <div className="flex items-center gap-2">
                                      <span className="text-[10px] text-slate-400 font-normal">المتبقي:</span>
                                      <span>{totals.totalRemaining.toLocaleString('ar-EG')} ر.س</span>
                                    </div>
                                    <div className="flex items-center gap-2 bg-emerald-50 text-emerald-600 px-2 py-0.5 rounded-lg border border-emerald-100">
                                      <span className="text-[10px] opacity-70">المسدد حالياً:</span>
                                      <span className="font-bold">{totals.totalSettledFiltered.toLocaleString('ar-EG')} ر.س</span>
                                    </div>
                                    {totals.totalLiquidation > 0 && (
                                      <div className="flex items-center gap-2 bg-red-50 text-red-600 px-2 py-0.5 rounded-lg border border-red-100">
                                        <span className="text-[10px] opacity-70">إجمالي التصفية:</span>
                                        <span className="font-bold">{totals.totalLiquidation.toLocaleString('ar-EG')} ر.س</span>
                                      </div>
                                    )}
                                  </div>
                                ) : header === headers[0] ? (
                                  <div className="flex items-center gap-2 text-indigo-600">
                                    <BarChart3 size={14} />
                                    <span>ملخص مالي</span>
                                  </div>
                                ) : ''}
                              </td>
                            );
                          })}
                        </tr>
                      </tfoot>
                    </table>
                  </div>
                  
                  {/* Table Footer / Summary */}
                  <div className="p-4 bg-slate-50 border-t border-slate-200 flex items-center justify-between text-[10px] font-bold text-slate-400 uppercase tracking-normal">
                    <span>نهاية النتائج المطابقة</span>
                    <span>إجمالي الصفوف المعروضة: {filteredData.length}</span>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </main>

      {/* Bottom Status Bar */}
      <footer className="h-10 bg-slate-800 px-6 flex items-center justify-between text-[10px] text-slate-400 uppercase tracking-normal shrink-0 mt-auto">
        <div className="flex items-center gap-4">
          <div className="flex items-center gap-2">
            <span className="w-1.5 h-1.5 bg-green-400 rounded-full shadow-[0_0_8px_rgba(74,222,128,0.5)]"></span>
            <span>نظام البحث: متصل</span>
          </div>
          <span className="w-1 h-1 bg-slate-600 rounded-full hidden sm:block"></span>
          <span className="hidden sm:block">آخر تحديث: {new Date().toLocaleTimeString('ar-EG')}</span>
        </div>
        <div className="flex items-center gap-2">
          <span className="text-indigo-400 font-bold">SMART ENGINE V1.0</span>
        </div>
      </footer>
    </div>
  );
}

function headerLabel(header: string) {
  return header.replace(/([A-Z])/g, ' $1').trim();
}

function highlightText(text: string, highlight: string) {
  if (!highlight.trim()) return <span>{text}</span>;
  
  // Escape special characters for regex to avoid errors if user types things like ( or [
  const escapedHighlight = highlight.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const parts = text.split(new RegExp(`(${escapedHighlight})`, 'gi'));
  
  return (
    <span className="inline-block">
      {parts.map((part, i) => (
        <span
          key={i}
          className={part.toLowerCase() === highlight.toLowerCase() ? 'bg-amber-300 text-amber-950 px-0.5 rounded-sm font-bold shadow-sm' : ''}
        >
          {part}
        </span>
      ))}
    </span>
  );
}

