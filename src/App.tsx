import React, { createContext, useContext, useEffect, useState, useMemo } from 'react';
import { 
  BrowserRouter as Router, 
  Routes, 
  Route, 
  Navigate, 
  Link, 
  useLocation, 
  useNavigate,
  useSearchParams
} from 'react-router-dom';
import { 
  collection, 
  doc, 
  onSnapshot, 
  query, 
  orderBy, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  serverTimestamp, 
} from 'firebase/firestore';
import ExcelJS from 'exceljs';
import { saveAs } from 'file-saver';
import { 
  LayoutDashboard, 
  Package, 
  History, 
  Plus, 
  Search, 
  Filter, 
  AlertTriangle, 
  ArrowUpRight, 
  ArrowDownLeft, 
  Edit2, 
  Trash2, 
  Menu, 
  X,
  ChevronRight,
  Loader2,
  Users,
  Building2,
  Download,
  Image as ImageIcon,
  Upload,
  Check,
  FileUp,
  Clock,
  Coffee,
  LogOut,
  LogIn,
  RotateCcw,
  Calendar,
  ClipboardList,
  ShoppingCart,
  BookOpen,
  CheckCircle2,
  QrCode,
  Maximize,
  ExternalLink
} from 'lucide-react';
import { motion, AnimatePresence } from 'motion/react';
import { QRCodeCanvas } from 'qrcode.react';
import { Html5QrcodeScanner } from 'html5-qrcode';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';
import { format, differenceInMinutes, startOfMonth, endOfMonth, isWithinInterval } from 'date-fns';

import { db } from './firebase';
import { Product, Movement, Requester, Attendance, ProductRequest, DailyReport } from './types';

// --- Utils ---
function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

// --- Components ---

const Button = React.forwardRef<HTMLButtonElement, React.ButtonHTMLAttributes<HTMLButtonElement> & { variant?: 'primary' | 'secondary' | 'danger' | 'ghost' | 'outline' | 'success', size?: 'sm' | 'md' | 'lg' }>(
  ({ className, variant = 'primary', size = 'md', ...props }, ref) => {
    const variants = {
      primary: 'bg-blue-600 text-white hover:bg-blue-700 shadow-sm',
      secondary: 'bg-gray-100 text-gray-900 hover:bg-gray-200',
      danger: 'bg-red-600 text-white hover:bg-red-700 shadow-sm',
      ghost: 'bg-transparent hover:bg-gray-100 text-gray-600',
      outline: 'bg-transparent border border-gray-300 text-gray-700 hover:bg-gray-50',
      success: 'bg-green-600 text-white hover:bg-green-700 shadow-sm',
    };
    const sizes = {
      sm: 'px-3 py-1.5 text-sm',
      md: 'px-4 py-2',
      lg: 'px-6 py-3 text-lg',
    };
    return (
      <button
        ref={ref}
        className={cn(
          'inline-flex items-center justify-center rounded-lg font-medium transition-colors focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-offset-2 disabled:opacity-50 disabled:pointer-events-none',
          variants[variant],
          sizes[size],
          className
        )}
        {...props}
      />
    );
  }
);

const Input = React.forwardRef<HTMLInputElement, React.InputHTMLAttributes<HTMLInputElement> & { label?: string, error?: string }>(
  ({ className, label, error, ...props }, ref) => {
    return (
      <div className="space-y-1.5 w-full">
        {label && <label className="text-sm font-medium text-gray-700">{label}</label>}
        <input
          ref={ref}
          className={cn(
            'flex h-10 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm placeholder:text-gray-400 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent disabled:cursor-not-allowed disabled:opacity-50 transition-all',
            error && 'border-red-500 focus:ring-red-500',
            className
          )}
          {...props}
        />
        {error && <p className="text-xs text-red-500">{error}</p>}
      </div>
    );
  }
);

const Select = React.forwardRef<HTMLSelectElement, React.SelectHTMLAttributes<HTMLSelectElement> & { label?: string, error?: string }>(
  ({ className, label, error, children, ...props }, ref) => {
    return (
      <div className="space-y-1.5 w-full">
        {label && <label className="text-sm font-medium text-gray-700">{label}</label>}
        <select
          ref={ref}
          className={cn(
            'flex h-10 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent disabled:cursor-not-allowed disabled:opacity-50 transition-all',
            error && 'border-red-500 focus:ring-red-500',
            className
          )}
          {...props}
        >
          {children}
        </select>
        {error && <p className="text-xs text-red-500">{error}</p>}
      </div>
    );
  }
);

const Modal: React.FC<{ isOpen: boolean, onClose: () => void, title: string, children: React.ReactNode }> = ({ isOpen, onClose, title, children }) => {
  return (
    <AnimatePresence>
      {isOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 sm:p-6">
          <motion.div
            initial={{ opacity: 0 }}
            animate={{ opacity: 1 }}
            exit={{ opacity: 0 }}
            onClick={onClose}
            className="absolute inset-0 bg-black/40 backdrop-blur-sm"
          />
          <motion.div
            initial={{ opacity: 0, scale: 0.95, y: 20 }}
            animate={{ opacity: 1, scale: 1, y: 0 }}
            exit={{ opacity: 0, scale: 0.95, y: 20 }}
            className="relative w-full max-w-lg overflow-hidden rounded-2xl bg-white shadow-2xl"
          >
            <div className="flex items-center justify-between border-b border-gray-100 px-6 py-4">
              <h3 className="text-lg font-semibold text-gray-900">{title}</h3>
              <button onClick={onClose} className="rounded-full p-1 hover:bg-gray-100 transition-colors">
                <X className="h-5 w-5 text-gray-500" />
              </button>
            </div>
            <div className="p-6">
              {children}
            </div>
          </motion.div>
        </div>
      )}
    </AnimatePresence>
  );
};

// --- Pages ---

const Layout: React.FC<{ children: React.ReactNode }> = ({ children }) => {
  const location = useLocation();
  const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);
  const [products, setProducts] = useState<Product[]>([]);

  useEffect(() => {
    const unsub = onSnapshot(collection(db, 'products'), (snap) => {
      setProducts(snap.docs.map(d => ({ id: d.id, ...d.data() } as Product)));
    });
    return unsub;
  }, []);

  const categories = useMemo(() => {
    const cats = new Set(products.map(p => p.category));
    return Array.from(cats).sort();
  }, [products]);

  const warehouseItems = [
    { name: 'Dashboard', path: '/', icon: LayoutDashboard },
    { name: 'Todo el Inventario', path: '/inventory', icon: Package },
    { name: 'Movimientos', path: '/movements', icon: History },
    { name: 'Personal Autorizado', path: '/requesters', icon: Users },
  ];

  const managementItems = [
    { name: 'Asistencia', path: '/attendance', icon: Clock },
    { name: 'Solicitudes', path: '/requests', icon: ClipboardList },
    { name: 'Bitácora', path: '/reports', icon: BookOpen },
    { name: 'Escáner QR', path: '/scanner', icon: Maximize },
  ];

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col md:flex-row">
      {/* Mobile Header */}
      <div className="md:hidden flex items-center justify-between px-6 py-4 bg-white border-b border-gray-200">
        <div className="flex items-center gap-2">
          <Building2 className="h-6 w-6 text-blue-600" />
          <span className="font-bold text-xl">Bodega Master</span>
        </div>
        <button onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}>
          {isMobileMenuOpen ? <X className="h-6 w-6" /> : <Menu className="h-6 w-6" />}
        </button>
      </div>

      {/* Sidebar */}
      <aside className={cn(
        "fixed inset-y-0 left-0 z-40 w-64 bg-white border-r border-gray-200 transform transition-transform duration-300 ease-in-out md:relative md:translate-x-0",
        isMobileMenuOpen ? "translate-x-0" : "-translate-x-full"
      )}>
        <div className="flex flex-col h-full">
          <div className="hidden md:flex items-center gap-3 px-8 py-8">
            <div className="p-2 bg-blue-600 rounded-xl text-white">
              <Building2 className="h-6 w-6" />
            </div>
            <span className="font-bold text-xl tracking-tight">Bodega Master</span>
          </div>

          <div className="flex-1 overflow-y-auto px-4 space-y-6">
            <nav className="space-y-1">
              <p className="px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-2">Bodega</p>
              {warehouseItems.map((item) => (
                <Link
                  key={item.path}
                  to={item.path}
                  onClick={() => setIsMobileMenuOpen(false)}
                  className={cn(
                    "flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all group",
                    location.pathname === item.path && !location.search
                      ? "bg-blue-50 text-blue-600" 
                      : "text-gray-500 hover:bg-gray-50 hover:text-gray-900"
                  )}
                >
                  <item.icon className={cn(
                    "h-5 w-5",
                    location.pathname === item.path && !location.search ? "text-blue-600" : "text-gray-400 group-hover:text-gray-600"
                  )} />
                  {item.name}
                </Link>
              ))}
            </nav>

            <nav className="space-y-1">
              <p className="px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-2">Gestión de Obra</p>
              {managementItems.map((item) => (
                <Link
                  key={item.path}
                  to={item.path}
                  onClick={() => setIsMobileMenuOpen(false)}
                  className={cn(
                    "flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-all group",
                    location.pathname === item.path && !location.search
                      ? "bg-blue-50 text-blue-600" 
                      : "text-gray-500 hover:bg-gray-50 hover:text-gray-900"
                  )}
                >
                  <item.icon className={cn(
                    "h-5 w-5",
                    location.pathname === item.path && !location.search ? "text-blue-600" : "text-gray-400 group-hover:text-gray-600"
                  )} />
                  {item.name}
                </Link>
              ))}
            </nav>

            <nav className="space-y-1">
              <p className="px-4 text-[10px] font-bold text-gray-400 uppercase tracking-widest mb-2">Categorías</p>
              {categories.length === 0 ? (
                <p className="px-4 text-xs text-gray-400 italic">Sin categorías</p>
              ) : (
                categories.map((cat) => (
                  <Link
                    key={cat}
                    to={`/inventory?category=${encodeURIComponent(cat)}`}
                    onClick={() => setIsMobileMenuOpen(false)}
                    className={cn(
                      "flex items-center gap-3 px-4 py-2 rounded-xl text-sm font-medium transition-all group",
                      location.search === `?category=${encodeURIComponent(cat)}`
                        ? "bg-blue-50 text-blue-600" 
                        : "text-gray-500 hover:bg-gray-50 hover:text-gray-900"
                    )}
                  >
                    <div className={cn(
                      "w-1.5 h-1.5 rounded-full",
                      location.search === `?category=${encodeURIComponent(cat)}` ? "bg-blue-600" : "bg-gray-300 group-hover:bg-gray-400"
                    )} />
                    {cat}
                  </Link>
                ))
              )}
            </nav>
          </div>
          
          <div className="p-8 text-xs text-gray-400 text-center border-t border-gray-50">
            Gestión Pública de Bodega
          </div>
        </div>
      </aside>

      {/* Main Content */}
      <main className="flex-1 overflow-y-auto p-6 md:p-10">
        <div className="max-w-7xl mx-auto">
          {children}
        </div>
      </main>
    </div>
  );
};

const Dashboard = () => {
  const [products, setProducts] = useState<Product[]>([]);
  const [movements, setMovements] = useState<Movement[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    const unsubProducts = onSnapshot(collection(db, 'products'), (snap) => {
      setProducts(snap.docs.map(d => ({ id: d.id, ...d.data() } as Product)));
      setLoading(false);
    });
    const unsubMovements = onSnapshot(query(collection(db, 'movements'), orderBy('date', 'desc')), (snap) => {
      setMovements(snap.docs.map(d => ({ id: d.id, ...d.data() } as Movement)));
    });
    return () => { unsubProducts(); unsubMovements(); };
  }, []);

  const stats = useMemo(() => {
    const totalProducts = products.length;
    const totalStock = products.reduce((acc, p) => acc + p.quantity, 0);
    const orderedProducts = products.filter(p => p.status === 'ordered').length;
    const recentMovements = movements.slice(0, 5);
    return { totalProducts, totalStock, orderedProducts, recentMovements };
  }, [products, movements]);

  if (loading) return <div className="flex items-center justify-center h-64"><Loader2 className="h-8 w-8 animate-spin text-blue-600" /></div>;

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-3xl font-bold text-gray-900">Panel de Resumen</h1>
        <p className="text-gray-500">Estado actual de la bodega</p>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <StatCard title="Total Productos" value={stats.totalProducts} icon={Package} color="blue" />
        <StatCard title="Artículos Totales" value={stats.totalStock} icon={Building2} color="green" />
        <StatCard title="Productos Pedidos" value={stats.orderedProducts} icon={ShoppingCart} color="yellow" isWarning={stats.orderedProducts > 0} />
        <StatCard title="Movimientos Hoy" value={movements.filter(m => {
          const d = m.date?.toDate ? m.date.toDate() : new Date(m.date);
          return d.toDateString() === new Date().toDateString();
        }).length} icon={History} color="purple" />
      </div>

      <div className="grid grid-cols-1 gap-8">
        <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
          <h2 className="text-xl font-bold text-gray-900 mb-6">Actividad Reciente</h2>
          <div className="space-y-6">
            {stats.recentMovements.length === 0 ? (
              <p className="text-gray-500 text-center py-8">Sin actividad reciente.</p>
            ) : (
              stats.recentMovements.map(m => (
                <div key={m.id} className="flex gap-4">
                  <div className={cn(
                    "mt-1 p-2 rounded-lg shrink-0",
                    m.type === 'entry' ? "bg-green-100 text-green-600" : "bg-orange-100 text-orange-600"
                  )}>
                    {m.type === 'entry' ? <ArrowDownLeft className="h-4 w-4" /> : <ArrowUpRight className="h-4 w-4" />}
                  </div>
                  <div className="flex-1 min-w-0">
                    <p className="text-sm font-semibold text-gray-900 truncate">{m.productName}</p>
                    <p className="text-xs text-gray-500">
                      {m.type === 'entry' ? 'Entrada' : 'Salida'} de {m.quantity} unidades
                    </p>
                    <p className="text-[10px] text-gray-400 mt-1">
                      {m.date?.toDate ? format(m.date.toDate(), 'HH:mm - dd/MM') : 'Reciente'}
                    </p>
                  </div>
                </div>
              ))
            )}
          </div>
        </div>
      </div>
    </div>
  );
};

const StatCard = ({ title, value, icon: Icon, color, isWarning }: any) => {
  const colors: any = {
    blue: 'bg-blue-50 text-blue-600',
    red: 'bg-red-50 text-red-600',
    green: 'bg-green-50 text-green-600',
    purple: 'bg-purple-50 text-purple-600',
    yellow: 'bg-yellow-50 text-yellow-600',
  };
  return (
    <div className={cn(
      "bg-white p-6 rounded-2xl shadow-sm border border-gray-100 transition-all hover:shadow-md",
      isWarning && "ring-2 ring-red-500 ring-offset-2"
    )}>
      <div className="flex items-center justify-between mb-4">
        <div className={cn("p-3 rounded-xl", colors[color])}>
          <Icon className="h-6 w-6" />
        </div>
      </div>
      <div>
        <p className="text-sm font-medium text-gray-500">{title}</p>
        <p className="text-2xl font-bold text-gray-900 mt-1">{value}</p>
      </div>
    </div>
  );
};

const Inventory = () => {
  const [searchParams, setSearchParams] = useSearchParams();
  const categoryFromUrl = searchParams.get('category');
  
  const [products, setProducts] = useState<Product[]>([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState('');
  const [categoryFilter, setCategoryFilter] = useState(categoryFromUrl || 'all');
  
  const [isModalOpen, setIsModalOpen] = useState(false);
  const [isMovementModalOpen, setIsMovementModalOpen] = useState(false);
  const [isQrModalOpen, setIsQrModalOpen] = useState(false);
  const [editingProduct, setEditingProduct] = useState<Product | null>(null);
  const [selectedProduct, setSelectedProduct] = useState<Product | null>(null);
  const [qrProduct, setQrProduct] = useState<Product | null>(null);
  const [imagePreview, setImagePreview] = useState<string | null>(null);
  const [requesters, setRequesters] = useState<Requester[]>([]);

  const [movements, setMovements] = useState<Movement[]>([]);

  useEffect(() => {
    const unsubRequesters = onSnapshot(query(collection(db, 'requesters'), orderBy('name')), (snap) => {
      setRequesters(snap.docs.map(d => ({ id: d.id, ...d.data() } as Requester)));
    });
    return unsubRequesters;
  }, []);

  useEffect(() => {
    if (categoryFromUrl) {
      setCategoryFilter(categoryFromUrl);
    } else {
      setCategoryFilter('all');
    }
  }, [categoryFromUrl]);

  useEffect(() => {
    const unsub = onSnapshot(collection(db, 'products'), (snap) => {
      setProducts(snap.docs.map(d => ({ id: d.id, ...d.data() } as Product)));
      setLoading(false);
    });
    const unsubMovements = onSnapshot(query(collection(db, 'movements'), orderBy('date', 'desc')), (snap) => {
      setMovements(snap.docs.map(d => ({ id: d.id, ...d.data() } as Movement)));
    });
    return () => { unsub(); unsubMovements(); };
  }, []);

  const categories = useMemo(() => {
    const cats = new Set(products.map(p => p.category));
    return ['all', ...Array.from(cats)];
  }, [products]);

  const filteredProducts = products.filter(p => {
    const matchesSearch = p.name.toLowerCase().includes(search.toLowerCase()) || p.category.toLowerCase().includes(search.toLowerCase());
    const matchesCategory = categoryFilter === 'all' || p.category === categoryFilter;
    return matchesSearch && matchesCategory;
  });

  const handleImageChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      const reader = new FileReader();
      reader.onloadend = () => {
        // Simple compression using canvas
        const img = new Image();
        img.src = reader.result as string;
        img.onload = () => {
          const canvas = document.createElement('canvas');
          const MAX_WIDTH = 400;
          const MAX_HEIGHT = 400;
          let width = img.width;
          let height = img.height;

          if (width > height) {
            if (width > MAX_WIDTH) {
              height *= MAX_WIDTH / width;
              width = MAX_WIDTH;
            }
          } else {
            if (height > MAX_HEIGHT) {
              width *= MAX_HEIGHT / height;
              height = MAX_HEIGHT;
            }
          }
          canvas.width = width;
          canvas.height = height;
          const ctx = canvas.getContext('2d');
          ctx?.drawImage(img, 0, 0, width, height);
          const compressedBase64 = canvas.toDataURL('image/jpeg', 0.7);
          setImagePreview(compressedBase64);
        };
      };
      reader.readAsDataURL(file);
    }
  };

  const handleSaveProduct = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    const formData = new FormData(e.currentTarget);
    const productData = {
      name: formData.get('name') as string,
      category: formData.get('category') as string,
      quantity: Number(formData.get('quantity')),
      unit: formData.get('unit') as string,
      status: formData.get('status') as 'available' | 'ordered',
      useQR: formData.get('useQR') === 'on',
      imageUrl: imagePreview || editingProduct?.imageUrl || '',
      updatedAt: serverTimestamp(),
    };

    try {
      if (editingProduct) {
        await updateDoc(doc(db, 'products', editingProduct.id), productData);
      } else {
        await addDoc(collection(db, 'products'), {
          ...productData,
          entryDate: serverTimestamp(),
        });
      }
      setIsModalOpen(false);
      setEditingProduct(null);
    } catch (err) {
      console.error(err);
      alert('Error al guardar el producto');
    }
  };

  const handleDeleteProduct = async (id: string) => {
    if (window.confirm('¿Estás seguro de eliminar este producto?')) {
      await deleteDoc(doc(db, 'products', id));
    }
  };

  const handleMovement = async (e: React.FormEvent<HTMLFormElement>) => {
    e.preventDefault();
    if (!selectedProduct) return;
    const formData = new FormData(e.currentTarget);
    const type = formData.get('type') as 'entry' | 'exit';
    const amount = Number(formData.get('amount'));
    const notes = formData.get('notes') as string;

    const newQuantity = type === 'entry' 
      ? selectedProduct.quantity + amount 
      : selectedProduct.quantity - amount;

    if (newQuantity < 0) {
      alert('No hay suficiente stock para esta salida');
      return;
    }

    try {
      // Update product quantity
      await updateDoc(doc(db, 'products', selectedProduct.id), {
        quantity: newQuantity,
        updatedAt: serverTimestamp(),
      });

      // Record movement
      await addDoc(collection(db, 'movements'), {
        productId: selectedProduct.id,
        productName: selectedProduct.name,
        type,
        quantity: amount,
        date: serverTimestamp(),
        userName: 'Usuario Anónimo',
        givenBy: formData.get('givenBy') as string,
        receivedBy: formData.get('receivedBy') as string,
        notes,
      });

      setIsMovementModalOpen(false);
      setSelectedProduct(null);
    } catch (err) {
      console.error(err);
      alert('Error al registrar el movimiento');
    }
  };

  const [isExporting, setIsExporting] = useState(false);
  const [exportSuccess, setExportSuccess] = useState(false);
  const [isImportingExcel, setIsImportingExcel] = useState(false);

  const handleImportExcel = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setIsImportingExcel(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const arrayBuffer = await file.arrayBuffer();
      await workbook.xlsx.load(arrayBuffer);
      const worksheet = workbook.getWorksheet(1);
      
      if (!worksheet) throw new Error("No se encontró la hoja de cálculo");

      const newProducts: any[] = [];
      worksheet.eachRow((row, rowNumber) => {
        if (rowNumber === 1) return; // Skip header

        const name = row.getCell(1).value?.toString();
        const category = row.getCell(2).value?.toString();
        const quantity = Number(row.getCell(3).value);
        const unit = row.getCell(4).value?.toString() || 'unidades';

        if (name && category && !isNaN(quantity)) {
          newProducts.push({
            name,
            category,
            quantity,
            unit,
            status: 'available',
            useQR: false,
            entryDate: serverTimestamp(),
            updatedAt: serverTimestamp(),
          });
        }
      });

      if (newProducts.length === 0) {
        alert("No se encontraron productos válidos en el archivo. El formato debe ser: Nombre, Categoría, Cantidad, Unidad, Stock Mínimo.");
        return;
      }

      const batch = newProducts.map(p => addDoc(collection(db, 'products'), p));
      await Promise.all(batch);
      alert(`Se importaron ${newProducts.length} productos exitosamente.`);
    } catch (err) {
      console.error(err);
      alert("Error al importar el archivo Excel. Asegúrate de que el formato sea correcto.");
    } finally {
      setIsImportingExcel(false);
      e.target.value = '';
    }
  };

  const downloadTemplate = () => {
    const workbook = new ExcelJS.Workbook();
    const sheet = workbook.addWorksheet('Plantilla Importación');
    sheet.columns = [
      { header: 'Nombre', key: 'name', width: 30 },
      { header: 'Categoría', key: 'category', width: 20 },
      { header: 'Cantidad', key: 'quantity', width: 15 },
      { header: 'Unidad', key: 'unit', width: 15 },
    ];
    sheet.addRow(['Cemento Gris', 'Obra Gris', 50, 'sacos']);
    sheet.addRow(['Varilla 3/8', 'Hierro', 100, 'unidades']);
    
    workbook.xlsx.writeBuffer().then(buffer => {
      saveAs(new Blob([buffer]), 'Plantilla_Importacion_Bodega.xlsx');
    });
  };

  const exportToExcel = async () => {
    setIsExporting(true);
    try {
      const workbook = new ExcelJS.Workbook();
    workbook.creator = 'Bodega Master';
    workbook.lastModifiedBy = 'Bodega Master';
    workbook.created = new Date();

    // --- Sheet 1: Inventario General ---
    const sheet1 = workbook.addWorksheet('Inventario General');
    sheet1.columns = [
      { header: 'Producto', key: 'name', width: 30 },
      { header: 'Categoría', key: 'category', width: 20 },
      { header: 'Cantidad Actual', key: 'quantity', width: 15 },
      { header: 'Unidad', key: 'unit', width: 10 },
      { header: 'Ubicación', key: 'location', width: 20 },
      { header: 'Última Fecha Mov.', key: 'lastDate', width: 20 },
      { header: 'Último Retirante', key: 'lastPerson', width: 25 },
    ];

    // Style headers
    sheet1.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet1.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } }; // Slate-800
    sheet1.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    filteredProducts.forEach(p => {
      const productMovements = movements.filter(m => m.productId === p.id);
      const lastMov = productMovements[0]; // Already sorted by date desc
      const lastExit = productMovements.find(m => m.type === 'exit');

      const row = sheet1.addRow({
        name: p.name,
        category: p.category,
        quantity: p.quantity,
        unit: p.unit,
        location: p.location || 'N/A',
        lastDate: lastMov ? format(lastMov.date?.toDate ? lastMov.date.toDate() : new Date(lastMov.date), 'dd/MM/yyyy HH:mm') : 'Sin mov.',
        lastPerson: lastExit ? lastExit.receivedBy : 'N/A'
      });

      // Style cells
      const cell = row.getCell('quantity');
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFDCFCE7' } }; // Green-100
      cell.font = { color: { argb: 'FF166534' }, bold: true }; // Green-800

      // Add borders to all cells in row
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    // --- Sheet 2: Historial de Movimientos ---
    const sheet2 = workbook.addWorksheet('Historial de Movimientos');
    sheet2.columns = [
      { header: 'Fecha y Hora', key: 'date', width: 20 },
      { header: 'Producto', key: 'product', width: 30 },
      { header: 'Tipo', key: 'type', width: 12 },
      { header: 'Cantidad', key: 'quantity', width: 12 },
      { header: 'Entregado por (Admin)', key: 'givenBy', width: 25 },
      { header: 'Recibido por (Solicitante)', key: 'receivedBy', width: 25 },
    ];

    sheet2.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
    sheet2.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
    sheet2.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

    movements.forEach(m => {
      const row = sheet2.addRow({
        date: m.date?.toDate ? format(m.date.toDate(), 'dd/MM/yyyy HH:mm') : '-',
        product: m.productName,
        type: m.type === 'entry' ? 'ENTRADA' : 'SALIDA',
        quantity: m.quantity,
        givenBy: m.givenBy || '-',
        receivedBy: m.receivedBy || '-'
      });

      const typeCell = row.getCell('type');
      if (m.type === 'entry') {
        typeCell.font = { color: { argb: 'FF166534' }, bold: true };
      } else {
        typeCell.font = { color: { argb: 'FF991B1B' }, bold: true };
      }

      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });
    });

    // --- Sheet 3: Resumen ---
    const sheet3 = workbook.addWorksheet('Resumen');
    sheet3.getColumn(1).width = 30;
    sheet3.getColumn(2).width = 40;

    const addSummaryRow = (label: string, value: any, isHeader = false) => {
      const row = sheet3.addRow([label, value]);
      if (isHeader) {
        row.getCell(1).font = { bold: true, size: 14 };
        row.getCell(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFCBD5E1' } };
      }
      row.getCell(1).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
      row.getCell(2).border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    };

    const mostRequested = movements
      .filter(m => m.type === 'exit')
      .reduce((acc: any, m) => {
        acc[m.productName] = (acc[m.productName] || 0) + m.quantity;
        return acc;
      }, {});
    
    const topProduct = Object.entries(mostRequested).sort((a: any, b: any) => b[1] - a[1])[0];
    const lastWithdrawal = movements.find(m => m.type === 'exit');

    sheet3.addRow(['REPORTE EJECUTIVO DE BODEGA', '']).getCell(1).font = { bold: true, size: 16 };
    sheet3.addRow(['Fecha de Generación:', format(new Date(), 'dd/MM/yyyy HH:mm')]);
    sheet3.addRow([]);

    addSummaryRow('INDICADOR', 'VALOR', true);
    addSummaryRow('Total de Productos en Catálogo', products.length);
    addSummaryRow('Total de Movimientos Registrados', movements.length);
    addSummaryRow('Producto Más Solicitado (Salidas)', topProduct ? `${topProduct[0]} (${topProduct[1]} unidades)` : 'N/A');
    addSummaryRow('Última Persona que Retiró Material', lastWithdrawal ? lastWithdrawal.receivedBy : 'N/A');

    // Generate and Save
    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `Reporte_Bodega_${format(new Date(), 'yyyy-MM-dd')}.xlsx`);
    setExportSuccess(true);
    setTimeout(() => setExportSuccess(false), 3000);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Inventario</h1>
          <p className="text-gray-500">Gestiona tus materiales y herramientas</p>
        </div>
        <div className="flex gap-2">
          <input 
            type="file" 
            id="excel-import" 
            className="hidden" 
            accept=".xlsx, .xls" 
            onChange={handleImportExcel}
          />
          <Button 
            variant="outline" 
            onClick={() => document.getElementById('excel-import')?.click()} 
            disabled={isImportingExcel}
            className="gap-2"
          >
            {isImportingExcel ? <Loader2 className="h-5 w-5 animate-spin" /> : <FileUp className="h-5 w-5" />}
            {isImportingExcel ? 'Importando...' : 'Importar Excel'}
          </Button>
          <Button 
            variant={exportSuccess ? "success" : "outline"} 
            onClick={exportToExcel} 
            disabled={isExporting} 
            className="gap-2 transition-all"
          >
            {isExporting ? <Loader2 className="h-5 w-5 animate-spin" /> : exportSuccess ? <Check className="h-5 w-5" /> : <Download className="h-5 w-5" />}
            {isExporting ? 'Exportando...' : exportSuccess ? '¡Listo!' : 'Exportar Excel'}
          </Button>
          <Button onClick={() => { setEditingProduct(null); setIsModalOpen(true); }} className="gap-2">
            <Plus className="h-5 w-5" />
            Nuevo Producto
          </Button>
          <Button variant="ghost" size="sm" onClick={downloadTemplate} className="text-xs text-gray-400 hover:text-blue-600">
            Descargar Plantilla
          </Button>
        </div>
      </div>

      <div className="flex flex-col md:flex-row gap-4">
        <div className="relative flex-1">
          <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-5 w-5 text-gray-400" />
          <Input 
            placeholder="Buscar por nombre o categoría..." 
            className="pl-10"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
          />
        </div>
        <div className="flex gap-2">
          <div className="relative">
            <Filter className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-gray-400" />
            <Select 
              className="pl-9 min-w-[150px]"
              value={categoryFilter}
              onChange={(e) => {
                const val = e.target.value;
                setCategoryFilter(val);
                if (val === 'all') {
                  searchParams.delete('category');
                } else {
                  searchParams.set('category', val);
                }
                setSearchParams(searchParams);
              }}
            >
              {categories.map(cat => (
                <option key={cat} value={cat}>{cat === 'all' ? 'Todas las categorías' : cat}</option>
              ))}
            </Select>
          </div>
        </div>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-100">
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Producto</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Categoría</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Stock Actual</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Estado</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider text-right">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr><td colSpan={4} className="px-6 py-12 text-center"><Loader2 className="h-8 w-8 animate-spin mx-auto text-blue-600" /></td></tr>
              ) : filteredProducts.length === 0 ? (
                <tr><td colSpan={4} className="px-6 py-12 text-center text-gray-500">No se encontraron productos.</td></tr>
              ) : filteredProducts.map(p => (
                <tr key={p.id} className="hover:bg-gray-50/50 transition-colors group">
                  <td className="px-6 py-4">
                    <div className="flex items-center gap-3">
                      {p.imageUrl ? (
                        <img 
                          src={p.imageUrl} 
                          alt={p.name} 
                          className="w-10 h-10 rounded-lg object-cover border border-gray-100"
                          referrerPolicy="no-referrer"
                        />
                      ) : (
                        <div className="w-10 h-10 rounded-lg flex items-center justify-center font-bold bg-blue-100 text-blue-600">
                          {p.name.charAt(0)}
                        </div>
                      )}
                      <div>
                        <p className="font-semibold text-gray-900">{p.name}</p>
                        <p className="text-xs text-gray-500">{p.unit}</p>
                      </div>
                    </div>
                  </td>
                  <td className="px-6 py-4">
                    <span className="px-2.5 py-1 rounded-full bg-gray-100 text-gray-600 text-xs font-medium">
                      {p.category}
                    </span>
                  </td>
                  <td className="px-6 py-4">
                    <div className="flex items-center gap-2">
                      <span className="font-bold text-lg text-gray-900">
                        {p.quantity}
                      </span>
                      <span className="text-xs text-gray-400 uppercase">{p.unit}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4">
                    <span className={cn(
                      "px-2.5 py-1 rounded-full text-xs font-medium",
                      p.status === 'available' ? "bg-green-100 text-green-600" : "bg-yellow-100 text-yellow-600"
                    )}>
                      {p.status === 'available' ? 'Disponible' : 'Pedido'}
                    </span>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <div className="flex items-center justify-end gap-2">
                      {p.useQR && (
                        <Button 
                          variant="ghost" 
                          size="sm" 
                          className="text-purple-600 hover:bg-purple-50"
                          onClick={() => { setQrProduct(p); setIsQrModalOpen(true); }}
                        >
                          <QrCode className="h-4 w-4" />
                        </Button>
                      )}
                      <Button 
                        variant="ghost" 
                        size="sm" 
                        className="text-blue-600 hover:bg-blue-50"
                        onClick={() => { setSelectedProduct(p); setIsMovementModalOpen(true); }}
                      >
                        Stock
                      </Button>
                      <Button 
                        variant="ghost" 
                        size="sm" 
                        onClick={() => { setEditingProduct(p); setIsModalOpen(true); }}
                      >
                        <Edit2 className="h-4 w-4" />
                      </Button>
                      <Button 
                        variant="ghost" 
                        size="sm" 
                        className="text-red-600 hover:bg-red-50"
                        onClick={() => handleDeleteProduct(p.id)}
                      >
                        <Trash2 className="h-4 w-4" />
                      </Button>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      {/* Product Modal */}
      <Modal 
        isOpen={isModalOpen} 
        onClose={() => { setIsModalOpen(false); setEditingProduct(null); setImagePreview(null); }}
        title={editingProduct ? 'Editar Producto' : 'Nuevo Producto'}
      >
        <form onSubmit={handleSaveProduct} className="space-y-4">
          <div className="flex justify-center mb-4">
            <div className="relative group">
              <div className="w-24 h-24 rounded-2xl bg-gray-100 border-2 border-dashed border-gray-300 flex items-center justify-center overflow-hidden">
                {imagePreview || editingProduct?.imageUrl ? (
                  <img 
                    src={imagePreview || editingProduct?.imageUrl} 
                    alt="Preview" 
                    className="w-full h-full object-cover"
                    referrerPolicy="no-referrer"
                  />
                ) : (
                  <ImageIcon className="h-8 w-8 text-gray-400" />
                )}
              </div>
              <label className="absolute inset-0 flex items-center justify-center bg-black/40 opacity-0 group-hover:opacity-100 transition-opacity rounded-2xl cursor-pointer">
                <Upload className="h-6 w-6 text-white" />
                <input type="file" className="hidden" accept="image/*" onChange={handleImageChange} />
              </label>
            </div>
          </div>
          <Input label="Nombre del Material" name="name" defaultValue={editingProduct?.name} placeholder="Ej: Cemento, Varilla..." required />
          <Input label="Categoría" name="category" defaultValue={editingProduct?.category} placeholder="Ej: Obra Gris, Acabados..." required />
          <div className="grid grid-cols-2 gap-4">
            <Input label="Stock Inicial" name="quantity" type="number" defaultValue={editingProduct?.quantity || 0} required />
            <Select label="Unidad" name="unit" defaultValue={editingProduct?.unit || 'kg'}>
              <option value="kg">Kilogramos (kg)</option>
              <option value="unidades">Unidades</option>
              <option value="m2">Metros Cuadrados (m2)</option>
              <option value="m3">Metros Cúbicos (m3)</option>
              <option value="sacos">Sacos / Bultos</option>
            </Select>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <Select label="Estado" name="status" defaultValue={editingProduct?.status || 'available'}>
              <option value="available">Disponible</option>
              <option value="ordered">Pedido</option>
            </Select>
            <div className="flex items-center gap-2 pt-8">
              <input 
                type="checkbox" 
                id="useQR" 
                name="useQR" 
                defaultChecked={editingProduct?.useQR || false}
                className="h-4 w-4 rounded border-gray-300 text-blue-600 focus:ring-blue-500"
              />
              <label htmlFor="useQR" className="text-sm font-medium text-gray-700">Activar QR</label>
            </div>
          </div>
          <div className="flex gap-3 pt-4">
            <Button type="button" variant="outline" className="flex-1" onClick={() => setIsModalOpen(false)}>Cancelar</Button>
            <Button type="submit" className="flex-1">Guardar</Button>
          </div>
        </form>
      </Modal>

      {/* Movement Modal */}
      <Modal 
        isOpen={isMovementModalOpen} 
        onClose={() => { setIsMovementModalOpen(false); setSelectedProduct(null); }}
        title={`Movimiento: ${selectedProduct?.name}`}
      >
        <form onSubmit={handleMovement} className="space-y-4">
          <div className="p-4 bg-gray-50 rounded-xl mb-4">
            <p className="text-sm text-gray-500">Stock Actual: <span className="font-bold text-gray-900">{selectedProduct?.quantity} {selectedProduct?.unit}</span></p>
          </div>
          <Select label="Tipo de Movimiento" name="type" required>
            <option value="entry">Entrada (+)</option>
            <option value="exit">Salida (-)</option>
          </Select>
          <Input label="Cantidad" name="amount" type="number" min="1" required />
          <div className="grid grid-cols-2 gap-4">
            <Input label="Entregado por (Admin)" name="givenBy" placeholder="Nombre del administrador" required />
            <div className="space-y-1.5 w-full">
              <label className="text-sm font-medium text-gray-700">Recibido por (Personal)</label>
              <select 
                name="receivedBy" 
                required
                className="flex h-10 w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
              >
                <option value="">Seleccionar persona...</option>
                {requesters.map(r => (
                  <option key={r.id} value={r.name}>{r.name}</option>
                ))}
                <option value="Otro">Otro (Especificar en notas)</option>
              </select>
            </div>
          </div>
          <Input label="Notas / Observaciones" name="notes" placeholder="Ej: Obra Calle 10, Reposición..." />
          <div className="flex gap-3 pt-4">
            <Button type="button" variant="outline" className="flex-1" onClick={() => setIsMovementModalOpen(false)}>Cancelar</Button>
            <Button type="submit" className="flex-1">Registrar</Button>
          </div>
        </form>
      </Modal>

      {/* QR Modal */}
      <Modal
        isOpen={isQrModalOpen}
        onClose={() => { setIsQrModalOpen(false); setQrProduct(null); }}
        title={`Código QR: ${qrProduct?.name}`}
      >
        {qrProduct && (
          <div className="flex flex-col items-center space-y-6 py-4">
            <div className="p-4 bg-white rounded-2xl shadow-inner border border-gray-100">
              <QRCodeCanvas 
                id="qr-canvas"
                value={JSON.stringify({ id: qrProduct.id, action: 'mark_as_ordered' })} 
                size={256}
                level="H"
                includeMargin={true}
              />
            </div>
            <div className="text-center">
              <p className="text-sm text-gray-500 mb-4">Escanea este código para marcar el producto como pedido.</p>
              <Button onClick={() => {
                const canvas = document.getElementById('qr-canvas') as HTMLCanvasElement;
                if (canvas) {
                  const url = canvas.toDataURL('image/png');
                  const link = document.createElement('a');
                  link.download = `QR_${qrProduct.name}.png`;
                  link.href = url;
                  link.click();
                }
              }} className="gap-2">
                <Download className="h-5 w-5" />
                Descargar QR
              </Button>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
};

const Movements = () => {
  const [movements, setMovements] = useState<Movement[]>([]);
  const [loading, setLoading] = useState(true);
  const [isExporting, setIsExporting] = useState(false);

  useEffect(() => {
    const unsub = onSnapshot(query(collection(db, 'movements'), orderBy('date', 'desc')), (snap) => {
      setMovements(snap.docs.map(d => ({ id: d.id, ...d.data() } as Movement)));
      setLoading(false);
    });
    return unsub;
  }, []);

  const exportMovementsToExcel = async () => {
    setIsExporting(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Historial de Movimientos');
      
      sheet.columns = [
        { header: 'Fecha', key: 'date', width: 15 },
        { header: 'Hora', key: 'time', width: 12 },
        { header: 'Producto', key: 'product', width: 30 },
        { header: 'Tipo', key: 'type', width: 12 },
        { header: 'Cantidad', key: 'quantity', width: 12 },
        { header: 'Entregado por', key: 'givenBy', width: 25 },
        { header: 'Recibido por', key: 'receivedBy', width: 25 },
        { header: 'Notas', key: 'notes', width: 40 },
      ];

      sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

      movements.forEach(m => {
        const dateObj = m.date?.toDate ? m.date.toDate() : new Date(m.date);
        const row = sheet.addRow({
          date: format(dateObj, 'dd/MM/yyyy'),
          time: format(dateObj, 'HH:mm:ss'),
          product: m.productName,
          type: m.type === 'entry' ? 'ENTRADA' : 'SALIDA',
          quantity: m.quantity,
          givenBy: m.givenBy || '-',
          receivedBy: m.receivedBy || '-',
          notes: m.notes || '-'
        });

        const typeCell = row.getCell('type');
        if (m.type === 'entry') {
          typeCell.font = { color: { argb: 'FF166534' }, bold: true };
        } else {
          typeCell.font = { color: { argb: 'FF991B1B' }, bold: true };
        }

        row.eachCell({ includeEmpty: true }, (cell) => {
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `Movimientos_Bodega_${format(new Date(), 'yyyy-MM-dd')}.xlsx`);
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Historial de Movimientos</h1>
          <p className="text-gray-500">Registro completo de entradas y salidas</p>
        </div>
        <Button 
          variant="outline" 
          onClick={exportMovementsToExcel} 
          disabled={isExporting} 
          className="gap-2"
        >
          {isExporting ? <Loader2 className="h-5 w-5 animate-spin" /> : <Download className="h-5 w-5" />}
          {isExporting ? 'Exportando...' : 'Exportar Movimientos'}
        </Button>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-100">
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Fecha</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Producto</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Tipo</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Cantidad</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Entrega / Recibe</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Notas</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr><td colSpan={6} className="px-6 py-12 text-center"><Loader2 className="h-8 w-8 animate-spin mx-auto text-blue-600" /></td></tr>
              ) : movements.length === 0 ? (
                <tr><td colSpan={6} className="px-6 py-12 text-center text-gray-500">No hay movimientos registrados.</td></tr>
              ) : movements.map(m => (
                <tr key={m.id} className="hover:bg-gray-50/50 transition-colors">
                  <td className="px-6 py-4 text-sm text-gray-500">
                    {m.date?.toDate ? format(m.date.toDate(), 'dd/MM/yyyy HH:mm') : '-'}
                  </td>
                  <td className="px-6 py-4">
                    <p className="font-semibold text-gray-900">{m.productName}</p>
                  </td>
                  <td className="px-6 py-4">
                    <span className={cn(
                      "px-2.5 py-1 rounded-full text-xs font-medium",
                      m.type === 'entry' ? "bg-green-100 text-green-600" : "bg-orange-100 text-orange-600"
                    )}>
                      {m.type === 'entry' ? 'Entrada' : 'Salida'}
                    </span>
                  </td>
                  <td className="px-6 py-4 font-bold text-gray-900">
                    {m.type === 'entry' ? '+' : '-'}{m.quantity}
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-500">
                    <div className="flex flex-col">
                      <span className="font-medium text-gray-700">E: {m.givenBy || '-'}</span>
                      <span className="text-gray-400">R: {m.receivedBy || '-'}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-500 italic">{m.notes || '-'}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};

const AttendancePage = () => {
  const [currentTime, setCurrentTime] = useState(new Date());
  const [requesters, setRequesters] = useState<Requester[]>([]);
  const [selectedWorkerId, setSelectedWorkerId] = useState<string>('');
  const [attendanceRecords, setAttendanceRecords] = useState<Attendance[]>([]);
  const [loading, setLoading] = useState(true);
  const [isExporting, setIsExporting] = useState(false);
  const [isEditModalOpen, setIsEditModalOpen] = useState(false);
  const [editingRecord, setEditingRecord] = useState<Attendance | null>(null);
  const [editForm, setEditForm] = useState({
    arrival: '',
    breakStart: '',
    breakEnd: '',
    departure: ''
  });

  useEffect(() => {
    const timer = setInterval(() => setCurrentTime(new Date()), 1000);
    return () => clearInterval(timer);
  }, []);

  useEffect(() => {
    const unsubRequesters = onSnapshot(query(collection(db, 'requesters'), orderBy('name')), (snap) => {
      setRequesters(snap.docs.map(d => ({ id: d.id, ...d.data() } as Requester)));
    });
    const unsubAttendance = onSnapshot(query(collection(db, 'attendance'), orderBy('date', 'desc')), (snap) => {
      setAttendanceRecords(snap.docs.map(d => ({ id: d.id, ...d.data() } as Attendance)));
      setLoading(false);
    });
    return () => {
      unsubRequesters();
      unsubAttendance();
    };
  }, []);

  const selectedWorker = requesters.find(r => r.id === selectedWorkerId);
  const todayStr = format(new Date(), 'yyyy-MM-dd');
  const todayRecord = attendanceRecords.find(r => r.userId === selectedWorkerId && r.date === todayStr);

  const currentStatus = todayRecord?.status || 'not_started';

  const handleAttendance = async (type: 'arrival' | 'breakStart' | 'breakEnd' | 'departure') => {
    if (!selectedWorkerId) {
      alert('Por favor selecciona un trabajador');
      return;
    }

    const now = new Date();
    const updateData: any = {
      updatedAt: serverTimestamp(),
    };

    if (type === 'arrival') {
      if (todayRecord) {
        alert('Ya has marcado llegada hoy');
        return;
      }
      const newRecord: Omit<Attendance, 'id'> = {
        userId: selectedWorkerId,
        userName: selectedWorker?.name || 'Desconocido',
        date: todayStr,
        arrival: now,
        status: 'working',
      };
      await addDoc(collection(db, 'attendance'), newRecord);
    } else {
      if (!todayRecord) {
        alert('Debes marcar llegada primero');
        return;
      }

      if (type === 'breakStart') {
        if (todayRecord.status !== 'working') {
          alert('Solo puedes iniciar receso si estás en jornada');
          return;
        }
        updateData.breakStart = now;
        updateData.status = 'on_break';
      } else if (type === 'breakEnd') {
        if (todayRecord.status !== 'on_break') {
          alert('Solo puedes terminar receso si estás en receso');
          return;
        }
        updateData.breakEnd = now;
        updateData.status = 'working';
      } else if (type === 'departure') {
        if (todayRecord.status === 'finished') {
          alert('Ya has marcado salida hoy');
          return;
        }
        if (todayRecord.status === 'on_break') {
          alert('Debes terminar el receso antes de marcar salida');
          return;
        }
        updateData.departure = now;
        updateData.status = 'finished';

        // Calculate hours
        const arrival = todayRecord.arrival.toDate ? todayRecord.arrival.toDate() : new Date(todayRecord.arrival);
        const breakStart = todayRecord.breakStart?.toDate ? todayRecord.breakStart.toDate() : (todayRecord.breakStart ? new Date(todayRecord.breakStart) : null);
        const breakEnd = todayRecord.breakEnd?.toDate ? todayRecord.breakEnd.toDate() : (todayRecord.breakEnd ? new Date(todayRecord.breakEnd) : null);
        
        const totalMins = differenceInMinutes(now, arrival);
        let breakMins = 0;
        if (breakStart && breakEnd) {
          breakMins = differenceInMinutes(breakEnd, breakStart);
        }
        updateData.totalHours = Number(((totalMins - breakMins) / 60).toFixed(2));
      }

      await updateDoc(doc(db, 'attendance', todayRecord.id), updateData);
    }
  };

  const calculateMonthlyHours = (workerId: string) => {
    const start = startOfMonth(new Date());
    const end = endOfMonth(new Date());
    return attendanceRecords
      .filter(r => r.userId === workerId && r.totalHours && isWithinInterval(r.arrival.toDate ? r.arrival.toDate() : new Date(r.arrival), { start, end }))
      .reduce((acc, r) => acc + (r.totalHours || 0), 0);
  };

  const exportAttendanceToExcel = async () => {
    setIsExporting(true);
    try {
      const workbook = new ExcelJS.Workbook();
      const sheet = workbook.addWorksheet('Control de Asistencia');
      
      sheet.columns = [
        { header: 'Fecha', key: 'date', width: 15 },
        { header: 'Trabajador', key: 'worker', width: 25 },
        { header: 'Llegada', key: 'arrival', width: 12 },
        { header: 'Inicio Receso', key: 'breakStart', width: 15 },
        { header: 'Fin Receso', key: 'breakEnd', width: 15 },
        { header: 'Salida', key: 'departure', width: 12 },
        { header: 'Horas Trabajadas', key: 'hours', width: 18 },
      ];

      sheet.getRow(1).font = { bold: true, color: { argb: 'FFFFFFFF' } };
      sheet.getRow(1).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF1E293B' } };
      sheet.getRow(1).alignment = { vertical: 'middle', horizontal: 'center' };

      attendanceRecords.forEach(r => {
        const formatTime = (ts: any) => ts ? format(ts.toDate ? ts.toDate() : new Date(ts), 'HH:mm:ss') : '-';
        sheet.addRow({
          date: r.date,
          worker: r.userName,
          arrival: formatTime(r.arrival),
          breakStart: formatTime(r.breakStart),
          breakEnd: formatTime(r.breakEnd),
          departure: formatTime(r.departure),
          hours: r.totalHours || 0
        });
      });

      const buffer = await workbook.xlsx.writeBuffer();
      saveAs(new Blob([buffer]), `Asistencia_Bodega_${format(new Date(), 'yyyy-MM-dd')}.xlsx`);
    } finally {
      setIsExporting(false);
    }
  };

  const openEditModal = (record: Attendance) => {
    setEditingRecord(record);
    const formatTimeForInput = (ts: any) => ts ? format(ts.toDate ? ts.toDate() : new Date(ts), 'HH:mm') : '';
    setEditForm({
      arrival: formatTimeForInput(record.arrival),
      breakStart: formatTimeForInput(record.breakStart),
      breakEnd: formatTimeForInput(record.breakEnd),
      departure: formatTimeForInput(record.departure)
    });
    setIsEditModalOpen(true);
  };

  const handleUpdateAttendance = async () => {
    if (!editingRecord) return;

    try {
      const updateData: any = {
        updatedAt: serverTimestamp()
      };

      const parseTime = (timeStr: string, baseDate: string) => {
        if (!timeStr) return null;
        return new Date(`${baseDate}T${timeStr}:00`);
      };

      const arrival = parseTime(editForm.arrival, editingRecord.date);
      const breakStart = parseTime(editForm.breakStart, editingRecord.date);
      const breakEnd = parseTime(editForm.breakEnd, editingRecord.date);
      const departure = parseTime(editForm.departure, editingRecord.date);

      if (arrival) updateData.arrival = arrival;
      updateData.breakStart = breakStart;
      updateData.breakEnd = breakEnd;
      updateData.departure = departure;

      // Recalculate hours
      if (arrival && departure) {
        const totalMins = differenceInMinutes(departure, arrival);
        let breakMins = 0;
        if (breakStart && breakEnd) {
          breakMins = differenceInMinutes(breakEnd, breakStart);
        }
        updateData.totalHours = Number(((totalMins - breakMins) / 60).toFixed(2));
      } else {
        updateData.totalHours = 0;
      }

      await updateDoc(doc(db, 'attendance', editingRecord.id), updateData);
      setIsEditModalOpen(false);
      setEditingRecord(null);
    } catch (error) {
      console.error('Error updating attendance:', error);
      alert('Error al actualizar el registro');
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Control de Asistencia</h1>
          <p className="text-gray-500">Registra y visualiza la jornada laboral</p>
        </div>
        <div className="flex items-center gap-4 bg-white px-6 py-3 rounded-2xl shadow-sm border border-gray-100">
          <Clock className="h-6 w-6 text-blue-600 animate-pulse" />
          <div className="flex flex-col">
            <span className="text-2xl font-mono font-bold text-gray-900">{format(currentTime, 'HH:mm:ss')}</span>
            <span className="text-xs text-gray-400 uppercase tracking-widest">{format(currentTime, 'EEEE, d MMMM')}</span>
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        {/* Registration Card */}
        <div className="lg:col-span-2 space-y-6">
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-8">
            <div className="max-w-md mx-auto space-y-8">
              <div className="space-y-2">
                <label className="text-sm font-medium text-gray-700">Seleccionar Trabajador</label>
                <select 
                  value={selectedWorkerId}
                  onChange={(e) => setSelectedWorkerId(e.target.value)}
                  className="flex h-12 w-full rounded-xl border border-gray-300 bg-white px-4 py-2 text-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
                >
                  <option value="">Seleccionar...</option>
                  {requesters.map(r => (
                    <option key={r.id} value={r.id}>{r.name}</option>
                  ))}
                </select>
              </div>

              {selectedWorkerId && (
                <div className="space-y-6">
                  <div className="flex items-center justify-between p-4 bg-blue-50 rounded-2xl border border-blue-100">
                    <div className="flex items-center gap-3">
                      <div className="w-12 h-12 rounded-full bg-blue-600 text-white flex items-center justify-center text-xl font-bold">
                        {selectedWorker?.name.charAt(0)}
                      </div>
                      <div>
                        <p className="font-bold text-gray-900">{selectedWorker?.name}</p>
                        <p className={cn(
                          "text-xs font-bold uppercase tracking-wider",
                          currentStatus === 'working' ? "text-green-600" :
                          currentStatus === 'on_break' ? "text-orange-600" :
                          currentStatus === 'finished' ? "text-gray-400" : "text-blue-600"
                        )}>
                          {currentStatus === 'working' ? '● En Jornada' :
                           currentStatus === 'on_break' ? '● En Receso' :
                           currentStatus === 'finished' ? '● Finalizado' : '● No Iniciado'}
                        </p>
                      </div>
                    </div>
                    <div className="text-right">
                      <p className="text-xs text-gray-400 uppercase font-bold">Mes Actual</p>
                      <p className="text-2xl font-bold text-blue-600">{calculateMonthlyHours(selectedWorkerId)}h</p>
                    </div>
                  </div>

                  <div className="grid grid-cols-2 gap-4">
                    <Button 
                      size="lg" 
                      className="h-24 flex-col gap-2 rounded-2xl"
                      disabled={currentStatus !== 'not_started'}
                      onClick={() => handleAttendance('arrival')}
                    >
                      <LogIn className="h-8 w-8" />
                      Llegada
                    </Button>
                    <Button 
                      size="lg" 
                      variant="secondary"
                      className="h-24 flex-col gap-2 rounded-2xl bg-orange-50 text-orange-600 hover:bg-orange-100 border-orange-100"
                      disabled={currentStatus !== 'working'}
                      onClick={() => handleAttendance('breakStart')}
                    >
                      <Coffee className="h-8 w-8" />
                      Inicio Receso
                    </Button>
                    <Button 
                      size="lg" 
                      variant="secondary"
                      className="h-24 flex-col gap-2 rounded-2xl bg-blue-50 text-blue-600 hover:bg-blue-100 border-blue-100"
                      disabled={currentStatus !== 'on_break'}
                      onClick={() => handleAttendance('breakEnd')}
                    >
                      <RotateCcw className="h-8 w-8" />
                      Fin Receso
                    </Button>
                    <Button 
                      size="lg" 
                      variant="danger"
                      className="h-24 flex-col gap-2 rounded-2xl"
                      disabled={currentStatus !== 'working'}
                      onClick={() => handleAttendance('departure')}
                    >
                      <LogOut className="h-8 w-8" />
                      Salida
                    </Button>
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>

        {/* Summary Card */}
        <div className="space-y-6">
          <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
            <div className="flex items-center justify-between mb-6">
              <h3 className="font-bold text-gray-900 flex items-center gap-2">
                <Calendar className="h-5 w-5 text-blue-600" />
                Resumen del Mes
              </h3>
              <Button variant="ghost" size="sm" onClick={exportAttendanceToExcel} disabled={isExporting}>
                {isExporting ? <Loader2 className="h-4 w-4 animate-spin" /> : <Download className="h-4 w-4" />}
              </Button>
            </div>
            <div className="space-y-4">
              {requesters.map(worker => {
                const hours = calculateMonthlyHours(worker.id);
                if (hours === 0) return null;
                return (
                  <div key={worker.id} className="flex items-center justify-between p-3 rounded-xl bg-gray-50 border border-gray-100">
                    <span className="text-sm font-medium text-gray-700">{worker.name}</span>
                    <span className="font-bold text-blue-600">{hours}h</span>
                  </div>
                );
              })}
              {requesters.every(w => calculateMonthlyHours(w.id) === 0) && (
                <p className="text-sm text-gray-400 text-center italic py-4">Sin registros este mes</p>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* History Table */}
      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="px-6 py-4 border-b border-gray-50 flex items-center justify-between">
          <h3 className="font-bold text-gray-900">Registros Recientes</h3>
        </div>
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-100">
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Fecha</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Trabajador</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Llegada</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Receso</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Salida</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Total</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider text-right">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr><td colSpan={7} className="px-6 py-12 text-center"><Loader2 className="h-8 w-8 animate-spin mx-auto text-blue-600" /></td></tr>
              ) : attendanceRecords.length === 0 ? (
                <tr><td colSpan={7} className="px-6 py-12 text-center text-gray-500">No hay registros de asistencia.</td></tr>
              ) : attendanceRecords.map(r => (
                <tr key={r.id} className="hover:bg-gray-50/50 transition-colors">
                  <td className="px-6 py-4 text-sm font-medium text-gray-900">{r.date}</td>
                  <td className="px-6 py-4 text-sm text-gray-700">{r.userName}</td>
                  <td className="px-6 py-4 text-sm text-gray-500">
                    {r.arrival ? format(r.arrival.toDate ? r.arrival.toDate() : new Date(r.arrival), 'HH:mm') : '-'}
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-500">
                    {r.breakStart ? format(r.breakStart.toDate ? r.breakStart.toDate() : new Date(r.breakStart), 'HH:mm') : '-'}
                    {r.breakEnd ? ` → ${format(r.breakEnd.toDate ? r.breakEnd.toDate() : new Date(r.breakEnd), 'HH:mm')}` : ''}
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-500">
                    {r.departure ? format(r.departure.toDate ? r.departure.toDate() : new Date(r.departure), 'HH:mm') : '-'}
                  </td>
                  <td className="px-6 py-4">
                    <span className="font-bold text-blue-600">{r.totalHours || 0}h</span>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <Button variant="ghost" size="sm" onClick={() => openEditModal(r)}>
                      <Edit2 className="h-4 w-4 text-gray-400 hover:text-blue-600" />
                    </Button>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <Modal
        isOpen={isEditModalOpen}
        onClose={() => setIsEditModalOpen(false)}
        title={`Editar Horarios - ${editingRecord?.userName} (${editingRecord?.date})`}
      >
        <div className="space-y-4">
          <div className="grid grid-cols-2 gap-4">
            <Input
              label="Hora de Llegada"
              type="time"
              value={editForm.arrival}
              onChange={(e) => setEditForm({ ...editForm, arrival: e.target.value })}
            />
            <Input
              label="Hora de Salida"
              type="time"
              value={editForm.departure}
              onChange={(e) => setEditForm({ ...editForm, departure: e.target.value })}
            />
          </div>
          <div className="grid grid-cols-2 gap-4">
            <Input
              label="Inicio Receso"
              type="time"
              value={editForm.breakStart}
              onChange={(e) => setEditForm({ ...editForm, breakStart: e.target.value })}
            />
            <Input
              label="Fin Receso"
              type="time"
              value={editForm.breakEnd}
              onChange={(e) => setEditForm({ ...editForm, breakEnd: e.target.value })}
            />
          </div>
          <div className="flex justify-end gap-3 pt-4">
            <Button variant="ghost" onClick={() => setIsEditModalOpen(false)}>
              Cancelar
            </Button>
            <Button onClick={handleUpdateAttendance}>
              Guardar Cambios
            </Button>
          </div>
        </div>
      </Modal>
    </div>
  );
};

const Requesters = () => {
  const [requesters, setRequesters] = useState<Requester[]>([]);
  const [loading, setLoading] = useState(true);
  const [newName, setNewName] = useState('');

  useEffect(() => {
    const unsub = onSnapshot(query(collection(db, 'requesters'), orderBy('name')), (snap) => {
      setRequesters(snap.docs.map(d => ({ id: d.id, ...d.data() } as Requester)));
      setLoading(false);
    });
    return unsub;
  }, []);

  const handleAddRequester = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!newName.trim()) return;
    try {
      await addDoc(collection(db, 'requesters'), {
        name: newName.trim(),
        createdAt: serverTimestamp()
      });
      setNewName('');
    } catch (err) {
      console.error(err);
      alert('Error al añadir personal');
    }
  };

  const handleDeleteRequester = async (id: string) => {
    if (window.confirm('¿Estás seguro de eliminar a esta persona de la lista autorizada?')) {
      await deleteDoc(doc(db, 'requesters', id));
    }
  };

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-3xl font-bold text-gray-900">Personal Autorizado</h1>
        <p className="text-gray-500">Gestiona la lista de personas que pueden retirar materiales</p>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 p-6">
        <form onSubmit={handleAddRequester} className="flex gap-3 mb-8">
          <Input 
            placeholder="Nombre completo de la persona..." 
            value={newName}
            onChange={(e) => setNewName(e.target.value)}
            className="flex-1"
          />
          <Button type="submit" className="gap-2 shrink-0">
            <Plus className="h-5 w-5" />
            Añadir a la Lista
          </Button>
        </form>

        <div className="space-y-3">
          {loading ? (
            <div className="flex justify-center py-8"><Loader2 className="h-8 w-8 animate-spin text-blue-600" /></div>
          ) : requesters.length === 0 ? (
            <p className="text-gray-500 text-center py-8 italic">No hay personal registrado aún.</p>
          ) : (
            requesters.map(r => (
              <div key={r.id} className="flex items-center justify-between p-4 rounded-xl bg-gray-50 border border-gray-100 group hover:bg-white hover:shadow-sm transition-all">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-full bg-blue-100 text-blue-600 flex items-center justify-center font-bold">
                    {r.name.charAt(0)}
                  </div>
                  <span className="font-semibold text-gray-900">{r.name}</span>
                </div>
                <Button 
                  variant="ghost" 
                  size="sm" 
                  className="text-red-600 opacity-0 group-hover:opacity-100 transition-opacity"
                  onClick={() => handleDeleteRequester(r.id)}
                >
                  <Trash2 className="h-4 w-4" />
                </Button>
              </div>
            ))
          )}
        </div>
      </div>
    </div>
  );
};

const RequestsPage = () => {
  const [requests, setRequests] = useState<ProductRequest[]>([]);
  const [requesters, setRequesters] = useState<Requester[]>([]);
  const [loading, setLoading] = useState(true);
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);

  const [formData, setFormData] = useState({
    productName: '',
    category: '',
    quantity: '',
    unit: '',
    priority: 'medium' as 'low' | 'medium' | 'high',
    comment: '',
    userId: ''
  });

  useEffect(() => {
    const unsubRequests = onSnapshot(query(collection(db, 'requests'), orderBy('createdAt', 'desc')), (snap) => {
      setRequests(snap.docs.map(d => ({ id: d.id, ...d.data() } as ProductRequest)));
      setLoading(false);
    });
    const unsubRequesters = onSnapshot(query(collection(db, 'requesters'), orderBy('name')), (snap) => {
      setRequesters(snap.docs.map(d => ({ id: d.id, ...d.data() } as Requester)));
    });
    return () => {
      unsubRequests();
      unsubRequesters();
    };
  }, []);

  const handleAddRequest = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.productName || !formData.quantity || !formData.userId) {
      alert('Por favor completa los campos obligatorios');
      return;
    }

    const selectedWorker = requesters.find(r => r.id === formData.userId);

    try {
      await addDoc(collection(db, 'requests'), {
        productName: formData.productName,
        category: formData.category || 'Sin Categoría',
        quantity: Number(formData.quantity),
        unit: formData.unit || 'unidades',
        priority: formData.priority,
        comment: formData.comment,
        status: 'pending',
        userId: formData.userId,
        userName: selectedWorker?.name || 'Desconocido',
        createdAt: serverTimestamp()
      });
      setIsAddModalOpen(false);
      setFormData({
        productName: '',
        category: '',
        quantity: '',
        unit: '',
        priority: 'medium',
        comment: '',
        userId: ''
      });
    } catch (err) {
      console.error(err);
      alert('Error al crear la solicitud');
    }
  };

  const updateStatus = async (request: ProductRequest, newStatus: ProductRequest['status']) => {
    try {
      const updateData: any = {
        status: newStatus,
        updatedAt: serverTimestamp()
      };

      if (newStatus === 'purchased' && request.status !== 'purchased') {
        const confirmAdd = window.confirm('¿Deseas agregar este producto al inventario automáticamente?');
        if (confirmAdd) {
          await addDoc(collection(db, 'products'), {
            name: request.productName,
            category: request.category,
            quantity: request.quantity,
            unit: request.unit,
            entryDate: serverTimestamp(),
            updatedAt: serverTimestamp()
          });

          await addDoc(collection(db, 'movements'), {
            productId: 'new',
            productName: request.productName,
            type: 'entry',
            quantity: request.quantity,
            date: serverTimestamp(),
            userId: request.userId,
            userName: request.userName,
            notes: `Auto-agregado desde solicitud: ${request.comment || ''}`
          });
        }
      }

      await updateDoc(doc(db, 'requests', request.id), updateData);
    } catch (err) {
      console.error(err);
      alert('Error al actualizar el estado');
    }
  };

  const getPriorityColor = (priority: string) => {
    switch (priority) {
      case 'high': return 'text-red-600 bg-red-50 border-red-100';
      case 'medium': return 'text-yellow-600 bg-yellow-50 border-yellow-100';
      case 'low': return 'text-green-600 bg-green-50 border-green-100';
      default: return 'text-gray-600 bg-gray-50 border-gray-100';
    }
  };

  const getStatusIcon = (status: string) => {
    switch (status) {
      case 'pending': return '🟡';
      case 'in_process': return '🔵';
      case 'purchased': return '🟢';
      case 'rejected': return '❌';
      default: return '⚪';
    }
  };

  const getStatusLabel = (status: string) => {
    switch (status) {
      case 'pending': return 'Pendiente';
      case 'in_process': return 'En Proceso';
      case 'purchased': return 'Comprado';
      case 'rejected': return 'Rechazado';
      default: return status;
    }
  };

  return (
    <div className="space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Solicitudes de Compra</h1>
          <p className="text-gray-500">Gestiona las necesidades de materiales y herramientas</p>
        </div>
        <Button onClick={() => setIsAddModalOpen(true)} className="gap-2">
          <Plus className="h-5 w-5" />
          Nueva Solicitud
        </Button>
      </div>

      <div className="bg-white rounded-2xl shadow-sm border border-gray-100 overflow-hidden">
        <div className="overflow-x-auto">
          <table className="w-full text-left">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-100">
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Producto</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Cantidad</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Prioridad</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Solicitante</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Fecha</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider">Estado</th>
                <th className="px-6 py-4 text-xs font-bold text-gray-500 uppercase tracking-wider text-right">Acciones</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-50">
              {loading ? (
                <tr><td colSpan={7} className="px-6 py-12 text-center"><Loader2 className="h-8 w-8 animate-spin mx-auto text-blue-600" /></td></tr>
              ) : requests.length === 0 ? (
                <tr><td colSpan={7} className="px-6 py-12 text-center text-gray-500">No hay solicitudes pendientes.</td></tr>
              ) : requests.map(req => (
                <tr key={req.id} className="hover:bg-gray-50/50 transition-colors">
                  <td className="px-6 py-4">
                    <div className="flex flex-col">
                      <span className="font-bold text-gray-900">{req.productName}</span>
                      <span className="text-xs text-gray-400">{req.category}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-700">
                    {req.quantity} {req.unit}
                  </td>
                  <td className="px-6 py-4">
                    <span className={cn(
                      "px-2.5 py-1 rounded-full text-xs font-bold border uppercase tracking-wider",
                      getPriorityColor(req.priority)
                    )}>
                      {req.priority === 'high' ? 'Alta' : req.priority === 'medium' ? 'Media' : 'Baja'}
                    </span>
                  </td>
                  <td className="px-6 py-4 text-sm text-gray-700">{req.userName}</td>
                  <td className="px-6 py-4 text-sm text-gray-500">
                    {req.createdAt ? format(req.createdAt.toDate ? req.createdAt.toDate() : new Date(req.createdAt), 'dd/MM/yyyy') : '-'}
                  </td>
                  <td className="px-6 py-4">
                    <div className="flex items-center gap-2">
                      <span>{getStatusIcon(req.status)}</span>
                      <span className="text-sm font-medium text-gray-700">{getStatusLabel(req.status)}</span>
                    </div>
                  </td>
                  <td className="px-6 py-4 text-right">
                    <div className="flex items-center justify-end gap-2">
                      <select 
                        value={req.status}
                        onChange={(e) => updateStatus(req, e.target.value as any)}
                        className="text-xs border rounded p-1 focus:outline-none focus:ring-1 focus:ring-blue-500"
                      >
                        <option value="pending">Pendiente</option>
                        <option value="in_process">En Proceso</option>
                        <option value="purchased">Comprado</option>
                        <option value="rejected">Rechazado</option>
                      </select>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>

      <Modal
        isOpen={isAddModalOpen}
        onClose={() => setIsAddModalOpen(false)}
        title="Nueva Solicitud de Producto"
      >
        <form onSubmit={handleAddRequest} className="space-y-4">
          <Input 
            label="Nombre del Producto"
            placeholder="Ej: Cemento Portland"
            value={formData.productName}
            onChange={(e) => setFormData({...formData, productName: e.target.value})}
            required
          />
          <div className="grid grid-cols-2 gap-4">
            <Input 
              label="Categoría"
              placeholder="Ej: Construcción"
              value={formData.category}
              onChange={(e) => setFormData({...formData, category: e.target.value})}
            />
            <Select
              label="Prioridad"
              value={formData.priority}
              onChange={(e) => setFormData({...formData, priority: e.target.value as any})}
            >
              <option value="low">Baja (Verde)</option>
              <option value="medium">Media (Amarillo)</option>
              <option value="high">Alta (Rojo)</option>
            </Select>
          </div>
          <div className="grid grid-cols-2 gap-4">
            <Input 
              label="Cantidad"
              type="number"
              value={formData.quantity}
              onChange={(e) => setFormData({...formData, quantity: e.target.value})}
              required
            />
            <Input 
              label="Unidad"
              placeholder="Ej: sacos, kg, m"
              value={formData.unit}
              onChange={(e) => setFormData({...formData, unit: e.target.value})}
            />
          </div>
          <Select
            label="Solicitado por"
            value={formData.userId}
            onChange={(e) => setFormData({...formData, userId: e.target.value})}
            required
          >
            <option value="">Seleccionar trabajador...</option>
            {requesters.map(r => (
              <option key={r.id} value={r.id}>{r.name}</option>
            ))}
          </Select>
          <div className="space-y-1.5">
            <label className="text-sm font-medium text-gray-700">Comentario (Opcional)</label>
            <textarea 
              className="flex min-h-[80px] w-full rounded-lg border border-gray-300 bg-white px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent transition-all"
              value={formData.comment}
              onChange={(e) => setFormData({...formData, comment: e.target.value})}
            />
          </div>
          <div className="flex justify-end gap-3 pt-4">
            <Button type="button" variant="ghost" onClick={() => setIsAddModalOpen(false)}>
              Cancelar
            </Button>
            <Button type="submit">
              Crear Solicitud
            </Button>
          </div>
        </form>
      </Modal>
    </div>
  );
};

const DailyReportsPage = () => {
  const [reports, setReports] = useState<DailyReport[]>([]);
  const [requesters, setRequesters] = useState<Requester[]>([]);
  const [loading, setLoading] = useState(true);
  const [isAddModalOpen, setIsAddModalOpen] = useState(false);
  const [selectedReport, setSelectedReport] = useState<DailyReport | null>(null);

  const [formData, setFormData] = useState({
    userId: '',
    advances: '',
    issues: '',
    delays: '',
    nextDayPlan: ''
  });

  useEffect(() => {
    const unsubReports = onSnapshot(query(collection(db, 'dailyReports'), orderBy('createdAt', 'desc')), (snap) => {
      setReports(snap.docs.map(d => ({ id: d.id, ...d.data() } as DailyReport)));
      setLoading(false);
    });
    const unsubRequesters = onSnapshot(query(collection(db, 'requesters'), orderBy('name')), (snap) => {
      setRequesters(snap.docs.map(d => ({ id: d.id, ...d.data() } as Requester)));
    });
    return () => {
      unsubReports();
      unsubRequesters();
    };
  }, []);

  const handleAddReport = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!formData.userId || !formData.advances) {
      alert('Por favor selecciona un usuario y describe los avances');
      return;
    }

    const selectedWorker = requesters.find(r => r.id === formData.userId);

    try {
      await addDoc(collection(db, 'dailyReports'), {
        date: serverTimestamp(),
        userId: formData.userId,
        userName: selectedWorker?.name || 'Desconocido',
        advances: formData.advances,
        issues: formData.issues,
        delays: formData.delays,
        nextDayPlan: formData.nextDayPlan,
        createdAt: serverTimestamp()
      });
      setIsAddModalOpen(false);
      setFormData({
        userId: '',
        advances: '',
        issues: '',
        delays: '',
        nextDayPlan: ''
      });
    } catch (err) {
      console.error(err);
      alert('Error al crear el reporte');
    }
  };

  const exportToExcel = async () => {
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Bitácora de Obra');

    worksheet.columns = [
      { header: 'Fecha', key: 'date', width: 15 },
      { header: 'Usuario', key: 'userName', width: 20 },
      { header: 'Avances', key: 'advances', width: 40 },
      { header: 'Inconvenientes', key: 'issues', width: 30 },
      { header: 'Atrasos', key: 'delays', width: 30 },
      { header: 'Plan Siguiente Día', key: 'nextDayPlan', width: 30 },
    ];

    reports.forEach(report => {
      worksheet.addRow({
        date: report.createdAt ? format(report.createdAt.toDate ? report.createdAt.toDate() : new Date(report.createdAt), 'dd/MM/yyyy') : '-',
        userName: report.userName,
        advances: report.advances,
        issues: report.issues,
        delays: report.delays,
        nextDayPlan: report.nextDayPlan,
      });
    });

    const buffer = await workbook.xlsx.writeBuffer();
    saveAs(new Blob([buffer]), `Bitacora_Obra_${format(new Date(), 'dd-MM-yyyy')}.xlsx`);
  };

  return (
    <div className="space-y-8">
      <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <h1 className="text-3xl font-bold text-gray-900">Bitácora de Obra</h1>
          <p className="text-gray-500">Registro diario de avances, incidentes y planificación</p>
        </div>
        <div className="flex gap-3">
          <Button variant="outline" onClick={exportToExcel} className="gap-2">
            <Download className="h-5 w-5" />
            Exportar Excel
          </Button>
          <Button onClick={() => setIsAddModalOpen(true)} className="gap-2">
            <Plus className="h-5 w-5" />
            Nuevo Reporte
          </Button>
        </div>
      </div>

      {loading ? (
        <div className="flex justify-center py-12"><Loader2 className="h-12 w-12 animate-spin text-blue-600" /></div>
      ) : reports.length === 0 ? (
        <div className="bg-white rounded-2xl p-12 text-center border border-dashed border-gray-200">
          <BookOpen className="h-12 w-12 text-gray-300 mx-auto mb-4" />
          <p className="text-gray-500 italic">No hay reportes registrados aún.</p>
        </div>
      ) : (
        <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
          {reports.map(report => (
            <motion.div 
              key={report.id}
              whileHover={{ y: -4 }}
              className="bg-white rounded-2xl p-6 shadow-sm border border-gray-100 cursor-pointer hover:shadow-md transition-all"
              onClick={() => setSelectedReport(report)}
            >
              <div className="flex justify-between items-start mb-4">
                <div className="flex items-center gap-2 text-blue-600">
                  <Calendar className="h-4 w-4" />
                  <span className="text-sm font-bold">
                    {report.createdAt ? format(report.createdAt.toDate ? report.createdAt.toDate() : new Date(report.createdAt), 'dd/MM/yyyy') : '-'}
                  </span>
                </div>
                <div className="flex gap-1">
                  {report.issues && <AlertTriangle className="h-4 w-4 text-red-500" />}
                  {report.delays && <Clock className="h-4 w-4 text-yellow-500" />}
                  <CheckCircle2 className="h-4 w-4 text-green-500" />
                </div>
              </div>
              
              <div className="mb-4">
                <p className="text-xs text-gray-400 uppercase font-bold tracking-wider mb-1">Responsable</p>
                <p className="font-semibold text-gray-900">{report.userName}</p>
              </div>

              <div>
                <p className="text-xs text-gray-400 uppercase font-bold tracking-wider mb-1">Resumen de Avances</p>
                <p className="text-sm text-gray-600 line-clamp-3">{report.advances}</p>
              </div>

              <div className="mt-6 pt-4 border-t border-gray-50 flex items-center justify-between text-blue-600 text-xs font-bold">
                <span>VER DETALLE COMPLETO</span>
                <ChevronRight className="h-4 w-4" />
              </div>
            </motion.div>
          ))}
        </div>
      )}

      {/* Modal Nuevo Reporte */}
      <Modal
        isOpen={isAddModalOpen}
        onClose={() => setIsAddModalOpen(false)}
        title="Nuevo Reporte Diario"
      >
        <form onSubmit={handleAddReport} className="space-y-4">
          <Select
            label="Responsable del Reporte"
            value={formData.userId}
            onChange={(e) => setFormData({...formData, userId: e.target.value})}
            required
          >
            <option value="">Seleccionar trabajador...</option>
            {requesters.map(r => (
              <option key={r.id} value={r.id}>{r.name}</option>
            ))}
          </Select>

          <div className="space-y-1.5">
            <label className="text-sm font-medium text-green-700 flex items-center gap-2">
              <CheckCircle2 className="h-4 w-4" />
              Avances del Día
            </label>
            <textarea 
              className="flex min-h-[100px] w-full rounded-lg border border-green-200 bg-green-50/30 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-green-500 transition-all"
              placeholder="¿Qué se logró hoy?"
              value={formData.advances}
              onChange={(e) => setFormData({...formData, advances: e.target.value})}
              required
            />
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="space-y-1.5">
              <label className="text-sm font-medium text-red-700 flex items-center gap-2">
                <AlertTriangle className="h-4 w-4" />
                Inconvenientes
              </label>
              <textarea 
                className="flex min-h-[80px] w-full rounded-lg border border-red-200 bg-red-50/30 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-red-500 transition-all"
                placeholder="Problemas encontrados..."
                value={formData.issues}
                onChange={(e) => setFormData({...formData, issues: e.target.value})}
              />
            </div>
            <div className="space-y-1.5">
              <label className="text-sm font-medium text-yellow-700 flex items-center gap-2">
                <Clock className="h-4 w-4" />
                Atrasos
              </label>
              <textarea 
                className="flex min-h-[80px] w-full rounded-lg border border-yellow-200 bg-yellow-50/30 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-yellow-500 transition-all"
                placeholder="Causas de demora..."
                value={formData.delays}
                onChange={(e) => setFormData({...formData, delays: e.target.value})}
              />
            </div>
          </div>

          <div className="space-y-1.5">
            <label className="text-sm font-medium text-blue-700 flex items-center gap-2">
              <Calendar className="h-4 w-4" />
              Plan para Mañana
            </label>
            <textarea 
              className="flex min-h-[80px] w-full rounded-lg border border-blue-200 bg-blue-50/30 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 transition-all"
              placeholder="¿Qué se planea hacer mañana?"
              value={formData.nextDayPlan}
              onChange={(e) => setFormData({...formData, nextDayPlan: e.target.value})}
            />
          </div>

          <div className="flex justify-end gap-3 pt-4">
            <Button type="button" variant="ghost" onClick={() => setIsAddModalOpen(false)}>
              Cancelar
            </Button>
            <Button type="submit">
              Guardar Bitácora
            </Button>
          </div>
        </form>
      </Modal>

      {/* Modal Detalle Reporte */}
      <Modal
        isOpen={!!selectedReport}
        onClose={() => setSelectedReport(null)}
        title={`Detalle de Bitácora - ${selectedReport?.createdAt ? format(selectedReport.createdAt.toDate ? selectedReport.createdAt.toDate() : new Date(selectedReport.createdAt), 'dd/MM/yyyy') : ''}`}
      >
        {selectedReport && (
          <div className="space-y-6">
            <div className="flex items-center justify-between p-4 bg-gray-50 rounded-xl border border-gray-100">
              <div>
                <p className="text-xs text-gray-400 uppercase font-bold tracking-wider mb-1">Responsable</p>
                <p className="font-bold text-gray-900">{selectedReport.userName}</p>
              </div>
              <div className="text-right">
                <p className="text-xs text-gray-400 uppercase font-bold tracking-wider mb-1">Fecha de Registro</p>
                <p className="text-sm text-gray-600">
                  {selectedReport.createdAt ? format(selectedReport.createdAt.toDate ? selectedReport.createdAt.toDate() : new Date(selectedReport.createdAt), 'dd/MM/yyyy HH:mm') : '-'}
                </p>
              </div>
            </div>

            <div className="space-y-4">
              <div className="p-4 rounded-xl bg-green-50 border border-green-100">
                <h4 className="text-sm font-bold text-green-800 flex items-center gap-2 mb-2">
                  <CheckCircle2 className="h-4 w-4" />
                  AVANCES DEL DÍA
                </h4>
                <p className="text-sm text-green-900 whitespace-pre-wrap">{selectedReport.advances}</p>
              </div>

              {selectedReport.issues && (
                <div className="p-4 rounded-xl bg-red-50 border border-red-100">
                  <h4 className="text-sm font-bold text-red-800 flex items-center gap-2 mb-2">
                    <AlertTriangle className="h-4 w-4" />
                    INCONVENIENTES / PROBLEMAS
                  </h4>
                  <p className="text-sm text-red-900 whitespace-pre-wrap">{selectedReport.issues}</p>
                </div>
              )}

              {selectedReport.delays && (
                <div className="p-4 rounded-xl bg-yellow-50 border border-yellow-100">
                  <h4 className="text-sm font-bold text-yellow-800 flex items-center gap-2 mb-2">
                    <Clock className="h-4 w-4" />
                    ATRASOS
                  </h4>
                  <p className="text-sm text-yellow-900 whitespace-pre-wrap">{selectedReport.delays}</p>
                </div>
              )}

              <div className="p-4 rounded-xl bg-blue-50 border border-blue-100">
                <h4 className="text-sm font-bold text-blue-800 flex items-center gap-2 mb-2">
                  <Calendar className="h-4 w-4" />
                  PLANIFICACIÓN PARA MAÑANA
                </h4>
                <p className="text-sm text-blue-900 whitespace-pre-wrap">{selectedReport.nextDayPlan || 'No especificado'}</p>
              </div>
            </div>

            <div className="flex justify-end pt-4">
              <Button onClick={() => setSelectedReport(null)}>
                Cerrar
              </Button>
            </div>
          </div>
        )}
      </Modal>
    </div>
  );
};

const ScannerPage = () => {
  const [scanResult, setScanResult] = useState<string | null>(null);
  const [isScanning, setIsScanning] = useState(true);
  const [status, setStatus] = useState<'idle' | 'scanning' | 'success' | 'error'>('idle');
  const [message, setMessage] = useState('');
  const navigate = useNavigate();

  useEffect(() => {
    let scanner: Html5QrcodeScanner | null = null;

    if (isScanning) {
      scanner = new Html5QrcodeScanner(
        "reader",
        { fps: 10, qrbox: { width: 250, height: 250 } },
        /* verbose= */ false
      );

      scanner.render(onScanSuccess, onScanFailure);
    }

    return () => {
      if (scanner) {
        scanner.clear().catch(error => console.error("Failed to clear scanner", error));
      }
    };
  }, [isScanning]);

  async function onScanSuccess(decodedText: string) {
    try {
      const data = JSON.parse(decodedText);
      if (data.id && data.action === 'mark_as_ordered') {
        setScanResult(data.id);
        setIsScanning(false);
        setStatus('scanning');
        setMessage('Procesando producto...');

        // Update product status in Firestore
        await updateDoc(doc(db, 'products', data.id), {
          status: 'ordered',
          updatedAt: serverTimestamp()
        });

        setStatus('success');
        setMessage('¡Producto marcado como pedido exitosamente!');
        
        // Redirect after a short delay
        setTimeout(() => {
          navigate('/inventory');
        }, 2000);
      } else {
        throw new Error('QR no válido para esta aplicación');
      }
    } catch (err) {
      console.error(err);
      setStatus('error');
      setMessage('Error: El código QR no es válido o el producto no existe.');
      setTimeout(() => {
        setStatus('idle');
        setIsScanning(true);
      }, 3000);
    }
  }

  function onScanFailure(error: any) {
    // console.warn(`Code scan error = ${error}`);
  }

  return (
    <div className="max-w-2xl mx-auto space-y-8">
      <div className="text-center">
        <h1 className="text-3xl font-bold text-gray-900">Escáner de Productos</h1>
        <p className="text-gray-500">Escanea el código QR para marcar un producto como pedido</p>
      </div>

      <div className="bg-white rounded-3xl shadow-xl overflow-hidden border border-gray-100">
        <div className="p-8">
          {status === 'idle' || status === 'error' ? (
            <div id="reader" className="overflow-hidden rounded-2xl border-2 border-dashed border-gray-200"></div>
          ) : (
            <div className="flex flex-col items-center justify-center py-12 space-y-4">
              {status === 'scanning' && <Loader2 className="h-16 w-16 animate-spin text-blue-600" />}
              {status === 'success' && (
                <div className="bg-green-100 p-4 rounded-full">
                  <CheckCircle2 className="h-16 w-16 text-green-600" />
                </div>
              )}
              <p className={cn(
                "text-lg font-semibold",
                status === 'success' ? "text-green-600" : "text-gray-900"
              )}>
                {message}
              </p>
            </div>
          )}

          {status === 'error' && (
            <div className="mt-4 p-4 bg-red-50 border border-red-100 rounded-xl flex items-center gap-3 text-red-600">
              <AlertTriangle className="h-5 w-5" />
              <p className="text-sm font-medium">{message}</p>
            </div>
          )}
        </div>
        
        <div className="bg-gray-50 px-8 py-4 flex justify-between items-center border-t border-gray-100">
          <div className="flex items-center gap-2 text-gray-500 text-sm">
            <QrCode className="h-4 w-4" />
            <span>Apunta la cámara al código QR</span>
          </div>
          <Button variant="ghost" onClick={() => navigate('/inventory')}>
            Cancelar
          </Button>
        </div>
      </div>

      <div className="bg-blue-50 rounded-2xl p-6 border border-blue-100">
        <h3 className="text-blue-800 font-bold flex items-center gap-2 mb-2">
          <Maximize className="h-5 w-5" />
          ¿Cómo funciona?
        </h3>
        <ul className="text-sm text-blue-700 space-y-2 list-disc list-inside">
          <li>Asegúrate de que el producto tenga el QR activado en el inventario.</li>
          <li>Apunta la cámara de tu dispositivo hacia el código impreso o en pantalla.</li>
          <li>El sistema identificará el producto y cambiará su estado a "Pedido" automáticamente.</li>
        </ul>
      </div>
    </div>
  );
};

// --- Main App ---

export default function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Layout><Dashboard /></Layout>} />
        <Route path="/inventory" element={<Layout><Inventory /></Layout>} />
        <Route path="/movements" element={<Layout><Movements /></Layout>} />
        <Route path="/requesters" element={<Layout><Requesters /></Layout>} />
        <Route path="/attendance" element={<Layout><AttendancePage /></Layout>} />
        <Route path="/requests" element={<Layout><RequestsPage /></Layout>} />
        <Route path="/reports" element={<Layout><DailyReportsPage /></Layout>} />
        <Route path="/scanner" element={<Layout><ScannerPage /></Layout>} />
        <Route path="*" element={<Navigate to="/" />} />
      </Routes>
    </Router>
  );
}
