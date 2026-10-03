import React, { useState, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { 
  auth, 
  db, 
  googleProvider, 
  OperationType, 
  handleFirestoreError 
} from './firebase';
import { 
  onAuthStateChanged, 
  signInWithPopup, 
  signOut, 
  User as FirebaseUser 
} from 'firebase/auth';
import { 
  doc, 
  setDoc, 
  getDoc, 
  getDocs, 
  collection, 
  query, 
  where, 
  orderBy, 
  addDoc, 
  onSnapshot,
  Timestamp,
  updateDoc
} from 'firebase/firestore';
import { 
  BookOpen, 
  Sparkles, 
  User, 
  PenTool, 
  CheckCircle2, 
  Heart, 
  Calendar, 
  Award, 
  MessageSquare, 
  HelpCircle, 
  Smile, 
  LogOut, 
  ChevronRight, 
  ChevronLeft, 
  Plus, 
  Search, 
  Filter, 
  AlertTriangle, 
  Frown, 
  Meh, 
  Star, 
  Flame, 
  RefreshCw, 
  BookMarked, 
  Activity, 
  Clock, 
  Bookmark,
  Users,
  Compass,
  FileText,
  Check,
  Shield,
  Zap,
  Mic
} from 'lucide-react';

// Types
interface UserProfile {
  uid: string;
  email: string;
  displayName: string;
  photoURL: string;
  role: 'guru_wali' | 'guru_non_wali' | 'murid' | 'guru_bk' | 'admin';
  kelas: string;
  createdAt: any;
}

interface Story {
  id: string;
  title: string;
  authorId: string;
  authorName: string;
  authorKelas: string;
  content: string;
  illustrationUrl: string;
  fact: string;
  feeling: string;
  feelingEmoji: string;
  finding: string;
  future: string;
  associatedHabits: string[];
  status: 'draft' | 'submitted' | 'reviewed';
  createdAt: any;
  updatedAt: any;
  aiSummary?: string;
  aiFeedback?: string;
}

interface Feedback {
  id: string;
  storyId: string;
  reviewerId: string;
  reviewerName: string;
  reviewerRole: 'guru_wali' | 'guru_non_wali' | 'guru_bk';
  feedbackText: string;
  stars: number;
  createdAt: any;
}

interface HabitTracker {
  id: string; // studentId_date
  studentId: string;
  studentName: string;
  studentKelas: string;
  date: string; // YYYY-MM-DD
  bangun_pagi: boolean;
  beribadah: boolean;
  berolahraga: boolean;
  makan_sehat: boolean;
  gemar_belajar: boolean;
  bermasyarakat: boolean;
  tidur_cepat: boolean;
  updatedAt: any;
}

// Habits constants
const HABITS_INFO = [
  { id: 'bangun_pagi', label: 'Bangun Pagi', desc: 'Bangun pagi segar tepat waktu', icon: '🌅', color: 'from-amber-400 to-orange-500' },
  { id: 'beribadah', label: 'Beribadah', desc: 'Melaksanakan ibadah sesuai keyakinan', icon: '🙏', color: 'from-purple-400 to-indigo-500' },
  { id: 'berolahraga', label: 'Berolahraga', desc: 'Menjaga tubuh sehat dengan bergerak aktif', icon: '🏃', color: 'from-blue-400 to-emerald-500' },
  { id: 'makan_sehat', label: 'Makan Sehat & Bergizi', desc: 'Makan buah, sayur, makanan bernutrisi seimbang', icon: '🍎', color: 'from-green-400 to-teal-500' },
  { id: 'gemar_belajar', label: 'Gemar Belajar', desc: 'Membaca buku atau mempelajari hal baru', icon: '📚', color: 'from-rose-400 to-pink-500' },
  { id: 'bermasyarakat', label: 'Bermasyarakat', desc: 'Tolong-menolong dan bersosialisasi dengan teman/tetangga', icon: '🤝', color: 'from-sky-400 to-indigo-500' },
  { id: 'tidur_cepat', label: 'Tidur Cepat', desc: 'Tidur tidak larut malam agar istirahat cukup', icon: '🌙', color: 'from-slate-600 to-indigo-950' },
];

const FEELING_EMOJIS = [
  { emoji: '😊', label: 'Gembira / Senang', desc: 'Merasa ceria, bahagia, dan penuh energi.' },
  { emoji: '🤩', label: 'Sangat Bersemangat', desc: 'Merasa antusias, termotivasi, dan bergairah.' },
  { emoji: '😇', label: 'Damai / Bersyukur', desc: 'Merasa tenang, nyaman, dan berterima kasih.' },
  { emoji: '😟', label: 'Cemas / Takut', desc: 'Merasa khawatir, gugup, atau tidak aman.' },
  { emoji: '😢', label: 'Sedih', desc: 'Merasa kecewa, sunyi, atau kurang beruntung.' },
  { emoji: '😠', label: 'Kesal / Marah', desc: 'Merasa jengkel atau tidak setuju dengan keadaan.' },
];

// Fallback illustration card template generators
function getCuteFallbackBackground(theme: string) {
  const gradients = [
    'linear-gradient(135deg, #FF9A9E 0%, #FECFEF 99%, #FEF9E7 100%)',
    'linear-gradient(135deg, #A1C4FD 0%, #C2E9FB 100%)',
    'linear-gradient(135deg, #FEE140 0%, #FA709A 100%)',
    'linear-gradient(135deg, #84FAB0 0%, #8FD3F4 100%)',
    'linear-gradient(135deg, #E0C3FC 0%, #8EC5FC 100%)',
    'linear-gradient(135deg, #F6D365 0%, #FDA085 100%)',
  ];
  let hash = 0;
  for (let i = 0; i < theme.length; i++) {
    hash = theme.charCodeAt(i) + ((hash << 5) - hash);
  }
  const index = Math.abs(hash) % gradients.length;
  return gradients[index];
}

// Helper to normalize and match student class with teacher's assigned class(es), e.g. "7B, 7C" or "7A, 7B" or "SEMUA"
export const isTeacherClassMatch = (teacherKelasString: string | undefined, studentKelas: string | undefined): boolean => {
  if (!teacherKelasString || !studentKelas) return false;
  const tk = teacherKelasString.trim().toUpperCase();
  const sk = studentKelas.trim().toUpperCase();
  
  if (tk === 'SEMUA' || sk === 'SEMUA') return true;

  const normalize = (k: string) => {
    let str = k.toUpperCase().replace(/^KELAS\s+/, '').trim();
    str = str.replace(/VII/g, '7').replace(/VIII/g, '8').replace(/IX/g, '9');
    str = str.replace(/[-\s]/g, ''); // "7-B" -> "7B"
    return str;
  };

  const target = normalize(sk);
  const teacherClasses = tk.split(',').map(normalize).filter(Boolean);

  return teacherClasses.some(tc => tc === target || target.includes(tc) || tc.includes(target));
};

// PREMIUM OFFLINE MOCK DATA (Gives a stellar testing state during Firestore limit block)
const MOCK_STORIES: Story[] = [
  {
    id: 'mock_story_1',
    title: 'Merapikan Tempat Tidurku Sendiri Setiap Pagi',
    authorId: 'mock_student_1',
    authorName: 'Ahmad Fauzi',
    authorKelas: '7-A',
    content: 'Setiap pagi setelah bangun tidur tepat pukul 05.00, sekarang saya membiasakan diri merapikan bantal, guling, dan selimut sendiri. Dulu Ibu sering memarahiku karena kamar selalu berantakan seperti kapal pecah. Sekarang kamarku bersih, wangi, dan saya merasa bersemangat luar biasa untuk pergi belajar di sekolah!',
    illustrationUrl: 'fallback_theme_bangun_pagi',
    fact: 'Merapikan tempat tidur secara mandiri tepat setelah bangun tidur pukul 05.00 pagi.',
    feeling: 'Merasa sangat bangga, mandiri, dan senang karena mendapat pujian senyuman dari Ibu.',
    feelingEmoji: '😊',
    finding: 'Kedisiplinan kecil yang dimulai sesaat setelah membuka mata di pagi hari membawa energi positif yang sangat besar sepanjang hari.',
    future: 'Saya berkomitmen untuk terus konsisten merapikan tempat tidur tanpa perlu dibangunkan atau disuruh Ibu lagi.',
    associatedHabits: ['bangun_pagi', 'gemar_belajar'],
    status: 'reviewed',
    createdAt: { seconds: Date.now() / 1000 - 86400 * 2 },
    updatedAt: { seconds: Date.now() / 1000 - 86400 * 2 },
    aiSummary: 'Ahmad merapikan tempat tidur sendiri setiap pagi dan menyadari bahwa disiplin pagi mendatangkan energi positif.',
    aiFeedback: 'Hebat sekali Ahmad! Kedisiplinan pagimu adalah cerminan anak Indonesia yang mandiri dan bertanggung jawab. Teruskan!'
  },
  {
    id: 'mock_story_2',
    title: 'Keindahan Membagi Bekal Buah Apel Bersama Dika',
    authorId: 'mock_student_2',
    authorName: 'Rania Amalia',
    authorKelas: '7-A',
    content: 'Pagi ini saya membawa dua buah bekal apel merah yang manis dan segar dari rumah. Saat jam istirahat makan bekal tiba, saya melihat Dika hanya duduk terdiam di pojok kelas dengan wajah lesu. Ternyata dia terburu-buru pagi ini sehingga lupa membawa uang saku maupun bekal makanan. Saya pun mendekatinya, memotong apel milik saya, dan memberikan setengahnya kepada Dika. Kami makan bersama sambil tertawa riang.',
    illustrationUrl: 'fallback_theme_makan_sehat',
    fact: 'Membagi buah apel bekal sekolah kepada sahabat bernama Dika yang kelaparan karena tidak membawa makanan.',
    feeling: 'Merasa gembira dan bersemangat karena bisa mengenyangkan perut teman serta menjalin kebersamaan.',
    feelingEmoji: '🤩',
    finding: 'Berbagi bekal sehat tidak membuat porsi makanan kita berkurang, melainkan melipatgandakan rasa persahabatan dan kebahagiaan.',
    future: 'Ke depan, saya akan selalu membiasakan diri peka terhadap teman-teman di sekitar kelas yang sedang kesulitan.',
    associatedHabits: ['makan_sehat', 'bermasyarakat'],
    status: 'submitted',
    createdAt: { seconds: Date.now() / 1000 - 3600 * 4 },
    updatedAt: { seconds: Date.now() / 1000 - 3600 * 4 },
    aiSummary: 'Rania membagi buah apel sehat miliknya bersama Dika yang lupa membawa bekal di sekolah.',
    aiFeedback: 'Rania adalah anak berhati emas! Tindakan berbagimu sangat mencerminkan nilai luhur Pancasila dan persahabatan sejati.'
  },
  {
    id: 'mock_story_3',
    title: 'Melaksanakan Ibadah Shalat Maghrib Berjamaah',
    authorId: 'mock_student_1',
    authorName: 'Ahmad Fauzi',
    authorKelas: '7-A',
    content: 'Sore tadi saya sedang asyik bermain game di handphone bersama kakak. Tiba-tiba suara adzan maghrib berkumandang dengan merdu. Saya pun langsung meletakkan handphone, mengajak kakak untuk berhenti bermain, dan segera mengambil air wudhu. Kami sekeluarga kemudian melaksanakan shalat berjamaah yang dipimpin langsung oleh Ayah di mushola rumah kami.',
    illustrationUrl: 'fallback_theme_beribadah',
    fact: 'Berhenti bermain game handphone tepat saat adzan berkumandang untuk shalat berjamaah.',
    feeling: 'Merasa damai, tenang, tentram, dan penuh dengan rasa syukur bersama seluruh keluarga.',
    feelingEmoji: '😇',
    finding: 'Menjaga hubungan dengan Tuhan melalui ibadah tepat waktu memberikan ketenangan batin yang tidak bisa ditandingi oleh game apapun.',
    future: 'Saya berjanji akan selalu meletakkan mainan dan segera bersiap beribadah begitu panggilan adzan terdengar.',
    associatedHabits: ['beribadah'],
    status: 'reviewed',
    createdAt: { seconds: Date.now() / 1000 - 86400 * 1 },
    updatedAt: { seconds: Date.now() / 1000 - 86400 * 1 },
    aiSummary: 'Ahmad menghentikan permainannya demi beribadah berjamaah tepat waktu bersama keluarganya.',
    aiFeedback: 'Luar biasa komitmen spiritualmu, Ahmad! Menyeimbangkan waktu bermain dan beribadah adalah tanda anak hebat.'
  },
  {
    id: 'mock_story_4',
    title: 'Khawatir Tidak Bisa Mengikuti Ujian Matematika',
    authorId: 'mock_student_3',
    authorName: 'Siti Rahma',
    authorKelas: '7-A',
    content: 'Malam ini saya merasa sangat cemas karena besok ada ujian matematika bab pecahan. Saya sudah mencoba membaca buku, tetapi rumusnya tampak sangat membingungkan. Saya merasa takut mendapat nilai jelek dan mengecewakan Ibu yang sudah bekerja keras menyekolahkanku.',
    illustrationUrl: 'fallback_theme_gemar_belajar',
    fact: 'Merasakan kecemasan dan ketakutan yang mendalam menjelang ujian matematika bab pecahan.',
    feeling: 'Merasa cemas, gugup, dan takut tidak bisa mengerjakan soal ujian esok hari.',
    feelingEmoji: '😟',
    finding: 'Belajar terburu-buru di malam terakhir ujian sangat membuat stres. Seharusnya saya mencicil belajar jauh-jauh hari.',
    future: 'Saya akan meminta bantuan bimbingan tambahan dari Bapak/Ibu guru BK atau belajar kelompok dengan rania.',
    associatedHabits: ['gemar_belajar'],
    status: 'submitted',
    createdAt: { seconds: Date.now() / 1000 - 3600 * 2 },
    updatedAt: { seconds: Date.now() / 1000 - 3600 * 2 },
    aiSummary: 'Siti mengalami kecemasan yang berlebih menjelang ujian matematika karena belajar dengan sistem kebut semalam.',
    aiFeedback: 'Tenang Siti, kamu sudah berusaha sebaik mungkin. Cobalah menarik nafas dalam-dalam dan istirahatlah yang cukup.'
  }
];

const MOCK_USERS: { [uid: string]: UserProfile } = {
  'mock_student_1': {
    uid: 'mock_student_1',
    email: 'ahmad@siswa.belajar.id',
    displayName: 'Ahmad Fauzi',
    photoURL: 'https://api.dicebear.com/7.x/adventurer/svg?seed=Ahmad',
    role: 'murid',
    kelas: 'VII-A',
    createdAt: null
  },
  'mock_student_2': {
    uid: 'mock_student_2',
    email: 'rania@siswa.belajar.id',
    displayName: 'Rania Amalia',
    photoURL: 'https://api.dicebear.com/7.x/adventurer/svg?seed=Rania',
    role: 'murid',
    kelas: 'VII-A',
    createdAt: null
  },
  'mock_student_3': {
    uid: 'mock_student_3',
    email: 'siti@siswa.belajar.id',
    displayName: 'Siti Rahma',
    photoURL: 'https://api.dicebear.com/7.x/adventurer/svg?seed=Siti',
    role: 'murid',
    kelas: 'VII-A',
    createdAt: null
  }
};

const MOCK_FEEDBACKS: { [storyId: string]: Feedback[] } = {
  'mock_story_1': [
    {
      id: 'mock_fb_1',
      storyId: 'mock_story_1',
      reviewerId: 'mock_teacher',
      reviewerName: 'Budi Hartono, S.Pd',
      reviewerRole: 'guru_wali',
      feedbackText: 'Sangat luar biasa Ahmad! Mandiri di pagi hari melatih kedisiplinan hidup yang akan sangat bermanfaat saat kamu dewasa kelak. Bapak sangat bangga!',
      stars: 5,
      createdAt: { seconds: Date.now() / 1000 - 86400 }
    }
  ],
  'mock_story_3': [
    {
      id: 'mock_fb_2',
      storyId: 'mock_story_3',
      reviewerId: 'mock_teacher',
      reviewerName: 'Budi Hartono, S.Pd',
      reviewerRole: 'guru_wali',
      feedbackText: 'Hebat nak. Menyeimbangkan aktivitas hiburan dengan kewajiban ibadah bersama keluarga membentuk karakter yang seimbang dan religius.',
      stars: 5,
      createdAt: { seconds: Date.now() / 1000 - 3600 * 12 }
    }
  ]
};

const MOCK_HABITS: HabitTracker[] = [
  {
    id: 'mock_student_1_2026-10-02',
    studentId: 'mock_student_1',
    studentName: 'Ahmad Fauzi',
    studentKelas: '7-A',
    date: '2026-10-02',
    bangun_pagi: true,
    beribadah: true,
    berolahraga: false,
    makan_sehat: true,
    gemar_belajar: true,
    bermasyarakat: true,
    tidur_cepat: true,
    updatedAt: null
  },
  {
    id: 'mock_student_2_2026-10-02',
    studentId: 'mock_student_2',
    studentName: 'Rania Amalia',
    studentKelas: '7-A',
    date: '2026-10-02',
    bangun_pagi: true,
    beribadah: true,
    berolahraga: true,
    makan_sehat: true,
    gemar_belajar: true,
    bermasyarakat: true,
    tidur_cepat: false,
    updatedAt: null
  }
];

export default function App() {
  const [currentUser, setCurrentUser] = useState<FirebaseUser | null>(null);
  const [userProfile, setUserProfile] = useState<UserProfile | null>(null);
  const [loading, setLoading] = useState(true);
  const [profileLoading, setProfileLoading] = useState(false);

  // Firestore Quota Detection & Offline Sandbox Fallback
  const [isQuotaExceeded, setIsQuotaExceeded] = useState(false);
  const [offlineMode, setOfflineMode] = useState(false);

  // Registration states
  const [regRole, setRegRole] = useState<'guru_wali' | 'guru_non_wali' | 'murid' | 'guru_bk' | 'admin'>('murid');
  const [regKelas, setRegKelas] = useState('');

  // Story states
  const [stories, setStories] = useState<Story[]>([]);
  const [feedbacks, setFeedbacks] = useState<{ [storyId: string]: Feedback[] }>({});
  const [habitLogs, setHabitLogs] = useState<HabitTracker[]>([]);
  const [allUsers, setAllUsers] = useState<{ [uid: string]: UserProfile }>({});

  // Active View Tab: 'gallery' | 'write' | 'habits' | 'classroom' | 'bk_corner'
  const [activeTab, setActiveTab] = useState<string>('gallery');

  // Filter gallery
  const [searchQuery, setSearchQuery] = useState('');
  const [filterKelas, setFilterKelas] = useState('Semua');
  const [filterHabit, setFilterHabit] = useState('Semua');
  const [filterEmotion, setFilterEmotion] = useState('Semua');

  // Selected story for detail modal
  const [selectedStory, setSelectedStory] = useState<Story | null>(null);
  const [selectedTeacherClassFilter, setSelectedTeacherClassFilter] = useState<string>('Semua');
  const [newFeedbackText, setNewFeedbackText] = useState('');
  const [newFeedbackStars, setNewFeedbackStars] = useState(5);
  const [submittingFeedback, setSubmittingFeedback] = useState(false);

  // Write Story states
  const [writeStep, setWriteStep] = useState(1);
  const [storyTitle, setStoryTitle] = useState('');
  const [storyContent, setStoryContent] = useState('');
  const [storyFact, setStoryFact] = useState('');
  const [storyFeeling, setStoryFeeling] = useState('');
  const [storyFeelingEmoji, setStoryFeelingEmoji] = useState('😊');
  const [storyFinding, setStoryFinding] = useState('');
  const [storyFuture, setStoryFuture] = useState('');
  const [storyHabits, setStoryHabits] = useState<string[]>([]);
  const [aiAnalysisResult, setAiAnalysisResult] = useState<any>(null);
  const [analyzingStory, setAnalyzingStory] = useState(false);
  const [generatingIllustration, setGeneratingIllustration] = useState(false);
  const [generatedImageUrl, setGeneratedImageUrl] = useState<string | null>(null);
  const [submittingStory, setSubmittingStory] = useState(false);

  // Student habits daily log
  const [selectedDate, setSelectedDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [todayHabitLog, setTodayHabitLog] = useState<Partial<HabitTracker>>({});
  const [savingHabits, setSavingHabits] = useState(false);

  // Admin Panel states
  const [adminInputMode, setAdminInputMode] = useState<'excel' | 'manual' | 'bulk'>('excel');
  const [adminNewUserName, setAdminNewUserName] = useState('');
  const [adminNewUserEmail, setAdminNewUserEmail] = useState('');
  const [adminNewUserRole, setAdminNewUserRole] = useState<'guru_wali' | 'murid' | 'guru_bk' | 'admin'>('murid');
  const [adminNewUserKelas, setAdminNewUserKelas] = useState('');
  const [adminSubmitting, setAdminSubmitting] = useState(false);
  const [bulkInputText, setBulkInputText] = useState('');
  const [isBulkUploading, setIsBulkUploading] = useState(false);
  const [bulkImportRole, setBulkImportRole] = useState<'murid' | 'guru_wali'>('murid');
  const [adminSubTab, setAdminSubTab] = useState<'register' | 'guru_wali' | 'murid' | 'stories'>('register');

  // Excel File Upload states
  const [excelParsedUsers, setExcelParsedUsers] = useState<UserProfile[]>([]);
  const [excelFileName, setExcelFileName] = useState<string>('');
  const [excelImporting, setExcelImporting] = useState<boolean>(false);
  
  // Inline edit states
  const [editingUserUid, setEditingUserUid] = useState<string | null>(null);
  const [editUserName, setEditUserName] = useState('');
  const [editUserEmail, setEditUserEmail] = useState('');
  const [editUserRole, setEditUserRole] = useState<'guru_wali' | 'murid' | 'guru_bk' | 'admin'>('murid');
  const [editUserKelas, setEditUserKelas] = useState('');

  // Inline edit states for stories
  const [editingStoryId, setEditingStoryId] = useState<string | null>(null);
  const [editStoryTitle, setEditStoryTitle] = useState('');
  const [editStoryContent, setEditStoryContent] = useState('');
  const [editStoryStatus, setEditStoryStatus] = useState<'draft' | 'submitted' | 'reviewed'>('submitted');
  const [editStoryKelas, setEditStoryKelas] = useState('');

  // Voice & Image Story States
  const [isRecording, setIsRecording] = useState<string | null>(null);
  const [uploadedImageUrl, setUploadedImageUrl] = useState<string | null>(null);
  const [analyzingImage, setAnalyzingImage] = useState(false);

  // Gemini-powered Audio Recorder states
  const [mediaRecorder, setMediaRecorder] = useState<MediaRecorder | null>(null);
  const [isAiRecording, setIsAiRecording] = useState<string | null>(null);
  const [transcribingAudio, setTranscribingAudio] = useState(false);
  const [micPermissionDeniedModal, setMicPermissionDeniedModal] = useState(false);

  // Upload audio file and transcribe with Gemini AI
  const handleAudioFileUpload = async (e: React.ChangeEvent<HTMLInputElement>, field: 'title' | 'content' | 'fact' | 'feeling' | 'finding' | 'future') => {
    const file = e.target.files?.[0];
    if (!file) return;

    setTranscribingAudio(true);
    setIsAiRecording(field);

    try {
      const reader = new FileReader();
      reader.readAsDataURL(file);
      reader.onloadend = async () => {
        try {
          const base64Audio = reader.result as string;
          
          const response = await fetch('/api/gemini/transcribe', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ audio: base64Audio, mimeType: file.type || 'audio/mp3' })
          });

          if (!response.ok) {
            throw new Error('Gagal mentranskripsi file audio.');
          }

          const data = await response.json();
          const speechToText = data.text;

          if (speechToText) {
            if (field === 'title') setStoryTitle(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
            else if (field === 'content') setStoryContent(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
            else if (field === 'fact') setStoryFact(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
            else if (field === 'feeling') setStoryFeeling(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
            else if (field === 'finding') setStoryFinding(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
            else if (field === 'future') setStoryFuture(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
            
            alert("Horeee! 🎉 Rekaman suaramu berhasil ditranskripsikan oleh Gemini AI!");
          }
        } catch (err) {
          console.error("Audio file transcription error:", err);
          alert("Gagal mentranskripsi file audio. Pastikan file berformat audio (.mp3, .wav, .m4a, .webm).");
        } finally {
          setTranscribingAudio(false);
          setIsAiRecording(null);
        }
      };
    } catch (err) {
      console.error("FileReader failed:", err);
      setTranscribingAudio(false);
      setIsAiRecording(null);
    }
  };

  // Start recording voice note using MediaRecorder for Gemini AI speech-to-text
  const startMediaRecording = async (field: 'title' | 'content' | 'fact' | 'feeling' | 'finding' | 'future') => {
    try {
      if (!navigator.mediaDevices || !navigator.mediaDevices.getUserMedia) {
        setMicPermissionDeniedModal(true);
        return;
      }

      const stream = await navigator.mediaDevices.getUserMedia({ audio: true });
      let recorder;
      try {
        recorder = new MediaRecorder(stream, { mimeType: 'audio/webm' });
      } catch (e) {
        try {
          recorder = new MediaRecorder(stream, { mimeType: 'audio/mp4' });
        } catch (e2) {
          recorder = new MediaRecorder(stream); // Fallback for browsers like iOS Safari
        }
      }

      const chunks: Blob[] = [];
      recorder.ondataavailable = (e) => {
        if (e.data.size > 0) {
          chunks.push(e.data);
        }
      };

      recorder.onstop = async () => {
        const audioBlob = new Blob(chunks, { type: recorder.mimeType });
        setTranscribingAudio(true);
        try {
          const reader = new FileReader();
          reader.readAsDataURL(audioBlob);
          reader.onloadend = async () => {
            try {
              const base64Audio = reader.result as string;
              
              const response = await fetch('/api/gemini/transcribe', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ audio: base64Audio, mimeType: audioBlob.type })
              });

              if (!response.ok) {
                throw new Error('Gagal mentranskripsi audio.');
              }

              const data = await response.json();
              const speechToText = data.text;

              if (speechToText) {
                if (field === 'title') setStoryTitle(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
                else if (field === 'content') setStoryContent(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
                else if (field === 'fact') setStoryFact(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
                else if (field === 'feeling') setStoryFeeling(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
                else if (field === 'finding') setStoryFinding(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
                else if (field === 'future') setStoryFuture(prev => (prev ? prev.trim() + ' ' + speechToText : speechToText));
              }
            } catch (err) {
              console.error("Transcription error in reader:", err);
              alert("Gagal mentranskripsi rekaman suara dengan Gemini AI.");
            } finally {
              setTranscribingAudio(false);
              setIsAiRecording(null);
            }
          };
        } catch (err) {
          console.error("Transcription failed outer:", err);
          alert("Gagal memproses suara Anda dengan Gemini AI.");
          setTranscribingAudio(false);
          setIsAiRecording(null);
        } finally {
          stream.getTracks().forEach(track => track.stop());
        }
      };

      recorder.start();
      setMediaRecorder(recorder);
      setIsAiRecording(field);
    } catch (error: any) {
      console.warn("Microphone access failed:", error);
      setIsAiRecording(null);
      setTranscribingAudio(false);
      setMicPermissionDeniedModal(true);
    }
  };

  const stopMediaRecording = () => {
    if (mediaRecorder && mediaRecorder.state !== 'inactive') {
      mediaRecorder.stop();
    }
  };

  // Real-time Voice Note transcription (Speech to Text)
  const startSpeechRecognition = (field: 'title' | 'content' | 'fact' | 'feeling' | 'finding' | 'future') => {
    const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
    if (!SpeechRecognition) {
      setMicPermissionDeniedModal(true);
      return;
    }

    const recognition = new SpeechRecognition();
    recognition.lang = 'id-ID';
    recognition.interimResults = false;
    recognition.maxAlternatives = 1;

    recognition.onstart = () => {
      setIsRecording(field);
    };

    recognition.onerror = (event: any) => {
      console.warn("Speech recognition error:", event.error);
      setIsRecording(null);
      if (event.error === 'not-allowed' || event.error === 'service-not-allowed' || event.error === 'audio-capture') {
        setMicPermissionDeniedModal(true);
      }
    };

    recognition.onend = () => {
      setIsRecording(null);
    };

    recognition.onresult = (event: any) => {
      const speechToText = event.results[0][0].transcript;
      if (field === 'title') setStoryTitle(prev => prev + ' ' + speechToText);
      else if (field === 'content') setStoryContent(prev => prev + ' ' + speechToText);
      else if (field === 'fact') setStoryFact(prev => prev + ' ' + speechToText);
      else if (field === 'feeling') setStoryFeeling(prev => prev + ' ' + speechToText);
      else if (field === 'finding') setStoryFinding(prev => prev + ' ' + speechToText);
      else if (field === 'future') setStoryFuture(prev => prev + ' ' + speechToText);
    };

    recognition.start();
  };

  // Convert uploaded image to Base64 and generate story with Gemini
  const handleImageUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onloadend = () => {
      setUploadedImageUrl(reader.result as string);
    };
    reader.readAsDataURL(file);
  };

  const handleGenerateStoryFromImage = async () => {
    if (!uploadedImageUrl) {
      alert("Silakan unggah foto kegiatan Anda terlebih dahulu!");
      return;
    }

    setAnalyzingImage(true);
    try {
      const response = await fetch('/api/gemini/analyze-image', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ image: uploadedImageUrl })
      });

      if (!response.ok) {
        throw new Error('Gagal menganalisis foto.');
      }

      const data = await response.json();
      setStoryTitle(data.title || '');
      setStoryContent(data.content || '');
      setStoryFact(data.fact || '');
      setStoryFeeling(data.feeling || '');
      setStoryFinding(data.finding || '');
      setStoryFuture(data.future || '');
      
      // Auto-set the generated image as the story illustration too!
      setGeneratedImageUrl(uploadedImageUrl);
      
      alert("Horeee! 🎉 Si CERDAS AI berhasil membuatkan cerita dan refleksi 4F berdasarkan fotomu! Silakan tinjau di langkah berikutnya.");
      setWriteStep(2); // Jump to Step 2 so they can review the 4F reflections
    } catch (error) {
      console.error("Image Analysis Error:", error);
      alert("Gagal menganalisis gambar dengan AI. Coba gunakan foto lain atau ketik cerita secara manual.");
    } finally {
      setAnalyzingImage(false);
    }
  };

  // Listen to Auth State
  useEffect(() => {
    const unsubscribe = onAuthStateChanged(auth, async (user) => {
      setCurrentUser(user);
      if (user) {
        setProfileLoading(true);
        try {
          // Attempt Firestore read
          const profileDoc = await getDoc(doc(db, 'users', user.uid));
          if (profileDoc.exists()) {
            const data = profileDoc.data() as UserProfile;
            setUserProfile(data);
            
            // Route user based on role
            if (data.role === 'murid') {
              setActiveTab('gallery');
            } else if (data.role === 'guru_wali') {
              setActiveTab('classroom');
            } else if (data.role === 'guru_bk') {
              setActiveTab('bk_corner');
            } else {
              setActiveTab('gallery');
            }
          } else {
            setUserProfile(null);
          }
        } catch (error: any) {
          console.warn("Gagal memuat profil pengguna dari Firestore (Quota mungkin tercapai):", error?.message);
          
          // Auto-trigger offline sandbox if quota error detected
          if (error.message?.includes('Quota') || error.message?.includes('quota') || error.message?.includes('exhausted') || error.message?.includes('resource-exhausted')) {
            setIsQuotaExceeded(true);
            setOfflineMode(true);
            
            // Set temporary local user profile based on Google Sign-In details
            const localProfile: UserProfile = {
              uid: user.uid,
              email: user.email || 'isumayasa91@guru.smp.belajar.id',
              displayName: user.displayName || 'I Sumayasa (Guru Wali)',
              photoURL: user.photoURL || 'https://api.dicebear.com/7.x/adventurer/svg?seed=' + user.uid,
              role: 'guru_wali', // Default homeroom teacher to test the app elegantly
              kelas: '7-A',
              createdAt: null
            };
            setUserProfile(localProfile);
            setActiveTab('classroom');
          }
        } finally {
          setProfileLoading(false);
        }
      } else {
        setUserProfile(null);
      }
      setLoading(false);
    });
    return unsubscribe;
  }, []);

  // Listen to All Stories and Users
  useEffect(() => {
    if (!currentUser || !userProfile || offlineMode) return;

    // Load users
    const usersUnsubscribe = onSnapshot(collection(db, 'users'), (snapshot) => {
      const usersMap: { [uid: string]: UserProfile } = {};
      snapshot.forEach((doc) => {
        usersMap[doc.id] = doc.data() as UserProfile;
      });
      setAllUsers(usersMap);
    }, (error: any) => {
      console.warn("Firestore listener quota limits exceeded:", error);
      setIsQuotaExceeded(true);
      setOfflineMode(true);
    });

    // Realtime stories
    let storiesQuery = query(collection(db, 'stories'), orderBy('createdAt', 'desc'));
    const storiesUnsubscribe = onSnapshot(storiesQuery, (snapshot) => {
      const storiesList: Story[] = [];
      snapshot.forEach((doc) => {
        storiesList.push({ id: doc.id, ...doc.data() } as Story);
      });
      setStories(storiesList);
    }, (error: any) => {
      console.warn("Firestore listener quota limits exceeded:", error);
      setIsQuotaExceeded(true);
      setOfflineMode(true);
    });

    // Habits tracker logs
    let habitsQuery;
    if (userProfile?.role === 'murid') {
      habitsQuery = query(collection(db, 'habits'), where('studentId', '==', currentUser.uid));
    } else {
      habitsQuery = query(collection(db, 'habits'), orderBy('date', 'desc'));
    }

    const habitsUnsubscribe = onSnapshot(habitsQuery, (snapshot) => {
      const habitsList: HabitTracker[] = [];
      snapshot.forEach((doc) => {
        habitsList.push({ id: doc.id, ...doc.data() } as HabitTracker);
      });
      setHabitLogs(habitsList);
    }, (error: any) => {
      console.warn("Habits snapshot query limited or quota reached:", error?.message);
    });

    return () => {
      usersUnsubscribe();
      storiesUnsubscribe();
      habitsUnsubscribe();
    };
  }, [currentUser, userProfile, offlineMode]);

  // Load Feedbacks for Selected Story when it changes
  useEffect(() => {
    if (!selectedStory || offlineMode) return;

    const feedbackQuery = query(collection(db, 'stories', selectedStory.id, 'feedbacks'), orderBy('createdAt', 'desc'));
    const unsubscribe = onSnapshot(feedbackQuery, (snapshot) => {
      const list: Feedback[] = [];
      snapshot.forEach((doc) => {
        list.push({ id: doc.id, ...doc.data() } as Feedback);
      });
      setFeedbacks(prev => ({ ...prev, [selectedStory.id]: list }));
    }, (error: any) => {
      console.warn("Feedback listener failed, using offline", error);
    });

    return unsubscribe;
  }, [selectedStory, offlineMode]);

  // Load and sync local state in offlineMode
  useEffect(() => {
    if (!offlineMode) return;

    // Load initial mock or local storage
    const localStories = localStorage.getItem('cerdas_stories');
    if (localStories) {
      setStories(JSON.parse(localStories));
    } else {
      setStories(MOCK_STORIES);
      localStorage.setItem('cerdas_stories', JSON.stringify(MOCK_STORIES));
    }

    const localUsers = localStorage.getItem('cerdas_users');
    if (localUsers) {
      setAllUsers(JSON.parse(localUsers));
    } else {
      setAllUsers(MOCK_USERS);
      localStorage.setItem('cerdas_users', JSON.stringify(MOCK_USERS));
    }

    const localFeedbacks = localStorage.getItem('cerdas_feedbacks');
    if (localFeedbacks) {
      setFeedbacks(JSON.parse(localFeedbacks));
    } else {
      setFeedbacks(MOCK_FEEDBACKS);
      localStorage.setItem('cerdas_feedbacks', JSON.stringify(MOCK_FEEDBACKS));
    }

    const localHabits = localStorage.getItem('cerdas_habits');
    if (localHabits) {
      setHabitLogs(JSON.parse(localHabits));
    } else {
      setHabitLogs(MOCK_HABITS);
      localStorage.setItem('cerdas_habits', JSON.stringify(MOCK_HABITS));
    }
  }, [offlineMode]);

  // Handle current student daily habits selection change
  useEffect(() => {
    if (!currentUser || userProfile?.role !== 'murid') return;

    const habitId = `${currentUser.uid}_${selectedDate}`;

    if (offlineMode) {
      const matchedLog = habitLogs.find(l => l.id === habitId);
      if (matchedLog) {
        setTodayHabitLog(matchedLog);
      } else {
        setTodayHabitLog({
          studentId: currentUser.uid,
          studentName: userProfile.displayName,
          studentKelas: userProfile.kelas,
          date: selectedDate,
          bangun_pagi: false,
          beribadah: false,
          berolahraga: false,
          makan_sehat: false,
          gemar_belajar: false,
          bermasyarakat: false,
          tidur_cepat: false,
        });
      }
      return;
    }

    const loadHabit = async () => {
      try {
        const habitDoc = await getDoc(doc(db, 'habits', habitId));
        if (habitDoc.exists()) {
          setTodayHabitLog(habitDoc.data() as HabitTracker);
        } else {
          setTodayHabitLog({
            studentId: currentUser.uid,
            studentName: userProfile.displayName,
            studentKelas: userProfile.kelas,
            date: selectedDate,
            bangun_pagi: false,
            beribadah: false,
            berolahraga: false,
            makan_sehat: false,
            gemar_belajar: false,
            bermasyarakat: false,
            tidur_cepat: false,
          });
        }
      } catch (e: any) {
        console.warn("Catatan kebiasaan belum tersedia di Firestore atau menggunakan nilai draf lokal:", e?.message);
        setTodayHabitLog({
          studentId: currentUser.uid,
          studentName: userProfile.displayName,
          studentKelas: userProfile.kelas,
          date: selectedDate,
          bangun_pagi: false,
          beribadah: false,
          berolahraga: false,
          makan_sehat: false,
          gemar_belajar: false,
          bermasyarakat: false,
          tidur_cepat: false,
        });
      }
    };
    loadHabit();
  }, [currentUser, userProfile, selectedDate, offlineMode, habitLogs]);

  const handleLogin = async () => {
    try {
      await signInWithPopup(auth, googleProvider);
    } catch (error) {
      console.error("Gagal masuk dengan Google:", error);
      // Fallback to manual offline login
      setIsQuotaExceeded(true);
      setOfflineMode(true);
      setCurrentUser({
        uid: 'mock_user_1',
        displayName: 'I Sumayasa (Demo)',
        email: 'isumayasa91@guru.smp.belajar.id',
        photoURL: 'https://api.dicebear.com/7.x/adventurer/svg?seed=isumayasa'
      } as any);
      setUserProfile({
        uid: 'mock_user_1',
        displayName: 'I Sumayasa (Demo)',
        email: 'isumayasa91@guru.smp.belajar.id',
        photoURL: 'https://api.dicebear.com/7.x/adventurer/svg?seed=isumayasa',
        role: 'guru_bk',
        kelas: 'SEMUA',
        createdAt: null
      });
      setActiveTab('bk_corner');
    }
  };

  const handleLogout = async () => {
    try {
      await signOut(auth);
    } catch (e) {
      console.warn("Logout error:", e);
    }
    setCurrentUser(null);
    setUserProfile(null);
    setSelectedStory(null);
    setActiveTab('gallery');
  };

  // Complete user profile registration
  const handleRegisterProfile = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentUser) return;

    setProfileLoading(true);
    const cleanKelas = regRole === 'murid' || regRole === 'guru_wali' ? regKelas.trim().toUpperCase() : 'SEMUA';
    
    if ((regRole === 'murid' || regRole === 'guru_wali') && !cleanKelas) {
      alert("Harap masukkan kelas Anda (contoh: 7-A, 8-B)");
      setProfileLoading(false);
      return;
    }

    if (offlineMode) {
      const profilePayload: UserProfile = {
        uid: currentUser.uid,
        email: currentUser.email || '',
        displayName: currentUser.displayName || 'Pengguna CERDAS',
        photoURL: currentUser.photoURL || 'https://api.dicebear.com/7.x/adventurer/svg?seed=' + currentUser.uid,
        role: regRole,
        kelas: cleanKelas,
        createdAt: { seconds: Date.now() / 1000 }
      };

      setUserProfile(profilePayload);
      const updatedUsers = { ...allUsers, [currentUser.uid]: profilePayload };
      setAllUsers(updatedUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(updatedUsers));

      if (regRole === 'murid') {
        setActiveTab('gallery');
      } else if (regRole === 'guru_wali') {
        setActiveTab('classroom');
      } else if (regRole === 'guru_bk') {
        setActiveTab('bk_corner');
      } else {
        setActiveTab('gallery');
      }
      setProfileLoading(false);
      return;
    }

    try {
      const profilePayload: UserProfile = {
        uid: currentUser.uid,
        email: currentUser.email || '',
        displayName: currentUser.displayName || 'Pengguna CERDAS',
        photoURL: currentUser.photoURL || 'https://api.dicebear.com/7.x/adventurer/svg?seed=' + currentUser.uid,
        role: regRole,
        kelas: cleanKelas,
        createdAt: Timestamp.now()
      };

      await setDoc(doc(db, 'users', currentUser.uid), profilePayload);
      setUserProfile(profilePayload);
      
      if (regRole === 'murid') {
        setActiveTab('gallery');
      } else if (regRole === 'guru_wali') {
        setActiveTab('classroom');
      } else if (regRole === 'guru_bk') {
        setActiveTab('bk_corner');
      } else {
        setActiveTab('gallery');
      }
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, `users/${currentUser.uid}`);
    } finally {
      setProfileLoading(false);
    }
  };

  // Student Story analysis with Si CERDAS AI
  const handleAnalyzeStoryWithAI = async () => {
    const cleanTitle = storyTitle.trim();
    const cleanContent = storyContent.trim();

    if (cleanTitle.length < 3) {
      alert("⚠️ Judul cerita terlalu singkat! Harap isi judul minimal 3 karakter.");
      return;
    }

    if (cleanContent.length < 30) {
      alert(`⚠️ Cerita terlalu singkat (${cleanContent.length} karakter)! Sistem tidak dapat menerima cerita kurang dari 30 karakter. Mohon tulis cerita lebih lengkap (30 - 5000 karakter).`);
      return;
    }

    if (cleanContent.length > 5000) {
      alert(`⚠️ Cerita melebihi batas (${cleanContent.length} karakter)! Maksimal panjang cerita adalah 5000 karakter.`);
      return;
    }

    if (!storyFact.trim()) {
      alert("Pastikan Fakta (Fact) refleksi telah diisi terlebih dahulu!");
      return;
    }

    setAnalyzingStory(true);
    setAiAnalysisResult(null);

    try {
      const response = await fetch('/api/gemini/analyze', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          title: storyTitle,
          content: storyContent,
          fact: storyFact,
          feeling: storyFeeling,
          finding: storyFinding,
          future: storyFuture,
          habits: storyHabits
        })
      });

      if (!response.ok) {
        throw new Error('Gagal mendapatkan ulasan AI.');
      }

      const data = await response.json();
      setAiAnalysisResult(data);
      
      if (data.detectedHabits && data.detectedHabits.length > 0) {
        const mergedHabits = Array.from(new Set([...storyHabits, ...data.detectedHabits.map((h: string) => h.toLowerCase().replace(/ /g, '_'))]));
        setStoryHabits(mergedHabits.filter(h => HABITS_INFO.some(info => info.id === h)));
      }
      
      setWriteStep(4);
    } catch (error) {
      console.error("Analysis Error:", error);
      // Give a friendly client-side fallback review if Gemini is offline
      const genericFeedback = {
        summary: `Cerita hebat tentang "${storyTitle}" yang merefleksikan nilai-nilai positif pilar kebiasaan baik anak bangsa.`,
        friendlyFeedback: `Bapak sangat terharu membaca kisahmu nak! Refleksi 4F-mu sangat matang, perbuatan berbagi serta ketekunanmu melambangkan generasi hebat Indonesia!`,
        detectedHabits: storyHabits,
        illustrationPrompt: `A lovely colorful children illustration, school kids, warm colors, watercolor styled`
      };
      setAiAnalysisResult(genericFeedback);
      setWriteStep(4);
    } finally {
      setAnalyzingStory(false);
    }
  };

  // Generate dynamic cover art via Gemini Flash Image
  const handleGenerateIllustration = async () => {
    const prompt = aiAnalysisResult?.illustrationPrompt || `A happy student storytelling scene in Indonesian school classroom, cute cartoon vector style`;
    setGeneratingIllustration(true);
    try {
      const response = await fetch('/api/gemini/generate-illustration', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ prompt })
      });
      
      if (!response.ok) throw new Error('Gagal memuat gambar');
      const data = await response.json();
      if (data.imageUrl) {
        setGeneratedImageUrl(data.imageUrl);
      } else {
        setGeneratedImageUrl('FALLBACK');
      }
    } catch (error) {
      console.error("Illustration Error:", error);
      setGeneratedImageUrl('FALLBACK');
    } finally {
      setGeneratingIllustration(false);
    }
  };

  // Submit/Draft story saving
  const handleSaveStory = async (status: 'draft' | 'submitted') => {
    if (!currentUser || !userProfile) return;

    const cleanTitle = storyTitle.trim();
    const cleanContent = storyContent.trim();

    if (cleanTitle.length < 3) {
      alert("⚠️ Judul cerita terlalu singkat! Harap isi judul minimal 3 karakter.");
      return;
    }

    if (cleanContent.length < 30) {
      alert(`⚠️ Cerita terlalu singkat (${cleanContent.length} karakter)! Setiap cerita minimal harus berisi 30 karakter agar sistem dapat menerima dan menyimpannya.`);
      return;
    }

    if (cleanContent.length > 5000) {
      alert(`⚠️ Cerita melebihi batas (${cleanContent.length} karakter)! Maksimal panjang cerita adalah 5000 karakter.`);
      return;
    }

    setSubmittingStory(true);
    const storyId = 'story_' + Math.random().toString(36).substring(2, 15);
    
    let coverImg = generatedImageUrl;
    if (!coverImg || coverImg === 'FALLBACK') {
      coverImg = `fallback_theme_${storyHabits[0] || 'gemar_belajar'}`;
    }

    const storyPayload: Story = {
      id: storyId,
      title: storyTitle,
      authorId: currentUser.uid,
      authorName: userProfile.displayName,
      authorKelas: userProfile.kelas,
      content: storyContent,
      illustrationUrl: coverImg,
      fact: storyFact,
      feeling: storyFeeling,
      feelingEmoji: storyFeelingEmoji,
      finding: storyFinding,
      future: storyFuture,
      associatedHabits: storyHabits,
      status: status,
      createdAt: { seconds: Date.now() / 1000 },
      updatedAt: { seconds: Date.now() / 1000 },
      aiSummary: aiAnalysisResult?.summary || "Cerita anak mandiri reflektif.",
      aiFeedback: aiAnalysisResult?.friendlyFeedback || "Hebat sekali ceritamu! Teruslah rajin berkarya."
    };

    if (offlineMode) {
      const updated = [storyPayload, ...stories];
      setStories(updated);
      localStorage.setItem('cerdas_stories', JSON.stringify(updated));
      
      // Reset forms
      setStoryTitle('');
      setStoryContent('');
      setStoryFact('');
      setStoryFeeling('');
      setStoryFeelingEmoji('😊');
      setStoryFinding('');
      setStoryFuture('');
      setStoryHabits([]);
      setAiAnalysisResult(null);
      setGeneratedImageUrl(null);
      setWriteStep(1);

      alert(status === 'submitted' ? "Yayy! [Offline Sandbox] Cerita hebatmu berhasil disimpan ke browser! 🎉" : "Ceritamu disimpan sebagai draf [Offline].");
      setActiveTab('gallery');
      setSubmittingStory(false);
      return;
    }

    try {
      const dbPayload = {
        ...storyPayload,
        createdAt: Timestamp.now(),
        updatedAt: Timestamp.now()
      };
      await setDoc(doc(db, 'stories', storyId), dbPayload);
      
      setStoryTitle('');
      setStoryContent('');
      setStoryFact('');
      setStoryFeeling('');
      setStoryFeelingEmoji('😊');
      setStoryFinding('');
      setStoryFuture('');
      setStoryHabits([]);
      setAiAnalysisResult(null);
      setGeneratedImageUrl(null);
      setWriteStep(1);

      alert(status === 'submitted' ? "Yayy! Cerita hebatmu berhasil dikirim ke Guru! 🎉" : "Ceritamu disimpan sebagai draf.");
      setActiveTab('gallery');
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, 'stories');
    } finally {
      setSubmittingStory(false);
    }
  };

  // Save student habits tracking check
  const handleSaveHabits = async () => {
    if (!currentUser || !userProfile) return;

    setSavingHabits(true);

    if (offlineMode) {
      const habitId = `${currentUser.uid}_${selectedDate}`;
      const payload = {
        ...todayHabitLog,
        id: habitId,
        studentId: currentUser.uid,
        studentName: userProfile.displayName,
        studentKelas: userProfile.kelas,
        date: selectedDate,
        updatedAt: { seconds: Date.now() / 1000 }
      };

      const updatedLogs = [payload as HabitTracker, ...habitLogs.filter(l => l.id !== habitId)];
      setHabitLogs(updatedLogs);
      localStorage.setItem('cerdas_habits', JSON.stringify(updatedLogs));
      alert("[Offline Sandbox] Catatan 7 kebiasaanmu hari ini berhasil disimpan! Tetap konsisten ya! 💪🏆");
      setSavingHabits(false);
      return;
    }

    try {
      const habitId = `${currentUser.uid}_${selectedDate}`;
      const payload = {
        ...todayHabitLog,
        id: habitId,
        studentId: currentUser.uid,
        studentName: userProfile.displayName,
        studentKelas: userProfile.kelas,
        date: selectedDate,
        updatedAt: Timestamp.now()
      };

      await setDoc(doc(db, 'habits', habitId), payload);
      alert("Hebat! Catatan 7 kebiasaanmu hari ini berhasil disimpan! Tetap konsisten ya! 💪🏆");
    } catch (error) {
      handleFirestoreError(error, OperationType.WRITE, `habits/${currentUser.uid}_${selectedDate}`);
    } finally {
      setSavingHabits(false);
    }
  };

  // Submit Feedback / Reviews (Guru / BK)
  const handleSubmitFeedback = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!currentUser || !userProfile || !selectedStory || !newFeedbackText.trim()) return;

    setSubmittingFeedback(true);
    const feedbackId = 'fb_' + Math.random().toString(36).substring(2, 15);
    const storyId = selectedStory.id;

    if (offlineMode) {
      const payload: Feedback = {
        id: feedbackId,
        storyId: storyId,
        reviewerId: currentUser.uid,
        reviewerName: userProfile.displayName,
        reviewerRole: userProfile.role as any,
        feedbackText: newFeedbackText,
        stars: newFeedbackStars,
        createdAt: { seconds: Date.now() / 1000 }
      };

      const currentFeedbacks = feedbacks[storyId] || [];
      const updatedFeedbacks = {
        ...feedbacks,
        [storyId]: [payload, ...currentFeedbacks]
      };
      setFeedbacks(updatedFeedbacks);
      localStorage.setItem('cerdas_feedbacks', JSON.stringify(updatedFeedbacks));

      // Update story status
      const updatedStories = stories.map(s => s.id === storyId ? { ...s, status: 'reviewed' as const } : s);
      setStories(updatedStories);
      localStorage.setItem('cerdas_stories', JSON.stringify(updatedStories));

      setSelectedStory(prev => prev ? { ...prev, status: 'reviewed' } : null);
      setNewFeedbackText('');
      alert("[Offline Sandbox] Apresiasi luar biasamu berhasil disimpan di browser! 🌟❤️");
      setSubmittingFeedback(false);
      return;
    }

    try {
      const payload: Feedback = {
        id: feedbackId,
        storyId: storyId,
        reviewerId: currentUser.uid,
        reviewerName: userProfile.displayName,
        reviewerRole: userProfile.role as any,
        feedbackText: newFeedbackText,
        stars: newFeedbackStars,
        createdAt: Timestamp.now()
      };

      await setDoc(doc(db, 'stories', storyId, 'feedbacks', feedbackId), payload);
      await updateDoc(doc(db, 'stories', storyId), {
        status: 'reviewed',
        updatedAt: Timestamp.now()
      });

      setSelectedStory(prev => prev ? { ...prev, status: 'reviewed' } : null);
      setNewFeedbackText('');
      alert("Apresiasi dan saran luar biasamu telah dikirimkan ke Murid! 🌟❤️");
    } catch (error) {
      handleFirestoreError(error, OperationType.CREATE, `stories/${storyId}/feedbacks/${feedbackId}`);
    } finally {
      setSubmittingFeedback(false);
    }
  };

  // Get active student logs statistics
  const getStudentHabitStats = (studentId: string) => {
    const studentLogs = habitLogs.filter(log => log.studentId === studentId);
    if (studentLogs.length === 0) return { score: 0, count: 0 };

    let totalTicks = 0;
    studentLogs.forEach(log => {
      if (log.bangun_pagi) totalTicks++;
      if (log.beribadah) totalTicks++;
      if (log.berolahraga) totalTicks++;
      if (log.makan_sehat) totalTicks++;
      if (log.gemar_belajar) totalTicks++;
      if (log.bermasyarakat) totalTicks++;
      if (log.tidur_cepat) totalTicks++;
    });

    const possibleTicks = studentLogs.length * 7;
    const percentage = Math.round((totalTicks / possibleTicks) * 100);

    return {
      score: percentage,
      count: studentLogs.length
    };
  };

  // Admin Panel User Creation
  const handleAdminCreateUser = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!adminNewUserName.trim() || !adminNewUserEmail.trim()) {
      alert("Harap isi Nama Lengkap dan Email!");
      return;
    }

    setAdminSubmitting(true);
    const mockUid = 'user_' + Math.random().toString(36).substring(2, 15);
    const cleanKelas = adminNewUserRole === 'murid' || adminNewUserRole === 'guru_wali' ? adminNewUserKelas.trim().toUpperCase() : 'SEMUA';

    const newUserPayload: UserProfile = {
      uid: mockUid,
      email: adminNewUserEmail.trim().toLowerCase(),
      displayName: adminNewUserName.trim(),
      photoURL: `https://api.dicebear.com/7.x/adventurer/svg?seed=${mockUid}`,
      role: adminNewUserRole,
      kelas: cleanKelas,
      createdAt: { seconds: Date.now() / 1000 }
    };

    if (offlineMode) {
      const updatedUsers = { ...allUsers, [mockUid]: newUserPayload };
      setAllUsers(updatedUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(updatedUsers));
      alert(`[Offline Sandbox] Berhasil menambahkan ${adminNewUserRole === 'murid' ? 'Murid' : adminNewUserRole === 'guru_wali' ? 'Guru Wali' : 'Pengguna'} baru!`);
      setAdminNewUserName('');
      setAdminNewUserEmail('');
      setAdminNewUserKelas('');
      setAdminSubmitting(false);
      return;
    }

    try {
      const dbPayload = {
        ...newUserPayload,
        createdAt: Timestamp.now()
      };
      await setDoc(doc(db, 'users', mockUid), dbPayload);
      alert(`Berhasil mendaftarkan ${adminNewUserRole === 'murid' ? 'Murid' : adminNewUserRole === 'guru_wali' ? 'Guru Wali' : 'Pengguna'} baru ke database!`);
      setAdminNewUserName('');
      setAdminNewUserEmail('');
      setAdminNewUserKelas('');
    } catch (err) {
      handleFirestoreError(err, OperationType.CREATE, `users/${mockUid}`);
    } finally {
      setAdminSubmitting(false);
    }
  };

  // Bulk copy-paste upload for users (murid & guru_wali)
  const handleAdminBulkUpload = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!bulkInputText.trim()) {
      alert("Harap masukkan data masal terlebih dahulu!");
      return;
    }

    setIsBulkUploading(true);
    const lines = bulkInputText.split('\n');
    const validUsers: UserProfile[] = [];
    const errors: string[] = [];

    lines.forEach((line, index) => {
      const trimmed = line.trim();
      if (!trimmed) return;

      const cols = trimmed.includes('\t') ? trimmed.split('\t') : trimmed.split(',');
      if (cols.length < 2) {
        errors.push(`Baris ${index + 1}: Nama dan Email tidak lengkap.`);
        return;
      }

      const name = cols[0].trim();
      const email = cols[1].trim().toLowerCase();
      
      if (!email.includes('@')) {
        errors.push(`Baris ${index + 1}: Email "${email}" tidak valid`);
        return;
      }

      // Detect role and class
      let mappedRole: 'murid' | 'guru_wali' | 'guru_bk' | 'admin' = bulkImportRole;
      let rawKelas = '7-A';

      if (cols.length >= 4) {
        // cols: [Name, Email, Role, Class]
        const rawRole = cols[2].trim().toLowerCase();
        if (rawRole.includes('wali') || rawRole.includes('guru')) {
          mappedRole = 'guru_wali';
        } else if (rawRole.includes('bk') || rawRole.includes('konselor')) {
          mappedRole = 'guru_bk';
        } else if (rawRole.includes('admin')) {
          mappedRole = 'admin';
        } else {
          mappedRole = 'murid';
        }
        rawKelas = cols[3].trim().toUpperCase();
      } else if (cols.length === 3) {
        // cols: [Name, Email, Class]
        rawKelas = cols[2].trim().toUpperCase();
      }

      const mockUid = 'user_bulk_' + Math.random().toString(36).substring(2, 10);
      validUsers.push({
        uid: mockUid,
        email,
        displayName: name,
        photoURL: `https://api.dicebear.com/7.x/adventurer/svg?seed=${mockUid}`,
        role: mappedRole,
        kelas: mappedRole === 'murid' || mappedRole === 'guru_wali' ? rawKelas : 'SEMUA',
        createdAt: { seconds: Date.now() / 1000 }
      });
    });

    if (validUsers.length === 0) {
      alert("Tidak ada data valid yang bisa diimpor. Harap cek kembali format.");
      setIsBulkUploading(false);
      return;
    }

    if (offlineMode) {
      const updatedUsers = { ...allUsers };
      validUsers.forEach(u => {
        updatedUsers[u.uid] = u;
      });
      setAllUsers(updatedUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(updatedUsers));
      alert(`[Offline Sandbox] Sukses mengimpor ${validUsers.length} pengguna baru secara masal! 🎉`);
      setBulkInputText('');
      setIsBulkUploading(false);
      return;
    }

    try {
      let count = 0;
      for (const u of validUsers) {
        const dbPayload = {
          ...u,
          createdAt: Timestamp.now()
        };
        await setDoc(doc(db, 'users', u.uid), dbPayload);
        count++;
      }
      alert(`Sukses mengimpor ${count} pengguna baru secara masal ke database! 🎉`);
      setBulkInputText('');
    } catch (err) {
      handleFirestoreError(err, OperationType.WRITE, 'users');
    } finally {
      setIsBulkUploading(false);
    }
  };

  // Excel File Parsing Handler (.xlsx, .xls, .csv)
  const handleExcelFileSelect = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setExcelFileName(file.name);
    const reader = new FileReader();

    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const rawData = XLSX.utils.sheet_to_json<Record<string, any>>(worksheet, { defval: '' });

        if (rawData.length === 0) {
          alert('File Excel kosong atau tidak berisi data yang dapat dibaca.');
          return;
        }

        const parsedList: UserProfile[] = [];

        rawData.forEach((row, idx) => {
          const findVal = (possibleKeys: string[]) => {
            for (const key of Object.keys(row)) {
              const cleanKey = key.trim().toLowerCase();
              if (possibleKeys.some(p => cleanKey.includes(p))) {
                return String(row[key]).trim();
              }
            }
            return '';
          };

          const name = findVal(['nama', 'name', 'siswa', 'guru', 'fullname']) || `Pengguna ${idx + 1}`;
          let email = findVal(['email', 'surel', 'mail']) || '';
          const rawRole = findVal(['peran', 'role', 'jabatan', 'posisi']).toLowerCase();
          const kelas = findVal(['kelas', 'class', 'rombel', 'tingkat']) || 'SEMUA';

          let role: 'murid' | 'guru_wali' | 'guru_bk' | 'admin' = 'murid';
          if (rawRole.includes('guru') && (rawRole.includes('wali') || rawRole.includes('kelas'))) {
            role = 'guru_wali';
          } else if (rawRole.includes('bk') || rawRole.includes('konselor')) {
            role = 'guru_bk';
          } else if (rawRole.includes('admin')) {
            role = 'admin';
          } else if (rawRole.includes('guru')) {
            role = 'guru_wali';
          }

          if (!email) {
            const cleanNameSlug = name.toLowerCase().replace(/[^a-z0-9]/g, '');
            email = `${cleanNameSlug || 'user'}_${Math.random().toString(36).substring(2, 6)}@smp.belajar.id`;
          }

          const uid = 'user_' + Math.random().toString(36).substring(2, 12);

          parsedList.push({
            uid,
            email: email.toLowerCase(),
            displayName: name,
            photoURL: `https://api.dicebear.com/7.x/adventurer/svg?seed=${uid}`,
            role,
            kelas: role === 'murid' || role === 'guru_wali' ? (kelas.toUpperCase() || '7-A') : 'SEMUA',
            createdAt: Timestamp.now()
          });
        });

        setExcelParsedUsers(parsedList);
      } catch (err: any) {
        console.error('Error reading Excel file:', err);
        alert('Gagal membaca file Excel. Harap pastikan format file adalah .xlsx, .xls, atau .csv');
      }
    };

    reader.readAsBinaryString(file);
  };

  // Excel Template Downloader
  const handleDownloadExcelTemplate = () => {
    const templateData = [
      {
        "Nama Lengkap": "Ahmad Rizky Pratama",
        "Email": "ahmad.rizky@siswa.belajar.id",
        "Peran": "murid",
        "Kelas": "7-A"
      },
      {
        "Nama Lengkap": "Siti Nurhaliza",
        "Email": "siti.nurhaliza@siswa.belajar.id",
        "Peran": "murid",
        "Kelas": "7-A"
      },
      {
        "Nama Lengkap": "Budi Santoso, S.Pd.",
        "Email": "budi.santoso@guru.smp.belajar.id",
        "Peran": "guru_wali",
        "Kelas": "7-A"
      },
      {
        "Nama Lengkap": "Dra. Rina BK",
        "Email": "rina.bk@guru.smp.belajar.id",
        "Peran": "guru_bk",
        "Kelas": "SEMUA"
      }
    ];

    const worksheet = XLSX.utils.json_to_sheet(templateData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data_Pengguna");
    XLSX.writeFile(workbook, "Templat_Pendaftaran_Pengguna_CERDAS.xlsx");
  };

  // Import Excel Parsed Users to Firestore
  const handleImportExcelUsers = async () => {
    if (excelParsedUsers.length === 0) return;

    setExcelImporting(true);

    if (offlineMode) {
      const updatedUsers = { ...allUsers };
      excelParsedUsers.forEach(u => {
        updatedUsers[u.uid] = u;
      });
      setAllUsers(updatedUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(updatedUsers));
      alert(`[Offline Sandbox] Sukses mengimpor ${excelParsedUsers.length} pengguna baru dari file Excel! 🎉`);
      setExcelParsedUsers([]);
      setExcelFileName('');
      setExcelImporting(false);
      return;
    }

    try {
      let successCount = 0;
      for (const u of excelParsedUsers) {
        try {
          await setDoc(doc(db, 'users', u.uid), u);
          successCount++;
        } catch (e) {
          console.warn(`Gagal menyimpan user ${u.displayName}:`, e);
        }
      }

      const updatedUsers = { ...allUsers };
      excelParsedUsers.forEach(u => {
        updatedUsers[u.uid] = u;
      });
      setAllUsers(updatedUsers);

      alert(`🎉 Berhasil mengimpor ${successCount} dari ${excelParsedUsers.length} pengguna dari file Excel ke database Firestore!`);
      setExcelParsedUsers([]);
      setExcelFileName('');
    } catch (err) {
      console.error("Gagal mengimpor Excel ke Firestore:", err);
      alert("Terjadi kendala saat mengimpor data ke database. Data tersimpan di tampilan lokal.");
    } finally {
      setExcelImporting(false);
    }
  };

  // Delete individual user
  const handleDeleteUser = async (uid: string, name: string) => {
    if (!window.confirm(`Apakah Anda yakin ingin menghapus pengguna "${name}"?`)) return;

    if (offlineMode) {
      const updatedUsers = { ...allUsers };
      delete updatedUsers[uid];
      setAllUsers(updatedUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(updatedUsers));
      alert("[Offline Sandbox] Berhasil menghapus pengguna.");
      return;
    }

    try {
      const { deleteDoc } = await import('firebase/firestore');
      await deleteDoc(doc(db, 'users', uid));
      alert("Pengguna berhasil dihapus.");
    } catch (err) {
      console.warn("Firestore delete blocked, deleting locally from view", err);
      const updatedUsers = { ...allUsers };
      delete updatedUsers[uid];
      setAllUsers(updatedUsers);
    }
  };

  // Inline user update
  const handleUpdateUser = async (uid: string) => {
    if (!editUserName.trim() || !editUserEmail.trim()) {
      alert("Nama dan Email tidak boleh kosong!");
      return;
    }

    const cleanKelas = editUserRole === 'murid' || editUserRole === 'guru_wali' ? editUserKelas.trim().toUpperCase() : 'SEMUA';
    const updatedPayload = {
      ...allUsers[uid],
      displayName: editUserName.trim(),
      email: editUserEmail.trim().toLowerCase(),
      role: editUserRole,
      kelas: cleanKelas
    };

    if (offlineMode) {
      const updatedUsers = { ...allUsers, [uid]: updatedPayload };
      setAllUsers(updatedUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(updatedUsers));
      setEditingUserUid(null);
      alert("[Offline Sandbox] Berhasil memperbarui data pengguna!");
      return;
    }

    try {
      await setDoc(doc(db, 'users', uid), updatedPayload);
      setEditingUserUid(null);
      alert("Berhasil memperbarui data pengguna!");
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, `users/${uid}`);
    }
  };

  // Delete ALL roster users (excluding current admin user)
  const handleDeleteAllUsers = async () => {
    if (!window.confirm("⚠️ PERINGATAN KRITIS: Apakah Anda yakin ingin menghapus seluruh data roster guru wali dan murid? Tindakan ini akan mengosongkan semua daftar pengguna.")) return;
    if (!window.confirm("Konfirmasi terakhir: Anda benar-benar yakin? Semua cerita kelas akan kehilangan asosiasi roster.")) return;

    const selfUid = currentUser?.uid;
    const keptUsers: { [uid: string]: UserProfile } = {};
    if (selfUid && allUsers[selfUid]) {
      keptUsers[selfUid] = allUsers[selfUid];
    }

    if (offlineMode) {
      setAllUsers(keptUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(keptUsers));
      alert("[Offline Sandbox] Seluruh roster pengguna berhasil dibersihkan!");
      return;
    }

    try {
      const uidsToDelete = Object.keys(allUsers).filter(uid => uid !== selfUid);
      const { deleteDoc } = await import('firebase/firestore');
      
      for (const uid of uidsToDelete) {
        try {
          await deleteDoc(doc(db, 'users', uid));
        } catch (e) {
          // If restricted, overwrite to placeholder
          await setDoc(doc(db, 'users', uid), {
            ...allUsers[uid],
            role: 'murid',
            displayName: '[Dihapus]',
            kelas: 'SEMUA'
          });
        }
      }
      alert("Seluruh roster pengguna berhasil dibersihkan!");
    } catch (err) {
      console.warn("Gagal membersihkan roster online, menyalin perubahan ke lokal", err);
      setAllUsers(keptUsers);
    }
  };

  // Delete all Guru Wali
  const handleDeleteAllGuruWali = async () => {
    if (!window.confirm("⚠️ PERINGATAN: Apakah Anda yakin ingin menghapus SELURUH data Guru Wali?")) return;
    
    const selfUid = currentUser?.uid;
    const keptUsers: { [uid: string]: UserProfile } = {};
    Object.entries(allUsers).forEach(([uid, u]) => {
      if (u.role !== 'guru_wali' || uid === selfUid) {
        keptUsers[uid] = u;
      }
    });

    if (offlineMode) {
      setAllUsers(keptUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(keptUsers));
      alert("[Offline Sandbox] Seluruh data Guru Wali berhasil dihapus!");
      return;
    }

    try {
      const uidsToDelete = Object.keys(allUsers).filter(uid => allUsers[uid].role === 'guru_wali' && uid !== selfUid);
      const { deleteDoc } = await import('firebase/firestore');
      for (const uid of uidsToDelete) {
        await deleteDoc(doc(db, 'users', uid));
      }
      alert("Seluruh data Guru Wali berhasil dihapus!");
    } catch (err) {
      console.warn("Gagal menghapus online, menyalin lokal", err);
      setAllUsers(keptUsers);
    }
  };

  // Delete all Murid
  const handleDeleteAllMurid = async () => {
    if (!window.confirm("⚠️ PERINGATAN: Apakah Anda yakin ingin menghapus SELURUH data Murid?")) return;
    
    const selfUid = currentUser?.uid;
    const keptUsers: { [uid: string]: UserProfile } = {};
    Object.entries(allUsers).forEach(([uid, u]) => {
      if (u.role !== 'murid' || uid === selfUid) {
        keptUsers[uid] = u;
      }
    });

    if (offlineMode) {
      setAllUsers(keptUsers);
      localStorage.setItem('cerdas_users', JSON.stringify(keptUsers));
      alert("[Offline Sandbox] Seluruh data Murid berhasil dihapus!");
      return;
    }

    try {
      const uidsToDelete = Object.keys(allUsers).filter(uid => allUsers[uid].role === 'murid' && uid !== selfUid);
      const { deleteDoc } = await import('firebase/firestore');
      for (const uid of uidsToDelete) {
        await deleteDoc(doc(db, 'users', uid));
      }
      alert("Seluruh data Murid berhasil dihapus!");
    } catch (err) {
      console.warn("Gagal menghapus online, menyalin lokal", err);
      setAllUsers(keptUsers);
    }
  };

  // Delete individual story
  const handleDeleteStory = async (storyId: string) => {
    if (!window.confirm("Apakah Anda yakin ingin menghapus cerita ini?")) return;

    if (offlineMode) {
      const updated = stories.filter(s => s.id !== storyId);
      setStories(updated);
      localStorage.setItem('cerdas_stories', JSON.stringify(updated));
      alert("[Offline Sandbox] Berhasil menghapus cerita.");
      return;
    }

    try {
      const { deleteDoc } = await import('firebase/firestore');
      await deleteDoc(doc(db, 'stories', storyId));
      alert("Cerita berhasil dihapus.");
    } catch (err) {
      console.warn("Firestore delete blocked, deleting locally from view", err);
      setStories(prev => prev.filter(s => s.id !== storyId));
    }
  };

  // Update story
  const handleUpdateStory = async (storyId: string) => {
    const cleanTitle = editStoryTitle.trim();
    const cleanContent = editStoryContent.trim();

    if (cleanTitle.length < 3) {
      alert("⚠️ Judul cerita minimal 3 karakter!");
      return;
    }

    if (cleanContent.length < 30) {
      alert(`⚠️ Isi cerita terlalu singkat (${cleanContent.length} karakter)! Minimal harus 30 karakter agar dapat tersimpan.`);
      return;
    }

    if (cleanContent.length > 5000) {
      alert(`⚠️ Isi cerita melebihi batas (${cleanContent.length} karakter)! Maksimal 5000 karakter.`);
      return;
    }

    const updated = stories.map(s => {
      if (s.id === storyId) {
        return {
          ...s,
          title: editStoryTitle.trim(),
          content: editStoryContent.trim(),
          status: editStoryStatus,
          authorKelas: editStoryKelas.trim().toUpperCase(),
          updatedAt: { seconds: Date.now() / 1000 }
        };
      }
      return s;
    });

    if (offlineMode) {
      setStories(updated);
      localStorage.setItem('cerdas_stories', JSON.stringify(updated));
      setEditingStoryId(null);
      alert("[Offline Sandbox] Berhasil memperbarui cerita!");
      return;
    }

    try {
      await updateDoc(doc(db, 'stories', storyId), {
        title: editStoryTitle.trim(),
        content: editStoryContent.trim(),
        status: editStoryStatus,
        authorKelas: editStoryKelas.trim().toUpperCase(),
        updatedAt: Timestamp.now()
      });
      setEditingStoryId(null);
      alert("Berhasil memperbarui cerita!");
    } catch (err) {
      handleFirestoreError(err, OperationType.UPDATE, `stories/${storyId}`);
    }
  };

  // Delete all stories
  const handleDeleteAllStories = async () => {
    if (!window.confirm("⚠️ PERINGATAN: Apakah Anda yakin ingin menghapus SELURUH cerita di dalam sistem?")) return;

    if (offlineMode) {
      setStories([]);
      localStorage.setItem('cerdas_stories', JSON.stringify([]));
      alert("[Offline Sandbox] Seluruh cerita berhasil dihapus!");
      return;
    }

    try {
      const { deleteDoc } = await import('firebase/firestore');
      for (const story of stories) {
        await deleteDoc(doc(db, 'stories', story.id));
      }
      setStories([]);
      alert("Seluruh cerita berhasil dihapus!");
    } catch (err) {
      console.warn("Gagal menghapus online, menyalin lokal", err);
      setStories([]);
    }
  };

  // Offline Sandbox role hot swapper
  const handleHotSwapRole = (role: 'guru_wali' | 'guru_non_wali' | 'murid' | 'guru_bk' | 'admin') => {
    if (!userProfile) return;
    const updated = {
      ...userProfile,
      role,
      kelas: role === 'murid' || role === 'guru_wali' ? 'VII-A' : 'SEMUA'
    };
    setUserProfile(updated);
    
    // Auto route tab
    if (role === 'murid') {
      setActiveTab('gallery');
    } else if (role === 'guru_wali') {
      setActiveTab('classroom');
    } else if (role === 'guru_bk') {
      setActiveTab('bk_corner');
    } else if (role === 'admin') {
      setActiveTab('admin');
    } else {
      setActiveTab('gallery');
    }
  };

  // Loading Screen
  if (loading) {
    return (
      <div className="min-h-screen bg-[#F0F8FA] flex flex-col items-center justify-center p-6">
        <div className="w-16 h-16 border-4 border-amber-400 border-t-transparent rounded-full animate-spin"></div>
        <p className="mt-4 text-slate-600 font-medium text-lg animate-pulse">Menghubungkan ke CERDAS Digital...</p>
      </div>
    );
  }

  // Login Screen
  if (!currentUser) {
    return (
      <div className="min-h-screen bg-[#F8FBFC] flex flex-col lg:flex-row items-center justify-center p-4 lg:p-12 gap-8 lg:gap-16">
        {/* Left Art Cover */}
        <div className="flex-1 max-w-lg text-center lg:text-left space-y-6">
          <div className="inline-flex items-center gap-2 px-3 py-1 bg-amber-100 text-amber-800 rounded-full text-xs font-semibold uppercase tracking-wider">
            <Sparkles className="w-4 h-4 text-amber-500 animate-spin" /> Literasi Anak Indonesia Hebat
          </div>
          <h1 className="text-4xl lg:text-6xl font-extrabold text-teal-800 tracking-tight leading-none">
            CERDAS
          </h1>
          <p className="text-xl font-display text-slate-600">
            Cerita Digital Anak Sempatik
          </p>
          <p className="text-slate-500 text-base leading-relaxed">
            Wadah interaktif menyenangkan untuk membantu anak Indonesia bercerita secara cerdas dengan metode refleksi <strong className="text-teal-700">4F (Fact, Feeling, Finding, Future)</strong> serta mengamalkan <strong className="text-amber-600">7 Kebiasaan Anak Indonesia Hebat</strong>.
          </p>

          <div className="grid grid-cols-2 gap-4 pt-4 text-left">
            <div className="p-4 bg-white rounded-xl shadow-xs border border-slate-100">
              <span className="text-2xl">🌱</span>
              <h3 className="font-bold text-slate-800 mt-1">Reflektif (4F)</h3>
              <p className="text-xs text-slate-500 mt-1">Peristiwa, Perasaan, Pembelajaran & Penerapan Masa Depan.</p>
            </div>
            <div className="p-4 bg-white rounded-xl shadow-xs border border-slate-100">
              <span className="text-2xl">🏆</span>
              <h3 className="font-bold text-slate-800 mt-1">7 Kebiasaan</h3>
              <p className="text-xs text-slate-500 mt-1">Ibadah, olahraga, makan sehat, gemar belajar, bangun pagi, dsb.</p>
            </div>
          </div>
        </div>

        {/* Right Authentication Form */}
        <div className="w-full max-w-md bg-white p-8 rounded-3xl shadow-xl border border-teal-50/50 flex flex-col items-center justify-center text-center space-y-6 relative overflow-hidden">
          <div className="absolute top-0 left-0 w-full h-3 bg-gradient-to-r from-teal-400 via-amber-300 to-rose-400"></div>
          
          <div className="p-4 bg-teal-50 rounded-2xl">
            <BookOpen className="w-12 h-12 text-teal-600" />
          </div>

          <div className="space-y-2">
            <h2 className="text-2xl font-bold text-slate-800">Mulai Petualanganmu!</h2>
            <p className="text-sm text-slate-500">Gunakan akun Google sekolah atau akun pribadi Anda</p>
          </div>

          <button 
            onClick={handleLogin}
            className="w-full flex items-center justify-center gap-3 bg-white hover:bg-slate-50 text-slate-800 border border-slate-200 px-6 py-3.5 rounded-2xl font-semibold shadow-sm hover:shadow-md transition-all active:scale-95 duration-200"
          >
            <svg className="w-5 h-5 shrink-0" viewBox="0 0 24 24" fill="none">
              <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#4285F4" />
              <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853" />
              <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.06H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.94l3.66-2.85z" fill="#FBBC05" />
              <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.06l3.66 2.85c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335" />
            </svg>
            Masuk dengan Google
          </button>

          <p className="text-xs text-slate-400">
            Aman • Terintegrasi dengan Google AI Studio & Firebase
          </p>
        </div>
      </div>
    );
  }

  // Registration & Role Selection (First login)
  if (currentUser && !userProfile) {
    return (
      <div className="min-h-screen bg-[#F0F8FA] flex items-center justify-center p-4">
        <div className="w-full max-w-xl bg-white rounded-3xl shadow-xl border border-teal-100 overflow-hidden relative">
          <div className="absolute top-0 left-0 w-full h-3 bg-gradient-to-r from-amber-400 via-emerald-400 to-blue-500"></div>
          
          <div className="p-8 space-y-6">
            <div className="text-center space-y-2">
              <h2 className="text-3xl font-extrabold text-teal-800">Lengkapi Profil CERDAS</h2>
              <p className="text-sm text-slate-500">Halo {currentUser.displayName}! Tentukan peran Anda untuk mulai berselancar di perpustakaan literasi digital.</p>
            </div>

            <form onSubmit={handleRegisterProfile} className="space-y-6">
              {/* Role Selection */}
              <div className="space-y-3">
                <label className="block text-sm font-bold text-slate-700">Pilih Peran Anda:</label>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <button
                    type="button"
                    onClick={() => { setRegRole('murid'); }}
                    className={`p-4 rounded-2xl border text-left flex flex-col justify-between transition-all duration-200 ${regRole === 'murid' ? 'border-teal-500 bg-teal-50/50 ring-2 ring-teal-200' : 'border-slate-200 hover:border-slate-300'}`}
                  >
                    <span className="text-2xl mb-2">🎒</span>
                    <div>
                      <h4 className="font-bold text-slate-800 text-sm">Murid / Siswa</h4>
                      <p className="text-xs text-slate-500 mt-0.5">Menulis cerita refleksi 4F & mencatat kebiasaan baik harian.</p>
                    </div>
                  </button>

                  <button
                    type="button"
                    onClick={() => { setRegRole('guru_wali'); }}
                    className={`p-4 rounded-2xl border text-left flex flex-col justify-between transition-all duration-200 ${regRole === 'guru_wali' ? 'border-amber-500 bg-amber-50/50 ring-2 ring-amber-200' : 'border-slate-200 hover:border-slate-300'}`}
                  >
                    <span className="text-2xl mb-2">👩‍🏫</span>
                    <div>
                      <h4 className="font-bold text-slate-800 text-sm">Guru Wali</h4>
                      <p className="text-xs text-slate-500 mt-0.5">Memantau progres kelas, membaca draf cerita & memberi ulasan.</p>
                    </div>
                  </button>

                  <button
                    type="button"
                    onClick={() => { setRegRole('guru_bk'); }}
                    className={`p-4 rounded-2xl border text-left flex flex-col justify-between transition-all duration-200 ${regRole === 'guru_bk' ? 'border-rose-500 bg-rose-50/50 ring-2 ring-rose-200' : 'border-slate-200 hover:border-slate-300'}`}
                  >
                    <span className="text-2xl mb-2">❤️</span>
                    <div>
                      <h4 className="font-bold text-slate-800 text-sm">Guru BK (Konselor)</h4>
                      <p className="text-xs text-slate-500 mt-0.5">Memantau kesejahteraan emosional murid lewat cerita.</p>
                    </div>
                  </button>
                </div>
              </div>

              {/* Class Designation (Only for Murid and Guru Wali) */}
              {(regRole === 'murid' || regRole === 'guru_wali') && (
                <div className="space-y-2 animate-fadeIn">
                  <label className="block text-sm font-bold text-slate-700">
                    {regRole === 'guru_wali' ? 'Kelas Binaan Guru Wali (Boleh >1 Kelas):' : 'Tingkat Kelas & Ruang:'}
                  </label>
                  <input
                    type="text"
                    required
                    placeholder={regRole === 'guru_wali' ? "Contoh: 7B, 7C (Pisahkan koma jika >1 kelas)" : "Contoh: VII-A atau 7B"}
                    value={regKelas}
                    onChange={(e) => setRegKelas(e.target.value)}
                    className="w-full border border-slate-200 px-4 py-3 rounded-2xl focus:outline-none focus:ring-2 focus:ring-teal-400 focus:border-transparent text-sm"
                  />
                  <p className="text-xs text-slate-400">
                    {regRole === 'guru_wali'
                      ? '💡 Guru wali yang membina lebih dari 1 kelas dapat mengetik kelas dipisahkan koma (contoh: 7A, 7B atau 7B, 7C).'
                      : 'Kelas akan mengelompokkan murid dengan guru walinya secara otomatis.'}
                  </p>
                </div>
              )}

              <button
                type="submit"
                disabled={profileLoading}
                className="w-full bg-teal-600 hover:bg-teal-700 text-white font-bold py-3.5 px-6 rounded-2xl shadow-md transition-all active:scale-98 disabled:opacity-50"
              >
                {profileLoading ? 'Mendaftarkan...' : 'Masuk ke Aplikasi CERDAS 🚀'}
              </button>
            </form>
          </div>
        </div>
      </div>
    );
  }

  // Active Story filtering in gallery
  const filteredStories = stories.filter(story => {
    const matchSearch = story.title.toLowerCase().includes(searchQuery.toLowerCase()) || 
                        story.content.toLowerCase().includes(searchQuery.toLowerCase()) ||
                        story.authorName.toLowerCase().includes(searchQuery.toLowerCase());
    
    const matchKelas = filterKelas === 'Semua' || story.authorKelas === filterKelas;
    const matchHabit = filterHabit === 'Semua' || story.associatedHabits.includes(filterHabit);
    const matchEmotion = filterEmotion === 'Semua' || story.feelingEmoji === filterEmotion;

    const isAuthor = story.authorId === currentUser?.uid;
    const isVisible = story.status !== 'draft' || isAuthor;

    return matchSearch && matchKelas && matchHabit && matchEmotion && isVisible;
  });

  const uniqueClasses = Array.from(new Set(stories.map(s => s.authorKelas))).filter(Boolean);

  return (
    <div className="min-h-screen bg-[#F8FBFC] flex flex-col">
      
      {/* ⚠️ HIGH-VISIBILITY FIRESTORE QUOTA WARNING BANNER */}
      {isQuotaExceeded && (
        <div className="bg-gradient-to-r from-amber-500 via-orange-600 to-rose-600 text-white px-6 py-4 shadow-md flex flex-col space-y-3 z-50">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="p-2 bg-white/20 rounded-xl">
                <Shield className="w-5 h-5 text-white animate-pulse" />
              </div>
              <div className="space-y-0.5">
                <h4 className="font-bold text-sm">Mode Sandbox Offline Aktif (Kuota Harian Firestore Tercapai)</h4>
                <p className="text-xs text-amber-50 leading-relaxed max-w-4xl">
                  Batas pembacaan gratis harian Spark Firestore pada akun ini telah terpenuhi. 
                  Jangan khawatir! Kami mengaktifkan **Mode Sandbox Offline (LocalStorage)** secara otomatis. 
                  Anda tetap dapat mencoba aplikasi ini secara penuh, menulis cerita, menganalisis dengan Gemini AI, dan menyimpan hasil ulasan di browser lokal Anda.
                </p>
              </div>
            </div>
            
            <a 
              href="https://console.firebase.google.com/project/gen-lang-client-0131415670/firestore/databases/ai-studio-cerdasceritadigi-37395f6c-764f-44b3-ad4e-8a1ee0649221/data?openUpgradeDialog=true" 
              target="_blank" 
              rel="noopener noreferrer"
              className="bg-white hover:bg-amber-50 text-orange-700 font-bold px-4 py-2 rounded-xl text-xs transition-colors self-start md:self-auto shadow-xs whitespace-nowrap inline-flex items-center gap-1.5"
            >
              <Zap className="w-3.5 h-3.5 text-orange-500" />
              Tingkatkan Quota / Upgrade
            </a>
          </div>

          {/* Super-Cool Multi-Role Hot Swapper to easily test different screens in Offline Sandbox */}
          <div className="pt-2 border-t border-white/20 flex flex-wrap items-center gap-3 text-xs">
            <span className="font-bold text-amber-100">🔌 Simulator Peran Demo:</span>
            <div className="flex flex-wrap gap-1 p-1 bg-black/25 rounded-xl">
              {[
                { role: 'murid', label: '🎒 Siswa (Ahmad)' },
                { role: 'guru_wali', label: '👩‍🏫 Guru Wali (Budi)' },
                { role: 'guru_bk', label: '❤️ Guru BK' },
                { role: 'admin', label: '🛡️ Admin' }
              ].map(swapper => (
                <button
                  key={swapper.role}
                  onClick={() => handleHotSwapRole(swapper.role as any)}
                  className={`px-3 py-1 rounded-lg font-bold transition-all ${userProfile?.role === swapper.role ? 'bg-white text-orange-800 shadow-sm' : 'text-amber-100 hover:text-white'}`}
                >
                  {swapper.label}
                </button>
              ))}
            </div>
          </div>
        </div>
      )}

      {/* 3-ZONE TOP BAR CONTRACT */}
      <header className="sticky top-0 z-40 bg-white border-b border-slate-100 px-6 py-4 flex items-center justify-between">
        {/* Zone 1: Single element brand wordmark */}
        <div className="flex items-center gap-2">
          <a href="#" onClick={() => setActiveTab('gallery')} className="text-xl md:text-2xl font-extrabold tracking-tight bg-gradient-to-r from-teal-600 to-amber-500 bg-clip-text text-transparent">
            CERDAS
          </a>
          <span className="inline-block text-xs font-semibold px-2.5 py-0.5 bg-teal-50 text-teal-700 border border-teal-100/80 rounded-full shadow-2xs">
            {userProfile?.role === 'murid' ? 'Siswa Sempatik Glory' : (userProfile?.role === 'guru_wali' ? `Guru Wali ${userProfile.kelas}` : (userProfile?.role === 'guru_bk' ? 'Guru BK / Konselor' : (userProfile?.role === 'admin' ? 'Administrator' : 'Siswa Sempatik Glory')))}
          </span>
        </div>

        {/* Zone 2: Navigation Links (single-line, truncated on narrow devices) */}
        <nav className="flex items-center gap-1 md:gap-4 font-medium text-sm text-slate-600">
          <button 
            onClick={() => setActiveTab('gallery')}
            className={`px-3 py-1.5 rounded-lg text-xs md:text-sm transition-colors whitespace-nowrap ${activeTab === 'gallery' ? 'bg-teal-50 text-teal-700 font-bold' : 'hover:text-slate-900'}`}
          >
            📚 Galeri Cerita
          </button>

          {userProfile?.role === 'murid' && (
            <>
              <button 
                onClick={() => setActiveTab('write')}
                className={`px-3 py-1.5 rounded-lg text-xs md:text-sm transition-colors whitespace-nowrap ${activeTab === 'write' ? 'bg-amber-50 text-amber-800 font-bold' : 'hover:text-slate-900'}`}
              >
                ✏️ Tulis Cerita
              </button>
              <button 
                onClick={() => setActiveTab('habits')}
                className={`px-3 py-1.5 rounded-lg text-xs md:text-sm transition-colors whitespace-nowrap ${activeTab === 'habits' ? 'bg-emerald-50 text-emerald-800 font-bold' : 'hover:text-slate-900'}`}
              >
                📅 Jurnal Kebiasaan
              </button>
            </>
          )}

          {userProfile?.role === 'guru_wali' && (
            <button 
              onClick={() => setActiveTab('classroom')}
              className={`px-3 py-1.5 rounded-lg text-xs md:text-sm transition-colors whitespace-nowrap ${activeTab === 'classroom' ? 'bg-amber-50 text-amber-800 font-bold' : 'hover:text-slate-900'}`}
            >
              👩‍🏫 Kelasku ({userProfile.kelas})
            </button>
          )}

          {userProfile?.role === 'guru_bk' && (
            <button 
              onClick={() => setActiveTab('bk_corner')}
              className={`px-3 py-1.5 rounded-lg text-xs md:text-sm transition-colors whitespace-nowrap ${activeTab === 'bk_corner' ? 'bg-rose-50 text-rose-800 font-bold' : 'hover:text-slate-900'}`}
            >
              ❤️ Sudut Konseling BK
            </button>
          )}

          {(userProfile?.role === 'admin' || currentUser?.email === 'isumayasa91@guru.smp.belajar.id' || offlineMode) && (
            <button 
              onClick={() => setActiveTab('admin')}
              className={`px-3 py-1.5 rounded-lg text-xs md:text-sm transition-colors whitespace-nowrap ${activeTab === 'admin' ? 'bg-indigo-50 text-indigo-800 font-bold' : 'hover:text-slate-900'}`}
            >
              🛡️ Menu Admin
            </button>
          )}
        </nav>

        {/* Zone 3: Account / Primary Actions */}
        <div className="flex items-center gap-3">
          <div className="hidden md:flex flex-col text-right truncate max-w-[150px]">
            <span className="text-xs font-bold text-slate-800 leading-tight truncate">
              {userProfile?.displayName || currentUser.displayName}
            </span>
            <span className="text-[10px] text-slate-400 truncate">
              {currentUser.email}
            </span>
          </div>

          <img 
            src={userProfile?.photoURL || currentUser.photoURL || 'https://api.dicebear.com/7.x/adventurer/svg'} 
            alt="Profil" 
            className="w-8 h-8 rounded-full ring-2 ring-teal-100"
          />

          <button 
            onClick={handleLogout}
            title="Keluar"
            className="p-1.5 hover:bg-slate-100 rounded-lg text-slate-400 hover:text-rose-500 transition-colors"
          >
            <LogOut className="w-4 h-4" />
          </button>
        </div>
      </header>

      {/* Main Container */}
      <main className="flex-1 max-w-7xl w-full mx-auto p-4 md:p-6 space-y-6">
        
        {/* ======================================= */}
        {/* TAB 1: STORY GALLERY / Digital Bookshelf */}
        {/* ======================================= */}
        {activeTab === 'gallery' && (
          <div className="space-y-6">
            
            {/* Jumbotron/Banner */}
            <div className="relative overflow-hidden bg-gradient-to-r from-teal-700 to-emerald-600 rounded-3xl p-6 md:p-8 text-white shadow-lg">
              <div className="relative z-10 max-w-xl space-y-3">
                <div className="inline-flex items-center gap-1 px-2.5 py-0.5 bg-teal-500/30 rounded-full text-xs font-semibold">
                  📖 Pojok Baca Digital Anak
                </div>
                <h1 className="text-3xl md:text-4xl font-extrabold tracking-tight">
                  Eksplorasi Dunia Imajinasi & Kebiasaan Hebat!
                </h1>
                <p className="text-sm md:text-base text-teal-100 leading-relaxed">
                  Temukan cerita inspiratif teman-teman se-Indonesia yang ditulis berdasarkan pengalaman nyata dengan bimbingan metode refleksi CERDAS 4F.
                </p>
              </div>
              <div className="absolute right-0 bottom-0 opacity-10 pointer-events-none transform translate-y-6 translate-x-4">
                <BookOpen className="w-96 h-96" />
              </div>
            </div>

            {/* Gallery Controls & Search */}
            <div className="bg-white p-4 rounded-2xl shadow-xs border border-slate-100 flex flex-col md:flex-row gap-3 items-center justify-between">
              
              <div className="relative w-full md:w-96">
                <Search className="absolute left-3.5 top-1/2 -translate-y-1/2 text-slate-400 w-4 h-4" />
                <input
                  type="text"
                  placeholder="Cari judul cerita, nama murid, atau isi kisah..."
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  className="w-full bg-slate-50 border-0 pl-10 pr-4 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-sm"
                />
              </div>

              {/* Filters */}
              <div className="flex flex-wrap items-center gap-3 w-full md:w-auto">
                <div className="flex items-center gap-1.5">
                  <span className="text-xs font-semibold text-slate-400">Kelas:</span>
                  <select
                    value={filterKelas}
                    onChange={(e) => setFilterKelas(e.target.value)}
                    className="bg-slate-50 border border-slate-200 rounded-lg text-xs px-2.5 py-1.5 text-slate-600 font-medium focus:outline-none"
                  >
                    <option value="Semua">Semua Kelas</option>
                    <option value="VII-A">Kelas VII-A</option>
                    <option value="VIII-B">Kelas VIII-B</option>
                    <option value="IX-C">Kelas IX-C</option>
                    {uniqueClasses.filter(c => c !== 'VII-A' && c !== 'VIII-B' && c !== 'IX-C').map(c => (
                      <option key={c} value={c}>{c}</option>
                    ))}
                  </select>
                </div>

                <div className="flex items-center gap-1.5">
                  <span className="text-xs font-semibold text-slate-400">Kebiasaan:</span>
                  <select
                    value={filterHabit}
                    onChange={(e) => setFilterHabit(e.target.value)}
                    className="bg-slate-50 border border-slate-200 rounded-lg text-xs px-2.5 py-1.5 text-slate-600 font-medium focus:outline-none"
                  >
                    <option value="Semua">Semua Kebiasaan</option>
                    {HABITS_INFO.map(h => (
                      <option key={h.id} value={h.id}>{h.label}</option>
                    ))}
                  </select>
                </div>

                <div className="flex items-center gap-1.5">
                  <span className="text-xs font-semibold text-slate-400">Emosi:</span>
                  <select
                    value={filterEmotion}
                    onChange={(e) => setFilterEmotion(e.target.value)}
                    className="bg-slate-50 border border-slate-200 rounded-lg text-xs px-2.5 py-1.5 text-slate-600 font-medium focus:outline-none"
                  >
                    <option value="Semua">Semua Emosi</option>
                    {FEELING_EMOJIS.map(fe => (
                      <option key={fe.emoji} value={fe.emoji}>{fe.emoji} {fe.label}</option>
                    ))}
                  </select>
                </div>

                {(filterKelas !== 'Semua' || filterHabit !== 'Semua' || filterEmotion !== 'Semua' || searchQuery) && (
                  <button 
                    onClick={() => {
                      setFilterKelas('Semua');
                      setFilterHabit('Semua');
                      setFilterEmotion('Semua');
                      setSearchQuery('');
                    }}
                    className="text-xs text-rose-500 font-bold hover:underline"
                  >
                    Atur Ulang
                  </button>
                )}
              </div>
            </div>

            {/* Bookshelf Grid */}
            {filteredStories.length === 0 ? (
              <div className="bg-white rounded-3xl p-12 text-center border border-dashed border-slate-200 space-y-3">
                <div className="text-4xl">📚🔍</div>
                <h3 className="font-bold text-slate-700 text-lg">Belum ada cerita digital yang cocok</h3>
                <p className="text-sm text-slate-400 max-w-md mx-auto">Coba ganti filter di atas atau jadilah murid pertama yang membagikan cerita refleksi menarikmu hari ini!</p>
                {userProfile?.role === 'murid' && (
                  <button 
                    onClick={() => setActiveTab('write')}
                    className="bg-amber-500 hover:bg-amber-600 text-white font-bold px-4 py-2 rounded-xl text-xs transition-colors shadow-sm"
                  >
                    Tulis Cerita Sekarang ✏️
                  </button>
                )}
              </div>
            ) : (
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-6">
                {filteredStories.map((story) => {
                  const firstHabit = story.associatedHabits[0];
                  const fallbackBg = getCuteFallbackBackground(story.title);

                  return (
                    <div 
                      key={story.id} 
                      onClick={() => setSelectedStory(story)}
                      className="group cursor-pointer bg-white rounded-3xl shadow-xs hover:shadow-md border border-slate-100 overflow-hidden transition-all duration-200 flex flex-col justify-between"
                    >
                      <div className="aspect-[4/3] w-full relative overflow-hidden flex items-center justify-center bg-slate-100">
                        {story.illustrationUrl && story.illustrationUrl.startsWith('data:') ? (
                          <img 
                            src={story.illustrationUrl} 
                            alt={story.title} 
                            referrerPolicy="no-referrer"
                            className="w-full h-full object-cover group-hover:scale-105 transition-transform duration-300" 
                          />
                        ) : (
                          <div 
                            style={{ backgroundImage: fallbackBg }} 
                            className="w-full h-full flex flex-col items-center justify-center p-6 text-center text-white"
                          >
                            <span className="text-5xl mb-2 animate-bounce">
                              {HABITS_INFO.find(h => h.id === firstHabit)?.icon || '📖'}
                            </span>
                            <span className="text-sm font-semibold tracking-wide opacity-90 block">Cerita CERDAS</span>
                            <span className="text-xs font-medium opacity-75 mt-0.5">Metode Refleksi 4F</span>
                          </div>
                        )}
                        
                        <div className="absolute top-3 left-3 flex gap-2">
                          {story.status === 'draft' && (
                            <span className="bg-slate-100/90 text-slate-700 font-extrabold text-[9px] px-2 py-0.5 rounded-md uppercase tracking-wider backdrop-blur-xs">Draf</span>
                          )}
                          {story.status === 'submitted' && (
                            <span className="bg-amber-100/90 text-amber-800 font-extrabold text-[9px] px-2 py-0.5 rounded-md uppercase tracking-wider backdrop-blur-xs animate-pulse">Perlu Ulasan</span>
                          )}
                          {story.status === 'reviewed' && (
                            <span className="bg-emerald-100/90 text-emerald-800 font-extrabold text-[9px] px-2 py-0.5 rounded-md uppercase tracking-wider backdrop-blur-xs">Selesai Diulas</span>
                          )}
                        </div>

                        <div className="absolute bottom-3 right-3 bg-white/90 backdrop-blur-xs w-8 h-8 rounded-full flex items-center justify-center text-lg shadow-sm" title={story.feeling}>
                          {story.feelingEmoji || '😊'}
                        </div>
                      </div>

                      <div className="p-5 flex-1 flex flex-col justify-between space-y-4">
                        <div className="space-y-2">
                          <h3 className="font-bold text-slate-800 text-base leading-snug group-hover:text-teal-700 transition-colors line-clamp-2">
                            {story.title}
                          </h3>
                          <p className="text-xs text-slate-400 line-clamp-2 leading-relaxed">
                            {story.content}
                          </p>
                        </div>

                        <div className="pt-2 border-t border-slate-50 flex items-center justify-between text-[11px] text-slate-500">
                          <div className="flex items-center gap-1 truncate max-w-[140px]">
                            <span className="font-bold text-slate-700">{story.authorName}</span>
                            <span>·</span>
                            <span>Kelas {story.authorKelas}</span>
                          </div>
                          <div className="flex items-center gap-1.5">
                            <span>
                              {story.createdAt?.seconds ? new Date(story.createdAt.seconds * 1000).toLocaleDateString('id-ID', { day: 'numeric', month: 'short' }) : 'Hari ini'}
                            </span>
                            {(story.authorId === currentUser?.uid || userProfile?.role === 'admin' || userProfile?.role === 'guru_wali' || userProfile?.role === 'guru_bk' || offlineMode) && (
                              <button
                                type="button"
                                onClick={(e) => {
                                  e.stopPropagation();
                                  handleDeleteStory(story.id);
                                }}
                                title="Hapus Cerita"
                                className="p-1 hover:bg-rose-50 text-slate-400 hover:text-rose-600 rounded-lg transition-colors"
                              >
                                🗑️
                              </button>
                            )}
                          </div>
                        </div>
                      </div>
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        )}

        {/* ======================================= */}
        {/* TAB 2: WRITE STORY REFLECTION WIZARD */}
        {/* ======================================= */}
        {activeTab === 'write' && userProfile?.role === 'murid' && (
          <div className="max-w-3xl mx-auto space-y-6">
            
            <div className="space-y-2 text-center md:text-left">
              <h1 className="text-3xl font-extrabold text-teal-800 flex items-center justify-center md:justify-start gap-2">Tulis Cerita Digital ✏️</h1>
              <p className="text-slate-500 text-sm">Refleksikan kegiatan hebatmu menggunakan elemen 4F dan sambungkan dengan 7 Kebiasaan Baik.</p>
            </div>

            {/* Custom interactive Timeline */}
            <div className="bg-white p-4 rounded-2xl shadow-xs border border-slate-100 flex items-center justify-between">
              {[
                { step: 1, label: 'Kisah Utama' },
                { step: 2, label: 'Refleksi 4F' },
                { step: 3, label: '7 Kebiasaan Hebat' },
                { step: 4, label: 'Si CERDAS AI' }
              ].map((item, index, arr) => (
                <React.Fragment key={item.step}>
                  <div className="flex items-center gap-2">
                    <div className={`w-8 h-8 rounded-full flex items-center justify-center text-xs font-bold transition-all ${writeStep >= item.step ? 'bg-teal-600 text-white font-extrabold' : 'bg-slate-100 text-slate-400'}`}>
                      {item.step}
                    </div>
                    <span className={`text-xs font-bold hidden md:inline ${writeStep === item.step ? 'text-teal-700' : 'text-slate-400'}`}>
                      {item.label}
                    </span>
                  </div>
                  {index < arr.length - 1 && (
                    <div className={`flex-1 h-0.5 mx-2 ${writeStep > item.step ? 'bg-teal-600' : 'bg-slate-100'}`} />
                  )}
                </React.Fragment>
              ))}
            </div>

            {/* Wizard Body Card */}
            <div className="bg-white rounded-3xl p-6 md:p-8 shadow-xs border border-slate-100 space-y-6">
              
              {/* STEP 1: Main Story */}
              {writeStep === 1 && (
                <div className="space-y-6">
                  
                  {/* Photo Upload & AI Auto-Story Teller Section */}
                  <div className="bg-indigo-50/50 p-5 rounded-3xl border border-indigo-100/60 space-y-4 animate-fadeIn">
                    <div className="flex items-center gap-2 text-indigo-900 font-bold text-sm">
                      <Sparkles className="w-4 h-4 text-indigo-500 animate-spin" />
                      <span>📸 Buat Cerita Otomatis dari Foto Kegiatan! (Opsional)</span>
                    </div>
                    <p className="text-xs text-slate-500 leading-relaxed">
                      Punya foto sedang merapikan tempat tidur, berolahraga, makan buah sehat, atau belajar? 
                      Unggah fotomu di bawah ini dan biarkan <strong>Si CERDAS AI</strong> menganalisis foto tersebut untuk membuat draf cerita beserta ulasan refleksi 4F lengkap secara otomatis!
                    </p>
                    
                    <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-3">
                      <input
                        type="file"
                        accept="image/*"
                        onChange={handleImageUpload}
                        className="block w-full text-xs text-slate-500 file:mr-4 file:py-2 file:px-4 file:rounded-xl file:border-0 file:text-xs file:font-semibold file:bg-indigo-50 file:text-indigo-700 file:cursor-pointer hover:file:bg-indigo-100"
                      />
                      
                      {uploadedImageUrl && (
                        <button
                          type="button"
                          disabled={analyzingImage}
                          onClick={handleGenerateStoryFromImage}
                          className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-5 py-2.5 rounded-xl shadow-xs transition-colors flex items-center justify-center gap-1.5 disabled:opacity-50 min-h-[38px]"
                        >
                          {analyzingImage ? (
                            <>
                              <RefreshCw className="w-4 h-4 animate-spin" />
                              Menganalisis Foto...
                            </>
                          ) : (
                            <>
                              <Sparkles className="w-4 h-4" />
                              Buat Cerita Otomatis dari Foto!
                            </>
                          )}
                        </button>
                      )}
                    </div>

                    {uploadedImageUrl && (
                      <div className="flex items-center gap-3 animate-fadeIn pt-1">
                        <div className="relative w-20 aspect-square rounded-2xl overflow-hidden border border-slate-200 shrink-0">
                          <img src={uploadedImageUrl} alt="Unggahan Kegiatan" className="w-full h-full object-cover" />
                          <button
                            type="button"
                            onClick={() => setUploadedImageUrl(null)}
                            className="absolute top-1 right-1 bg-black/60 hover:bg-black/80 text-white p-1 rounded-full text-[8px]"
                          >
                            Hapus
                          </button>
                        </div>
                        <div className="text-[11px] text-indigo-700 font-semibold bg-indigo-50/50 p-2 rounded-xl">
                          ✓ Foto berhasil terunggah! Klik tombol <strong>"Buat Cerita Otomatis"</strong> di atas jika ingin AI menyusun draf cerita dari foto ini.
                        </div>
                      </div>
                    )}
                  </div>

                  <div className="bg-amber-50 p-4 rounded-2xl text-xs text-amber-900 border border-amber-100 leading-relaxed space-y-1">
                    <div><strong>💡 Tips Menulis:</strong> Mulailah dengan judul menarik dan tulis kisah pengalaman nyatamu secara jujur.</div>
                    <div className="text-[11px] font-semibold text-amber-800">
                      ⚠️ <strong>Ketentuan Panjang Cerita:</strong> Setiap cerita minimal berisi <strong>30 karakter</strong> dan maksimal <strong>5000 karakter</strong>. Jika terlalu singkat (kurang dari 30 karakter), sistem tidak akan menerima atau menyimpannya.
                    </div>
                  </div>

                  {/* Title Field with Voice note */}
                  <div className="space-y-2">
                    <label className="block text-sm font-bold text-slate-700 flex items-center justify-between">
                      <span>Judul Cerita:</span>
                      <div className="flex gap-2">
                        {isRecording === 'title' && (
                          <span className="text-[10px] text-red-500 font-bold animate-pulse flex items-center gap-1">🔴 Mendengarkan...</span>
                        )}
                        {isAiRecording === 'title' && (
                          <span className="text-[10px] text-indigo-600 font-bold animate-pulse flex items-center gap-1">🎙️ Merekam (Gemini AI)...</span>
                        )}
                        {transcribingAudio && (
                          <span className="text-[10px] text-teal-600 font-bold animate-pulse flex items-center gap-1">🧠 Memproses dengan AI...</span>
                        )}
                      </div>
                    </label>
                    <div className="flex items-center gap-2">
                      <input
                        type="text"
                        placeholder="Contoh: Belajar Berbagi Sarapan Sehat bersama Sahabat"
                        value={storyTitle}
                        onChange={(e) => setStoryTitle(e.target.value)}
                        className="flex-1 border border-slate-200 px-4 py-3 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-sm font-medium text-slate-700"
                      />
                      <div className="flex items-center gap-1 shrink-0">
                        <button
                          type="button"
                          title="Dikte Suara Langsung (Web Speech)"
                          onClick={() => startSpeechRecognition('title')}
                          className={`p-3 rounded-xl border transition-all ${isRecording === 'title' ? 'bg-red-100 text-red-600 border-red-300 animate-pulse' : 'bg-slate-50 text-slate-500 hover:bg-slate-100 border-slate-200'}`}
                        >
                          <Mic className="w-4 h-4" />
                        </button>
                        <button
                          type="button"
                          title={isAiRecording === 'title' ? "Selesai merekam & kirim ke Gemini AI" : "Rekam Suara & Transkripsi dengan Gemini AI"}
                          onClick={isAiRecording === 'title' ? stopMediaRecording : () => startMediaRecording('title')}
                          className={`p-2.5 rounded-xl border transition-all flex items-center gap-1 text-[11px] font-bold ${isAiRecording === 'title' ? 'bg-indigo-600 text-white border-indigo-500 animate-pulse' : 'bg-indigo-50 text-indigo-600 hover:bg-indigo-100 border-indigo-200'}`}
                        >
                          <Sparkles className="w-3.5 h-3.5" />
                          {isAiRecording === 'title' ? 'Selesai ⏺️' : 'Rekam AI'}
                        </button>
                      </div>
                    </div>
                  </div>

                  {/* Content Field with Voice note */}
                  <div className="space-y-2">
                    <label className="block text-sm font-bold text-slate-700 flex items-center justify-between">
                      <span>Kisah Utamamu (Cerita Bebas):</span>
                      <div className="flex gap-2">
                        {isRecording === 'content' && (
                          <span className="text-[10px] text-red-500 font-bold animate-pulse flex items-center gap-1">🔴 Mendengarkan...</span>
                        )}
                        {isAiRecording === 'content' && (
                          <span className="text-[10px] text-indigo-600 font-bold animate-pulse flex items-center gap-1">🎙️ Merekam (Gemini AI)...</span>
                        )}
                        {transcribingAudio && (
                          <span className="text-[10px] text-teal-600 font-bold animate-pulse flex items-center gap-1">🧠 Memproses dengan AI...</span>
                        )}
                      </div>
                    </label>
                    <div className="relative">
                      <textarea
                        rows={6}
                        placeholder="Pagi ini saya membawa bekal dua buah apel dari rumah. Saat jam istirahat, saya melihat Dika tidak membawa bekal makanan..."
                        value={storyContent}
                        onChange={(e) => setStoryContent(e.target.value)}
                        className="w-full border border-slate-200 p-4 pr-12 pb-14 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-sm leading-relaxed font-medium text-slate-600"
                      />
                      <div className="absolute bottom-3 right-3 flex items-center gap-1.5">
                        <button
                          type="button"
                          title="Dikte Suara Langsung (Web Speech)"
                          onClick={() => startSpeechRecognition('content')}
                          className={`p-2.5 rounded-xl border transition-all ${isRecording === 'content' ? 'bg-red-100 text-red-600 border-red-300 animate-pulse' : 'bg-slate-50 text-slate-400 hover:bg-slate-100 border-slate-200'}`}
                        >
                          <Mic className="w-4 h-4" />
                        </button>
                        <button
                          type="button"
                          title={isAiRecording === 'content' ? "Selesai merekam & kirim ke Gemini AI" : "Rekam Suara & Transkripsi dengan Gemini AI"}
                          onClick={isAiRecording === 'content' ? stopMediaRecording : () => startMediaRecording('content')}
                          className={`p-2 rounded-xl border transition-all flex items-center gap-1 text-[10px] font-bold ${isAiRecording === 'content' ? 'bg-indigo-600 text-white border-indigo-500 animate-pulse' : 'bg-indigo-50 text-indigo-600 hover:bg-indigo-100 border-indigo-200'}`}
                        >
                          <Sparkles className="w-3.5 h-3.5" />
                          {isAiRecording === 'content' ? 'Selesai ⏺️' : 'Rekam AI'}
                        </button>
                      </div>
                    </div>
                    <div className="flex items-center justify-between text-xs pt-1">
                      <span className={`font-semibold ${storyContent.trim().length < 30 || storyContent.length > 5000 ? 'text-rose-600 font-bold' : 'text-emerald-600 font-bold'}`}>
                        {storyContent.trim().length === 0 ? (
                          <span className="text-slate-400 font-normal">Minimal 30 karakter, Maksimal 5000 karakter</span>
                        ) : storyContent.trim().length < 30 ? (
                          <span>⚠️ Terlalu sedikit! Kurang {30 - storyContent.trim().length} karakter lagi (Min. 30)</span>
                        ) : storyContent.length > 5000 ? (
                          <span>⚠️ Melebihi batas! Maksimal 5000 karakter</span>
                        ) : (
                          <span>✓ Panjang cerita memenuhi syarat (Siap disimpan/dikirim)</span>
                        )}
                      </span>
                      <span className={`font-bold ${storyContent.length > 5000 || storyContent.trim().length < 30 ? 'text-rose-500' : 'text-slate-400'}`}>
                        {storyContent.length} / 5000 karakter
                      </span>
                    </div>
                  </div>
                </div>
              )}

              {/* STEP 2: 4F Reflections with individual mic buttons */}
              {writeStep === 2 && (
                <div className="space-y-6">
                  
                  {/* FACT */}
                  <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-2 relative">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className="text-base">📝</span>
                        <label className="block text-sm font-bold text-teal-800">Fact (Fakta / Peristiwa):</label>
                      </div>
                      
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => startSpeechRecognition('fact')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isRecording === 'fact' ? 'bg-red-100 text-red-600 border-red-300 animate-pulse' : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-100'}`}
                        >
                          <Mic className="w-3 h-3" />
                          {isRecording === 'fact' ? 'Mendengarkan...' : 'Dikte'}
                        </button>
                        <button
                          type="button"
                          onClick={isAiRecording === 'fact' ? stopMediaRecording : () => startMediaRecording('fact')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isAiRecording === 'fact' ? 'bg-indigo-600 text-white border-indigo-500 animate-pulse' : 'bg-indigo-50 text-indigo-600 border-indigo-100 hover:bg-indigo-100'}`}
                        >
                          <Sparkles className="w-3 h-3" />
                          {isAiRecording === 'fact' ? 'Selesai ⏺️' : 'Rekam AI'}
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center text-[11px] text-slate-500">
                      <span>Tuliskan ringkasan fakta kejadian nyata. Apa yang terjadi, kapan, dan siapa saja yang terlibat?</span>
                      {transcribingAudio && isAiRecording === 'fact' && (
                        <span className="text-teal-600 font-bold animate-pulse">🧠 Memproses suara...</span>
                      )}
                    </div>
                    <textarea
                      rows={2}
                      placeholder="Faktanya saya membagi buah apel saya menjadi dua dan memakannya bersama Dika di meja kelas..."
                      value={storyFact}
                      onChange={(e) => setStoryFact(e.target.value)}
                      className="w-full bg-white border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-medium text-slate-600"
                    />
                  </div>

                  {/* FEELING */}
                  <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-3 relative">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className="text-base">💖</span>
                        <label className="block text-sm font-bold text-rose-800">Feeling (Perasaan Anda):</label>
                      </div>
                      
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => startSpeechRecognition('feeling')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isRecording === 'feeling' ? 'bg-red-100 text-red-600 border-red-300 animate-pulse' : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-100'}`}
                        >
                          <Mic className="w-3 h-3" />
                          {isRecording === 'feeling' ? 'Mendengarkan...' : 'Dikte'}
                        </button>
                        <button
                          type="button"
                          onClick={isAiRecording === 'feeling' ? stopMediaRecording : () => startMediaRecording('feeling')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isAiRecording === 'feeling' ? 'bg-indigo-600 text-white border-indigo-500 animate-pulse' : 'bg-indigo-50 text-indigo-600 border-indigo-100 hover:bg-indigo-100'}`}
                        >
                          <Sparkles className="w-3 h-3" />
                          {isAiRecording === 'feeling' ? 'Selesai ⏺️' : 'Rekam AI'}
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center text-[11px] text-slate-500">
                      <span>Pilih emoji emosional yang paling mewakili perasaanmu saat kejadian itu berlangsung:</span>
                      {transcribingAudio && isAiRecording === 'feeling' && (
                        <span className="text-teal-600 font-bold animate-pulse">🧠 Memproses suara...</span>
                      )}
                    </div>
                    
                    <div className="grid grid-cols-3 sm:grid-cols-6 gap-2">
                      {FEELING_EMOJIS.map(fe => (
                        <button
                          key={fe.emoji}
                          type="button"
                          onClick={() => setStoryFeelingEmoji(fe.emoji)}
                          className={`p-2.5 rounded-xl border flex flex-col items-center justify-center transition-all ${storyFeelingEmoji === fe.emoji ? 'border-rose-500 bg-rose-50/50 scale-105 ring-2 ring-rose-200' : 'bg-white border-slate-100 hover:border-slate-200'}`}
                        >
                          <span className="text-2xl">{fe.emoji}</span>
                          <span className="text-[10px] font-bold text-slate-600 mt-1">{fe.label.split(' ')[0]}</span>
                        </button>
                      ))}
                    </div>

                    <textarea
                      rows={2}
                      placeholder="Jelaskan perasaanmu. Mengapa kamu merasakan hal tersebut?"
                      value={storyFeeling}
                      onChange={(e) => setStoryFeeling(e.target.value)}
                      className="w-full bg-white border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-medium text-slate-600"
                    />
                  </div>

                  {/* FINDING */}
                  <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-2 relative">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className="text-base">💡</span>
                        <label className="block text-sm font-bold text-amber-800">Finding (Temuan / Pembelajaran):</label>
                      </div>
                      
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => startSpeechRecognition('finding')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isRecording === 'finding' ? 'bg-red-100 text-red-600 border-red-300 animate-pulse' : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-100'}`}
                        >
                          <Mic className="w-3 h-3" />
                          {isRecording === 'finding' ? 'Mendengarkan...' : 'Dikte'}
                        </button>
                        <button
                          type="button"
                          onClick={isAiRecording === 'finding' ? stopMediaRecording : () => startMediaRecording('finding')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isAiRecording === 'finding' ? 'bg-indigo-600 text-white border-indigo-500 animate-pulse' : 'bg-indigo-50 text-indigo-600 border-indigo-100 hover:bg-indigo-100'}`}
                        >
                          <Sparkles className="w-3 h-3" />
                          {isAiRecording === 'finding' ? 'Selesai ⏺️' : 'Rekam AI'}
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center text-[11px] text-slate-500">
                      <span>Pelajaran berharga apa yang didapatkan dari pengalaman ini? Nilai kebaikan apa yang kamu temukan?</span>
                      {transcribingAudio && isAiRecording === 'finding' && (
                        <span className="text-teal-600 font-bold animate-pulse">🧠 Memproses suara...</span>
                      )}
                    </div>
                    <textarea
                      rows={2}
                      placeholder="Saya belajar bahwa berbagi tidak mengurangi apa yang kita miliki, melainkan justru melipatgandakan kebahagiaan..."
                      value={storyFinding}
                      onChange={(e) => setStoryFinding(e.target.value)}
                      className="w-full bg-white border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-medium text-slate-600"
                    />
                  </div>

                  {/* FUTURE */}
                  <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-2 relative">
                    <div className="flex items-center justify-between">
                      <div className="flex items-center gap-2">
                        <span className="text-base">🚀</span>
                        <label className="block text-sm font-bold text-blue-800">Future (Masa Depan / Penerapan):</label>
                      </div>
                      
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => startSpeechRecognition('future')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isRecording === 'future' ? 'bg-red-100 text-red-600 border-red-300 animate-pulse' : 'bg-white text-slate-500 border-slate-200 hover:bg-slate-100'}`}
                        >
                          <Mic className="w-3 h-3" />
                          {isRecording === 'future' ? 'Mendengarkan...' : 'Dikte'}
                        </button>
                        <button
                          type="button"
                          onClick={isAiRecording === 'future' ? stopMediaRecording : () => startMediaRecording('future')}
                          className={`p-1.5 rounded-lg border text-[10px] font-bold flex items-center gap-1 transition-all ${isAiRecording === 'future' ? 'bg-indigo-600 text-white border-indigo-500 animate-pulse' : 'bg-indigo-50 text-indigo-600 border-indigo-100 hover:bg-indigo-100'}`}
                        >
                          <Sparkles className="w-3 h-3" />
                          {isAiRecording === 'future' ? 'Selesai ⏺️' : 'Rekam AI'}
                        </button>
                      </div>
                    </div>
                    <div className="flex justify-between items-center text-[11px] text-slate-500">
                      <span>Bagaimana kamu akan menerapkan pembelajaran baik ini di masa depan agar menjadi anak Indonesia yang lebih hebat?</span>
                      {transcribingAudio && isAiRecording === 'future' && (
                        <span className="text-teal-600 font-bold animate-pulse">🧠 Memproses suara...</span>
                      )}
                    </div>
                    <textarea
                      rows={2}
                      placeholder="Ke depan, saya akan selalu berinisiatif membantu atau berbagi dengan teman-teman yang kesulitan..."
                      value={storyFuture}
                      onChange={(e) => setStoryFuture(e.target.value)}
                      className="w-full bg-white border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-medium text-slate-600"
                    />
                  </div>

                </div>
              )}

              {/* STEP 3: 7 Habits */}
              {writeStep === 3 && (
                <div className="space-y-4">
                  <div className="space-y-1">
                    <h3 className="font-bold text-slate-800 text-base">Kaitkan dengan 7 Kebiasaan Anak Indonesia Hebat</h3>
                    <p className="text-xs text-slate-400">Pilih satu atau beberapa kebiasaan baik di bawah ini yang paling berkaitan erat dengan kisahmu!</p>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                    {HABITS_INFO.map(habit => {
                      const isChecked = storyHabits.includes(habit.id);
                      return (
                        <div 
                          key={habit.id}
                          onClick={() => {
                            if (isChecked) {
                              setStoryHabits(prev => prev.filter(h => h !== habit.id));
                            } else {
                              setStoryHabits(prev => [...prev, habit.id]);
                            }
                          }}
                          className={`p-4 rounded-2xl border cursor-pointer flex items-start gap-3.5 transition-all ${isChecked ? 'border-emerald-500 bg-emerald-50/50' : 'border-slate-100 bg-white hover:border-slate-200'}`}
                        >
                          <span className="text-2xl p-1.5 bg-slate-50 rounded-xl">{habit.icon}</span>
                          <div className="space-y-0.5">
                            <h4 className="font-bold text-xs text-slate-800">{habit.label}</h4>
                            <p className="text-[10px] text-slate-400 leading-tight">{habit.desc}</p>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>
              )}

              {/* STEP 4: AI Summary */}
              {writeStep === 4 && (
                <div className="space-y-6">
                  
                  <div className="p-5 bg-teal-50 border border-teal-100 rounded-3xl relative overflow-hidden flex flex-col md:flex-row items-start gap-4">
                    <div className="p-3 bg-teal-600 rounded-2xl text-white">
                      <Sparkles className="w-6 h-6 animate-pulse" />
                    </div>
                    <div className="space-y-2">
                      <h4 className="font-bold text-teal-900 text-sm">Masukan Cerdas dari Si CERDAS AI</h4>
                      <p className="text-xs text-teal-800 leading-relaxed italic">
                        "{aiAnalysisResult?.friendlyFeedback || "Luar biasa! Cerita reflektifmu sangat menyentuh hati dan bernilai positif."}"
                      </p>
                      
                      {aiAnalysisResult?.detectedHabits && aiAnalysisResult.detectedHabits.length > 0 && (
                        <div className="pt-2 text-[11px] text-teal-700 flex flex-wrap gap-1.5">
                          <strong>Habit Terdeteksi AI:</strong>
                          {aiAnalysisResult.detectedHabits.map((h: string) => (
                            <span key={h} className="underline text-teal-800 font-bold">{h}</span>
                          ))}
                        </div>
                      )}
                    </div>
                  </div>

                  {/* Story Illustration Preview */}
                  <div className="space-y-2">
                    <label className="block text-sm font-bold text-slate-700">Visual/Ilustrasi Cerita:</label>
                    <p className="text-xs text-slate-400">Si CERDAS AI merekomendasikan prompt lukisan berikut:</p>
                    <div className="bg-slate-50 p-3 rounded-xl text-xs text-slate-600 border border-slate-100 italic">
                      {aiAnalysisResult?.illustrationPrompt || "A cute watercolor painting of students sharing a fruit basket, smiling kids, sun rays, digital art"}
                    </div>

                    <div className="flex flex-col items-center justify-center p-6 border-2 border-dashed border-slate-200 bg-slate-50/50 rounded-2xl text-center space-y-4">
                      {generatedImageUrl === 'FALLBACK' ? (
                        <div className="space-y-2">
                          <div className="text-4xl animate-bounce">🎨💻</div>
                          <p className="text-xs text-slate-500 font-medium">Menggunakan template ilustrasi premium berbasis kebiasaan pilihanmu.</p>
                          <p className="text-[10px] text-slate-400">Gambar penutup akan disematkan otomatis saat cerita disimpan.</p>
                        </div>
                      ) : generatedImageUrl ? (
                        <div className="relative max-w-xs mx-auto aspect-square rounded-2xl overflow-hidden shadow-sm">
                          <img src={generatedImageUrl} alt="Generated Art" className="w-full h-full object-cover" />
                          <button 
                            onClick={() => setGeneratedImageUrl(null)}
                            className="absolute top-2 right-2 bg-black/60 hover:bg-black/80 text-white p-1 rounded-full text-xs"
                          >
                            Ulangi
                          </button>
                        </div>
                      ) : (
                        <button
                          type="button"
                          disabled={generatingIllustration}
                          onClick={handleGenerateIllustration}
                          className="flex items-center gap-2 bg-gradient-to-r from-teal-600 to-amber-500 hover:from-teal-700 hover:to-amber-600 text-white font-bold text-xs px-5 py-3 rounded-xl shadow-xs transition-transform active:scale-95 duration-100 disabled:opacity-50"
                        >
                          {generatingIllustration ? (
                            <>
                              <RefreshCw className="w-4 h-4 animate-spin" />
                              Menggambar dengan AI...
                            </>
                          ) : (
                            <>
                              <Sparkles className="w-4 h-4" />
                              Buat Ilustrasi Sampul dengan AI!
                            </>
                          )}
                        </button>
                      )}
                    </div>
                  </div>

                </div>
              )}

              {/* Action Buttons */}
              <div className="flex items-center justify-between pt-4 border-t border-slate-50">
                <button
                  type="button"
                  disabled={writeStep === 1}
                  onClick={() => setWriteStep(prev => prev - 1)}
                  className="flex items-center gap-1.5 text-slate-600 hover:text-slate-800 font-bold text-xs"
                >
                  <ChevronLeft className="w-4 h-4" /> Kembali
                </button>

                {writeStep < 4 ? (
                  <button
                    type="button"
                    onClick={() => {
                      if (writeStep === 1) {
                        const cleanTitle = storyTitle.trim();
                        const cleanContent = storyContent.trim();
                        if (cleanTitle.length < 3) {
                          alert("⚠️ Judul cerita terlalu singkat! Harap isi judul minimal 3 karakter.");
                          return;
                        }
                        if (cleanContent.length < 30) {
                          alert(`⚠️ Cerita terlalu singkat (${cleanContent.length} karakter)! Sistem tidak menerima cerita kurang dari 30 karakter. Mohon tulis cerita lebih lengkap (minimal 30, maksimal 5000 karakter).`);
                          return;
                        }
                        if (cleanContent.length > 5000) {
                          alert(`⚠️ Cerita terlalu panjang (${cleanContent.length} karakter)! Maksimal panjang cerita adalah 5000 karakter.`);
                          return;
                        }
                      }
                      if (writeStep === 2 && !storyFact.trim()) {
                        alert("Harap isi Fact (Fakta) kejadian nyata Anda.");
                        return;
                      }
                      if (writeStep === 3) {
                        handleAnalyzeStoryWithAI();
                        return;
                      }
                      setWriteStep(prev => prev + 1);
                    }}
                    className="flex items-center gap-1 bg-teal-600 hover:bg-teal-700 text-white font-bold text-xs px-5 py-2.5 rounded-xl shadow-xs active:scale-95 duration-100"
                  >
                    {analyzingStory ? (
                      <>
                        <RefreshCw className="w-4 h-4 animate-spin" />
                        Sedang Menganalisis...
                      </>
                    ) : (
                      <>
                        Lanjut <ChevronRight className="w-4 h-4" />
                      </>
                    )}
                  </button>
                ) : (
                  <div className="flex items-center gap-3">
                    <button
                      type="button"
                      disabled={submittingStory}
                      onClick={() => handleSaveStory('draft')}
                      className="text-slate-500 hover:text-slate-800 font-bold text-xs px-4 py-2"
                    >
                      Simpan Draf
                    </button>
                    <button
                      type="button"
                      disabled={submittingStory}
                      onClick={() => handleSaveStory('submitted')}
                      className="bg-teal-600 hover:bg-teal-700 text-white font-bold text-xs px-6 py-2.5 rounded-xl shadow-xs active:scale-95 duration-100"
                    >
                      {submittingStory ? 'Mengirim...' : 'Kirim Cerita Hebat! 🚀'}
                    </button>
                  </div>
                )}
              </div>

            </div>

          </div>
        )}

        {/* ======================================= */}
        {/* TAB 3: STUDENT HABITS JURNAL / TRACKER */}
        {/* ======================================= */}
        {activeTab === 'habits' && userProfile?.role === 'murid' && (
          <div className="max-w-4xl mx-auto space-y-6">
            
            <div className="space-y-1 text-center md:text-left">
              <h1 className="text-3xl font-extrabold text-teal-800">Jurnal 7 Kebiasaan Harian 📅</h1>
              <p className="text-slate-500 text-sm">Konsistensi membentuk karakter anak hebat. Catat kebiasaan muliamu setiap hari!</p>
            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-6 flex flex-col justify-between">
                <div className="space-y-4">
                  <div className="space-y-2">
                    <label className="block text-xs font-bold text-slate-400 uppercase tracking-wider">Pilih Tanggal:</label>
                    <input
                      type="date"
                      max={new Date().toISOString().split('T')[0]}
                      value={selectedDate}
                      onChange={(e) => setSelectedDate(e.target.value)}
                      className="w-full bg-slate-50 border border-slate-200 rounded-xl px-4 py-2.5 focus:outline-none focus:ring-2 focus:ring-teal-400 text-sm font-bold text-slate-700"
                    />
                  </div>

                  <div className="p-4 bg-teal-50 rounded-2xl border border-teal-100/50 space-y-3">
                    <div className="flex items-center gap-2">
                      <Flame className="w-5 h-5 text-amber-500 fill-amber-500 animate-pulse" />
                      <h4 className="font-bold text-teal-900 text-xs uppercase tracking-wider">Statistik Literasi & Karakter</h4>
                    </div>
                    
                    <div className="grid grid-cols-2 gap-2 pt-2 text-center">
                      <div className="bg-white p-2.5 rounded-xl shadow-xs">
                        <span className="text-[10px] text-slate-400 block font-semibold">Tingkat Kebiasaan</span>
                        <span className="text-xl font-extrabold text-teal-700">{getStudentHabitStats(currentUser.uid).score}%</span>
                      </div>
                      <div className="bg-white p-2.5 rounded-xl shadow-xs">
                        <span className="text-[10px] text-slate-400 block font-semibold">Hari Tercatat</span>
                        <span className="text-xl font-extrabold text-teal-700">{getStudentHabitStats(currentUser.uid).count} Hari</span>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="text-xs text-slate-400 text-center italic border-t border-slate-100 pt-4">
                  "Kita adalah apa yang kita lakukan berulang kali. Kebiasaan mulia, masa depan bahagia!"
                </div>
              </div>

              <div className="lg:col-span-2 bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-6">
                <div className="flex items-center justify-between">
                  <h3 className="font-bold text-slate-800 text-base">Amalkan Kebiasaan Hari Ini:</h3>
                  <span className="text-xs text-slate-500">
                    {new Date(selectedDate).toLocaleDateString('id-ID', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' })}
                  </span>
                </div>

                <div className="space-y-3">
                  {HABITS_INFO.map(habit => {
                    const isChecked = !!(todayHabitLog as any)[habit.id];
                    return (
                      <div
                        key={habit.id}
                        onClick={() => {
                          setTodayHabitLog(prev => ({
                            ...prev,
                            [habit.id]: !isChecked
                          }));
                        }}
                        className={`p-3.5 rounded-2xl border cursor-pointer flex items-center justify-between transition-all ${isChecked ? 'border-emerald-500 bg-emerald-50/40' : 'border-slate-100 bg-slate-50/50 hover:border-slate-200'}`}
                      >
                        <div className="flex items-center gap-3">
                          <span className="text-2xl">{habit.icon}</span>
                          <div className="space-y-0.5">
                            <span className="text-xs font-bold text-slate-800 block">{habit.label}</span>
                            <span className="text-[10px] text-slate-400">{habit.desc}</span>
                          </div>
                        </div>

                        <div className={`w-6 h-6 rounded-full flex items-center justify-center transition-all ${isChecked ? 'bg-emerald-500 text-white scale-105' : 'bg-white border border-slate-200'}`}>
                          {isChecked && <Check className="w-3.5 h-3.5 stroke-[3px]" />}
                        </div>
                      </div>
                    );
                  })}
                </div>

                <button
                  type="button"
                  disabled={savingHabits}
                  onClick={handleSaveHabits}
                  className="w-full bg-teal-600 hover:bg-teal-700 text-white font-bold py-3 px-6 rounded-2xl shadow-xs transition-all active:scale-98 disabled:opacity-50 text-xs"
                >
                  {savingHabits ? 'Sedang Menyimpan...' : 'Simpan Log Kebiasaan Hebat 🏆'}
                </button>
              </div>

            </div>

          </div>
        )}

        {/* ======================================= */}
        {/* TAB 4: HOMEROOM TEACHER DASHBOARD */}
        {/* ======================================= */}
        {activeTab === 'classroom' && userProfile?.role === 'guru_wali' && (() => {
          const teacherClassesList = userProfile.kelas ? userProfile.kelas.split(',').map(c => c.trim().toUpperCase()).filter(Boolean) : [];
          const activeClassFilter = selectedTeacherClassFilter === 'Semua' ? userProfile.kelas : selectedTeacherClassFilter;

          return (
          <div className="space-y-6">
            
            <div className="flex flex-col md:flex-row md:items-center justify-between gap-4">
              <div className="space-y-1">
                <h1 className="text-3xl font-extrabold text-teal-800">Ruang Kelas {userProfile.kelas} 👩‍🏫</h1>
                <p className="text-slate-500 text-sm">Kelola progres literasi siswa, pantau kebiasaan baik harian, dan berikan ulasan cerita mereka.</p>
              </div>

              {teacherClassesList.length > 1 && (
                <div className="flex items-center gap-2 bg-white px-4 py-2.5 rounded-2xl border border-slate-200 shadow-2xs">
                  <span className="text-xs font-bold text-slate-500">Filter Kelas Binaan:</span>
                  <select
                    value={selectedTeacherClassFilter}
                    onChange={(e) => setSelectedTeacherClassFilter(e.target.value)}
                    className="bg-slate-50 border border-slate-200 text-teal-800 font-extrabold text-xs px-3 py-1.5 rounded-xl focus:outline-none cursor-pointer"
                  >
                    <option value="Semua">✨ Semua Kelas Binaan ({userProfile.kelas})</option>
                    {teacherClassesList.map(c => (
                      <option key={c} value={c}>Kelas {c}</option>
                    ))}
                  </select>
                </div>
              )}
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              
              <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-xs flex items-center gap-4">
                <div className="p-3 bg-teal-50 text-teal-600 rounded-2xl">
                  <BookMarked className="w-6 h-6" />
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 block font-bold uppercase tracking-wider">Total Cerita Siswa</span>
                  <span className="text-2xl font-extrabold text-slate-800">
                    {stories.filter(s => isTeacherClassMatch(activeClassFilter, s.authorKelas)).length} Cerita
                  </span>
                </div>
              </div>

              <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-xs flex items-center gap-4">
                <div className="p-3 bg-amber-50 text-amber-500 rounded-2xl">
                  <Clock className="w-6 h-6 animate-pulse" />
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 block font-bold uppercase tracking-wider">Menunggu Ulasan</span>
                  <span className="text-2xl font-extrabold text-slate-800">
                    {stories.filter(s => isTeacherClassMatch(activeClassFilter, s.authorKelas) && s.status === 'submitted').length} Kisah
                  </span>
                </div>
              </div>

              <div className="bg-white p-5 rounded-3xl border border-slate-100 shadow-xs flex items-center gap-4">
                <div className="p-3 bg-emerald-50 text-emerald-600 rounded-2xl">
                  <Activity className="w-6 h-6" />
                </div>
                <div>
                  <span className="text-[10px] text-slate-400 block font-bold uppercase tracking-wider">Indeks Kebiasaan Kelas</span>
                  <span className="text-2xl font-extrabold text-slate-800">
                    {Math.round(
                      habitLogs.filter(l => isTeacherClassMatch(activeClassFilter, l.studentKelas)).length > 0
                        ? habitLogs
                            .filter(l => isTeacherClassMatch(activeClassFilter, l.studentKelas))
                            .reduce((acc, curr) => {
                              let ticks = 0;
                              if (curr.bangun_pagi) ticks++;
                              if (curr.beribadah) ticks++;
                              if (curr.berolahraga) ticks++;
                              if (curr.makan_sehat) ticks++;
                              if (curr.gemar_belajar) ticks++;
                              if (curr.bermasyarakat) ticks++;
                              if (curr.tidur_cepat) ticks++;
                              return acc + (ticks / 7);
                            }, 0) / habitLogs.filter(l => isTeacherClassMatch(activeClassFilter, l.studentKelas)).length * 100
                        : 0
                    )}%
                  </span>
                </div>
              </div>

            </div>

            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              
              <div className="lg:col-span-2 bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-4">
                <h3 className="font-bold text-slate-800 text-base">Kotak Masuk Cerita Kelas</h3>
                <p className="text-xs text-slate-400">Bacalah refleksi murid dan berikan bimbingan serta apresiasi bintang yang memotivasi.</p>
                
                {stories.filter(s => isTeacherClassMatch(activeClassFilter, s.authorKelas)).length === 0 ? (
                  <div className="p-12 text-center text-slate-400 italic text-sm">
                    Belum ada siswa di kelas Anda ({activeClassFilter}) yang memublikasikan cerita.
                  </div>
                ) : (
                  <div className="space-y-3 pt-2">
                    {stories
                      .filter(s => isTeacherClassMatch(activeClassFilter, s.authorKelas))
                      .map(story => (
                        <div
                          key={story.id}
                          onClick={() => setSelectedStory(story)}
                          className="p-4 rounded-2xl border border-slate-100 hover:border-teal-200 hover:bg-slate-50/50 cursor-pointer flex items-center justify-between transition-all duration-150"
                        >
                          <div className="space-y-1 max-w-lg">
                            <div className="flex items-center gap-2">
                              <span className="text-xs font-bold text-slate-700">{story.authorName}</span>
                              <span className="text-xs text-slate-400">·</span>
                              <span className="text-[10px] font-bold bg-slate-100 text-slate-600 px-1.5 py-0.5 rounded-md">Kelas {story.authorKelas}</span>
                              <span className="text-xs text-slate-400">·</span>
                              <span className="text-[10px] text-slate-400">
                                {story.createdAt?.seconds ? new Date(story.createdAt.seconds * 1000).toLocaleDateString('id-ID') : 'Hari ini'}
                              </span>
                            </div>
                            <h4 className="font-bold text-sm text-slate-800 truncate">{story.title}</h4>
                            <div className="text-[10px] text-slate-400 truncate flex items-center gap-1">
                              <strong>Refleksi:</strong>
                              <span className="font-medium text-teal-600 truncate">{story.fact.substring(0, 50)}...</span>
                            </div>
                          </div>

                          <div className="flex items-center gap-2">
                            {story.status === 'submitted' ? (
                              <span className="bg-amber-100 text-amber-800 font-extrabold text-[9px] px-2 py-0.5 rounded-md uppercase tracking-wider animate-pulse">Perlu Ulasan</span>
                            ) : (
                              <span className="bg-slate-100 text-slate-500 font-extrabold text-[9px] px-2 py-0.5 rounded-md uppercase tracking-wider">Selesai Ulas</span>
                            )}
                            <button
                              type="button"
                              onClick={(e) => {
                                e.stopPropagation();
                                handleDeleteStory(story.id);
                              }}
                              title="Hapus Cerita"
                              className="p-1 hover:bg-rose-50 text-slate-400 hover:text-rose-600 rounded-lg transition-colors"
                            >
                              🗑️
                            </button>
                            <ChevronRight className="w-4 h-4 text-slate-400" />
                          </div>
                        </div>
                      ))}
                  </div>
                )}
              </div>

              <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-4">
                <h3 className="font-bold text-slate-800 text-base">Rapor Kebiasaan Siswa</h3>
                <p className="text-xs text-slate-400">Peringkat kepatuhan pilar kebiasaan baik siswa.</p>

                <div className="space-y-3 pt-2">
                  {Object.values(allUsers)
                    .filter(u => u.role === 'murid' && isTeacherClassMatch(activeClassFilter, u.kelas))
                    .map(student => {
                      const stats = getStudentHabitStats(student.uid);
                      return (
                        <div key={student.uid} className="flex items-center justify-between p-3 bg-slate-50 rounded-2xl border border-slate-100">
                          <div className="flex items-center gap-2.5">
                            <img src={student.photoURL} alt={student.displayName} className="w-8 h-8 rounded-full" />
                            <div className="truncate max-w-[120px]">
                              <span className="text-xs font-bold text-slate-800 block truncate">{student.displayName}</span>
                              <span className="text-[10px] text-slate-400 block">Kelas {student.kelas} · {stats.count} hari</span>
                            </div>
                          </div>

                          <div className="text-right">
                            <span className="text-xs font-extrabold text-teal-700 block">{stats.score}%</span>
                            <div className="w-16 h-1 bg-slate-200 rounded-full overflow-hidden mt-0.5">
                              <div style={{ width: `${stats.score}%` }} className="h-full bg-teal-500" />
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  
                  {Object.values(allUsers).filter(u => u.role === 'murid' && isTeacherClassMatch(activeClassFilter, u.kelas)).length === 0 && (
                    <p className="text-xs text-slate-400 italic text-center py-4">Belum ada murid terdaftar di kelas Anda ({activeClassFilter}).</p>
                  )}
                </div>
              </div>

            </div>

          </div>
          );
        })()}

        {/* ======================================= */}
        {/* TAB 5: GUIDANCE COUNSELOR (BK) CORNER */}
        {/* ======================================= */}
        {activeTab === 'bk_corner' && userProfile?.role === 'guru_bk' && (
          <div className="space-y-6">
            
            <div className="space-y-1">
              <h1 className="text-3xl font-extrabold text-teal-800">Pojok Bimbingan & Konseling (BK) 🛋️❤️</h1>
              <p className="text-slate-500 text-sm">Menjaga kehangatan batin dan kesejahteraan emosional seluruh anak didik dengan deteksi emosi.</p>
            </div>

            <div className="bg-gradient-to-r from-rose-500 to-rose-600 rounded-3xl p-6 text-white shadow-md flex flex-col md:flex-row items-center justify-between gap-6">
              <div className="space-y-2 max-w-xl">
                <h3 className="font-bold text-lg">Indikator Kesejahteraan Emosi Siswa</h3>
                <p className="text-xs text-rose-100 leading-relaxed">
                  Sistem filter emosi menyortir cerita-cerita siswa yang berstatus 'Sedih' atau 'Cemas/Takut' dari refleksi 4F. Berikan pelukan semangat lewat catatan konseling Anda secara tepat waktu.
                </p>
              </div>

              <div className="flex gap-4">
                <div className="bg-white/10 px-5 py-4 rounded-2xl text-center">
                  <span className="text-2xl block">😢</span>
                  <span className="text-xs font-bold block mt-1">
                    {stories.filter(s => s.feelingEmoji === '😢' || s.feelingEmoji === '😟' || s.feelingEmoji === '😠').length} Siswa Butuh Dukungan
                  </span>
                </div>
              </div>
            </div>

            {/* List Flagged Stories */}
            <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-4">
              <h3 className="font-bold text-slate-800 text-base">Kotak Refleksi Butuh Sentuhan BK</h3>
              <p className="text-xs text-slate-400">Berikut adalah cerita dari siswa di seluruh kelas yang mencatatkan emosi penuh kecemasan atau kesedihan:</p>

              <div className="space-y-3 pt-2">
                {stories
                  .filter(s => s.feelingEmoji === '😢' || s.feelingEmoji === '😟' || s.feelingEmoji === '😠')
                  .map(story => (
                    <div
                      key={story.id}
                      onClick={() => setSelectedStory(story)}
                      className="p-4 rounded-2xl border border-rose-100 hover:bg-rose-50/20 cursor-pointer flex items-center justify-between transition-colors"
                    >
                      <div className="space-y-1">
                        <div className="flex items-center gap-2">
                          <span className="text-xs font-bold text-slate-800">{story.authorName}</span>
                          <span className="text-xs text-slate-400">·</span>
                          <span className="text-[10px] bg-rose-50 text-rose-700 px-2 py-0.5 rounded-full font-bold">
                            Perasaan: {story.feelingEmoji} ({story.feeling.substring(0, 30)}...)
                          </span>
                        </div>
                        <h4 className="font-bold text-sm text-slate-800 mt-1">{story.title}</h4>
                        <p className="text-xs text-slate-500 line-clamp-1">{story.content}</p>
                      </div>

                      <div className="flex items-center gap-2 text-rose-600 font-bold text-xs">
                        Beri Sentuhan <ChevronRight className="w-4 h-4 text-rose-400" />
                      </div>
                    </div>
                  ))}

                {stories.filter(s => s.feelingEmoji === '😢' || s.feelingEmoji === '😟' || s.feelingEmoji === '😠').length === 0 && (
                  <p className="text-slate-400 text-sm italic text-center py-6">Alhamdulillah, seluruh kondisi emosi siswa hari ini dalam keadaan stabil dan bahagia. 😊💖</p>
                )}
              </div>
            </div>

          </div>
        )}

        {/* ======================================= */}
        {/* TAB 6: ADMINISTRATOR CONTROL PANEL */}
        {/* ======================================= */}
        {activeTab === 'admin' && (userProfile?.role === 'admin' || currentUser?.email === 'isumayasa91@guru.smp.belajar.id' || offlineMode) && (
          <div className="space-y-6 animate-fadeIn pb-12">
            
            <div className="flex flex-col md:flex-row justify-between items-start md:items-center gap-4">
              <div className="space-y-1">
                <h1 className="text-3xl font-extrabold text-indigo-900 flex items-center gap-2">Panel Administrasi CERDAS 🛡️</h1>
                <p className="text-slate-500 text-sm">Kelola data murid, data guru wali, serta lakukan sinkronisasi cerita kelas secara otomatis.</p>
              </div>
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-indigo-700 bg-indigo-50 px-3 py-1.5 rounded-xl border border-indigo-100 flex items-center gap-1.5">
                  <Activity className="w-3.5 h-3.5 text-indigo-500 animate-pulse" />
                  Status Sistem: Sinkron
                </span>
              </div>
            </div>

            {/* Admin Sub-navigation Tabs */}
            <div className="flex border-b border-slate-200 gap-1 overflow-x-auto pb-px">
              {[
                { id: 'register', label: '📝 Registrasi & Impor Baru', icon: '📝' },
                { id: 'guru_wali', label: '👩‍🏫 Data Guru Wali', icon: '👩‍🏫' },
                { id: 'murid', label: '🎒 Data Murid / Siswa', icon: '🎒' },
                { id: 'stories', label: '📚 Kelola Cerita Kelas', icon: '📚' }
              ].map(subTab => (
                <button
                  key={subTab.id}
                  onClick={() => setAdminSubTab(subTab.id as any)}
                  className={`px-4 py-3 font-bold text-xs flex items-center gap-2 border-b-2 transition-all whitespace-nowrap ${adminSubTab === subTab.id ? 'border-indigo-600 text-indigo-700 font-extrabold bg-indigo-50/30' : 'border-transparent text-slate-500 hover:text-slate-700'}`}
                >
                  <span>{subTab.icon}</span>
                  <span>{subTab.label}</span>
                </button>
              ))}
            </div>

            {/* Sub-tab 1: Registration Form (Manual / Bulk) */}
            {adminSubTab === 'register' && (
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* Form Tambah Pengguna */}
                <div className="lg:col-span-2 bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-6">
                  <div className="flex justify-between items-center border-b border-slate-100 pb-3">
                    <h3 className="font-bold text-slate-800 text-base flex items-center gap-2">
                      <span>👤 Mendaftarkan Pengguna Baru</span>
                    </h3>
                    
                    {/* Segmented input mode selector */}
                    <div className="flex gap-1 p-1 bg-slate-100 rounded-xl text-[10px]">
                      <button
                        type="button"
                        onClick={() => setAdminInputMode('excel')}
                        className={`px-3 py-1.5 rounded-lg font-bold text-center transition-all ${adminInputMode === 'excel' ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-600 hover:text-slate-800'}`}
                      >
                        📊 Upload File Excel
                      </button>
                      <button
                        type="button"
                        onClick={() => setAdminInputMode('manual')}
                        className={`px-3 py-1.5 rounded-lg font-bold text-center transition-all ${adminInputMode === 'manual' ? 'bg-white text-slate-800 shadow-xs' : 'text-slate-500 hover:text-slate-700'}`}
                      >
                        Input Manual
                      </button>
                      <button
                        type="button"
                        onClick={() => setAdminInputMode('bulk')}
                        className={`px-3 py-1.5 rounded-lg font-bold text-center transition-all ${adminInputMode === 'bulk' ? 'bg-white text-slate-800 shadow-xs' : 'text-slate-500 hover:text-slate-700'}`}
                      >
                        Copy-Paste Teks
                      </button>
                    </div>
                  </div>

                  {adminInputMode === 'excel' ? (
                    <div className="space-y-5 animate-fadeIn">
                      <div className="bg-gradient-to-r from-emerald-50 to-teal-50 p-4 rounded-2xl border border-emerald-100/80 leading-relaxed text-xs text-emerald-950 space-y-2">
                        <div className="flex justify-between items-center">
                          <strong className="text-emerald-900 font-extrabold flex items-center gap-1.5 text-sm">
                            <span>📊 Impor Pengguna dari File Excel / CSV</span>
                          </strong>
                          <button
                            type="button"
                            onClick={handleDownloadExcelTemplate}
                            className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-xl text-[11px] transition-colors shadow-2xs flex items-center gap-1"
                          >
                            <span>📥 Download Format Templat (.xlsx)</span>
                          </button>
                        </div>
                        <p className="text-slate-600">
                          Unggah file Excel (<strong>.xlsx</strong>, <strong>.xls</strong>) atau <strong>.csv</strong> berisi daftar nama siswa/guru. Kolom yang didukung: <em>Nama Lengkap, Email, Peran (murid/guru_wali/guru_bk/admin), Kelas</em>.
                        </p>
                      </div>

                      {/* File Upload Box */}
                      <div className="border-2 border-dashed border-indigo-200 hover:border-indigo-400 bg-indigo-50/30 transition-all rounded-2xl p-6 text-center space-y-3 cursor-pointer relative group">
                        <input
                          type="file"
                          accept=".xlsx, .xls, .csv"
                          onChange={handleExcelFileSelect}
                          className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
                        />
                        <div className="p-3 bg-indigo-100/80 text-indigo-700 rounded-full w-12 h-12 flex items-center justify-center mx-auto group-hover:scale-110 transition-transform">
                          <span className="text-xl">📁</span>
                        </div>
                        <div>
                          <p className="text-xs font-bold text-slate-700">
                            {excelFileName ? `File Terpilih: ${excelFileName}` : 'Klik atau Tarik File Excel (.xlsx / .csv) Ke Sini'}
                          </p>
                          <p className="text-[11px] text-slate-400 mt-0.5">
                            Sistem akan membaca dan memeriksa baris data secara otomatis
                          </p>
                        </div>
                      </div>

                      {/* Parsed Preview Table */}
                      {excelParsedUsers.length > 0 && (
                        <div className="space-y-3 pt-2">
                          <div className="flex justify-between items-center">
                            <span className="text-xs font-bold text-slate-700">
                              📋 Hasil Pembacaan Data ({excelParsedUsers.length} Pengguna Terdeteksi)
                            </span>
                            <span className="text-[10px] font-semibold text-emerald-700 bg-emerald-50 px-2.5 py-0.5 rounded-full border border-emerald-100">
                              ✓ Data Siap Diimpor
                            </span>
                          </div>

                          <div className="max-h-60 overflow-y-auto border border-slate-200 rounded-2xl text-xs bg-slate-50/50">
                            <table className="w-full text-left border-collapse">
                              <thead className="bg-slate-100 sticky top-0 text-slate-600 font-bold text-[11px] border-b border-slate-200">
                                <tr>
                                  <th className="p-2.5">No</th>
                                  <th className="p-2.5">Nama Lengkap</th>
                                  <th className="p-2.5">Email Akun</th>
                                  <th className="p-2.5">Peran</th>
                                  <th className="p-2.5">Kelas</th>
                                </tr>
                              </thead>
                              <tbody className="divide-y divide-slate-100 text-[11px]">
                                {excelParsedUsers.map((u, i) => (
                                  <tr key={u.uid} className="hover:bg-white transition-colors">
                                    <td className="p-2.5 text-slate-400 font-mono">{i + 1}</td>
                                    <td className="p-2.5 font-bold text-slate-800">{u.displayName}</td>
                                    <td className="p-2.5 text-slate-600 font-mono text-[10px]">{u.email}</td>
                                    <td className="p-2.5">
                                      <span className={`px-2 py-0.5 rounded-full font-bold text-[10px] ${u.role === 'murid' ? 'bg-emerald-100 text-emerald-800' : u.role === 'guru_wali' ? 'bg-amber-100 text-amber-800' : 'bg-indigo-100 text-indigo-800'}`}>
                                        {u.role === 'murid' ? 'Murid' : u.role === 'guru_wali' ? 'Guru Wali' : u.role === 'guru_bk' ? 'Guru BK' : 'Admin'}
                                      </span>
                                    </td>
                                    <td className="p-2.5 font-bold text-slate-700">{u.kelas || 'SEMUA'}</td>
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          </div>

                          <button
                            type="button"
                            disabled={excelImporting}
                            onClick={handleImportExcelUsers}
                            className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 rounded-2xl text-xs shadow-md transition-all disabled:opacity-50 flex items-center justify-center gap-2"
                          >
                            <span>🚀 Impor {excelParsedUsers.length} Pengguna dari Excel Ke Database Sekarang</span>
                          </button>
                        </div>
                      )}
                    </div>
                  ) : adminInputMode === 'bulk' ? (
                    <form onSubmit={handleAdminBulkUpload} className="space-y-4">
                      <div className="bg-amber-50/70 p-4 rounded-2xl text-xs text-amber-900 border border-amber-100 leading-relaxed space-y-1">
                        <strong>💡 Fitur Impor Masal:</strong>
                        <p>Salin data guru wali atau murid dari Excel atau Google Sheets Anda dan tempel di bawah. Sistem akan secara otomatis mendaftarkan semua akun ini.</p>
                      </div>

                      {/* Dropdown Default Target Role */}
                      <div className="space-y-2">
                        <label className="block text-xs font-bold text-slate-600">Pilih Peran Default Impor:</label>
                        <div className="flex gap-2">
                          {[
                            { id: 'murid', label: '🎒 Siswa / Murid', color: 'border-emerald-200' },
                            { id: 'guru_wali', label: '👩‍🏫 Guru Wali', color: 'border-amber-200' }
                          ].map(opt => (
                            <button
                              key={opt.id}
                              type="button"
                              onClick={() => setBulkImportRole(opt.id as any)}
                              className={`flex-1 py-2.5 px-4 border rounded-2xl text-xs font-bold transition-all flex items-center justify-center gap-1.5 ${bulkImportRole === opt.id ? 'bg-indigo-50 border-indigo-500 text-indigo-800 scale-102 ring-2 ring-indigo-100' : 'bg-white border-slate-200 hover:bg-slate-50 text-slate-600'}`}
                            >
                              {opt.label}
                            </button>
                          ))}
                        </div>
                      </div>
                      
                      <div className="bg-slate-50 p-3 rounded-xl border border-slate-100 text-[10px] text-slate-500 font-mono space-y-1">
                        <span className="font-bold text-slate-600 uppercase block">Format Kolom (Pisahkan Koma atau Tab):</span>
                        <code>Nama Lengkap, Email Sekolah, Kelas</code>
                        <span className="block mt-1 text-slate-400">Contoh baris:</span>
                        <code>Ahmad Fauzi, ahmad@siswa.belajar.id, VII-A</code>
                        <code>Budi Hartono S.Pd, budi@guru.belajar.id, VII-A</code>
                      </div>

                      <div className="space-y-1">
                        <label className="block text-xs font-bold text-slate-600">Tempelkan Baris Data Excel Di Sini:</label>
                        <textarea
                          rows={6}
                          required
                          placeholder="Andi Wijaya, andi@siswa.belajar.id, VII-A&#10;Siti Aminah S.Pd, siti@guru.belajar.id, VII-A"
                          value={bulkInputText}
                          onChange={(e) => setBulkInputText(e.target.value)}
                          className="w-full border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-mono leading-relaxed"
                        />
                      </div>

                      <div className="flex gap-2">
                        <button
                          type="button"
                          onClick={() => {
                            const template = bulkImportRole === 'murid' 
                              ? "Andi Wijaya, andi@siswa.belajar.id, VII-A\nSahrul Gunawan, sahrul@siswa.belajar.id, VII-A\nRina Roslina, rina@siswa.belajar.id, VII-B"
                              : "Budi Hartono S.Pd, budi@guru.belajar.id, VII-A\nSiti Aminah M.Pd, siti@guru.belajar.id, VII-B";
                            setBulkInputText(template);
                          }}
                          className="flex-1 bg-slate-50 hover:bg-slate-100 text-slate-600 border border-slate-200 font-bold py-2.5 rounded-xl text-xs transition-colors"
                        >
                          Contoh Template
                        </button>
                        
                        <button
                          type="submit"
                          disabled={isBulkUploading}
                          className="flex-1 bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-2.5 rounded-xl text-xs shadow-xs transition-colors disabled:opacity-50"
                        >
                          {isBulkUploading ? 'Mengimpor...' : 'Unggah Masal Sekarang 🚀'}
                        </button>
                      </div>
                    </form>
                  ) : (
                    <form onSubmit={handleAdminCreateUser} className="space-y-4">
                      <div className="space-y-1">
                        <label className="block text-xs font-bold text-slate-600">Nama Lengkap:</label>
                        <input
                          type="text"
                          required
                          placeholder="Contoh: Budi Hartono, S.Pd / Ahmad Fauzi"
                          value={adminNewUserName}
                          onChange={(e) => setAdminNewUserName(e.target.value)}
                          className="w-full border border-slate-200 px-3.5 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-medium text-slate-700"
                        />
                      </div>

                      <div className="space-y-1">
                        <label className="block text-xs font-bold text-slate-600">Email Akun Belajar/Pribadi:</label>
                        <input
                          type="email"
                          required
                          placeholder="Contoh: andi@siswa.belajar.id"
                          value={adminNewUserEmail}
                          onChange={(e) => setAdminNewUserEmail(e.target.value)}
                          className="w-full border border-slate-200 px-3.5 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-medium text-slate-700"
                        />
                      </div>

                      <div className="space-y-1">
                        <label className="block text-xs font-bold text-slate-600">Peran Pengguna:</label>
                        <select
                          value={adminNewUserRole}
                          onChange={(e: any) => setAdminNewUserRole(e.target.value)}
                          className="w-full border border-slate-200 px-3.5 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs bg-white text-slate-700 font-semibold"
                        >
                          <option value="murid">🎒 Murid / Siswa</option>
                          <option value="guru_wali">👩‍🏫 Guru Wali (Homeroom)</option>
                          <option value="guru_bk">❤️ Guru BK (Konselor)</option>
                        </select>
                      </div>

                      {(adminNewUserRole === 'murid' || adminNewUserRole === 'guru_wali') && (
                        <div className="space-y-1 animate-fadeIn">
                          <label className="block text-xs font-bold text-slate-600">
                            {adminNewUserRole === 'guru_wali' ? 'Penugasan Kelas Binaan (Boleh >1 Kelas):' : 'Penugasan Kelas Siswa:'}
                          </label>
                          <input
                            type="text"
                            required
                            placeholder={adminNewUserRole === 'guru_wali' ? "Contoh: 7B, 7C (Pisahkan koma jika >1 kelas)" : "Contoh: VII-A atau 7B"}
                            value={adminNewUserKelas}
                            onChange={(e) => setAdminNewUserKelas(e.target.value)}
                            className="w-full border border-slate-200 px-3.5 py-2.5 rounded-xl focus:outline-none focus:ring-2 focus:ring-teal-400 text-xs font-medium text-slate-700"
                          />
                          {adminNewUserRole === 'guru_wali' && (
                            <p className="text-[10px] text-indigo-600 font-medium">💡 Untuk guru binaan lebih dari 1 kelas, pisahkan dengan koma (contoh: 7B, 7C).</p>
                          )}
                        </div>
                      )}

                      <button
                        type="submit"
                        disabled={adminSubmitting}
                        className="w-full bg-indigo-600 hover:bg-indigo-700 text-white font-bold py-3 px-4 rounded-xl text-xs shadow-xs transition-colors active:scale-98 disabled:opacity-50"
                      >
                        {adminSubmitting ? 'Mendaftarkan...' : 'Daftarkan Pengguna Baru 🚀'}
                      </button>
                    </form>
                  )}
                </div>

                {/* Dashboard summary info */}
                <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-xs flex flex-col justify-between">
                  <div className="space-y-4">
                    <h3 className="font-bold text-slate-800 text-sm border-b border-slate-100 pb-2 flex items-center gap-1.5">
                      <Award className="w-4 h-4 text-indigo-500" />
                      Statistik Sinkronisasi CERDAS
                    </h3>
                    
                    <div className="space-y-3">
                      <div className="flex justify-between items-center p-3 bg-indigo-50/50 rounded-2xl">
                        <span className="text-xs font-bold text-slate-600">Guru Wali:</span>
                        <span className="text-sm font-extrabold text-indigo-800">{Object.values(allUsers).filter(u => u.role === 'guru_wali').length} Akun</span>
                      </div>
                      
                      <div className="flex justify-between items-center p-3 bg-emerald-50/50 rounded-2xl">
                        <span className="text-xs font-bold text-slate-600">Siswa / Murid:</span>
                        <span className="text-sm font-extrabold text-emerald-800">{Object.values(allUsers).filter(u => u.role === 'murid').length} Siswa</span>
                      </div>

                      <div className="flex justify-between items-center p-3 bg-rose-50/50 rounded-2xl">
                        <span className="text-xs font-bold text-slate-600">Guru BK (Konselor):</span>
                        <span className="text-sm font-extrabold text-rose-800">{Object.values(allUsers).filter(u => u.role === 'guru_bk').length} Akun</span>
                      </div>

                      <div className="flex justify-between items-center p-3 bg-amber-50/50 rounded-2xl">
                        <span className="text-xs font-bold text-slate-600">Total Cerita Tersimpan:</span>
                        <span className="text-sm font-extrabold text-amber-800">{stories.length} Judul</span>
                      </div>
                    </div>

                    <div className="p-3.5 bg-slate-50 rounded-2xl border border-slate-100 space-y-1.5 text-[11px] text-slate-500 leading-relaxed">
                      <strong>✨ Mekanisme Auto-Sync:</strong>
                      <p>Sistem mencocokkan kode <strong>Kelas</strong> (misal: <code>7-A</code>) antara Murid dan Guru Wali. Ketika murid mengunggah cerita refleksi, cerita tersebut akan masuk ke inbox ulasan Guru Wali yang mengajar kelas yang sama.</p>
                    </div>
                  </div>

                  <div className="text-[10px] text-slate-400 text-center italic mt-4">
                    Sistem Roster Digital • Google AI Studio Build
                  </div>
                </div>

              </div>
            )}

            {/* Sub-tab 2: Homeroom Teacher Management (Guru Wali) */}
            {adminSubTab === 'guru_wali' && (
              <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-4">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 border-b border-slate-100 pb-4">
                  <div>
                    <h3 className="font-bold text-slate-800 text-base flex items-center gap-1.5">
                      <span>👩‍🏫 Data Guru Wali</span>
                    </h3>
                    <p className="text-xs text-slate-400 mt-0.5">Guru wali bertanggung jawab memantau progres literasi kelas dan mengulas cerita murid binaannya.</p>
                  </div>
                  
                  <button
                    onClick={handleDeleteAllGuruWali}
                    disabled={Object.values(allUsers).filter(u => u.role === 'guru_wali').length === 0}
                    className="text-xs font-bold text-rose-600 hover:text-rose-800 bg-rose-50 hover:bg-rose-100 px-4 py-2 rounded-xl border border-rose-100 transition-colors disabled:opacity-50"
                  >
                    ⚠️ Hapus Seluruh Guru Wali
                  </button>
                </div>

                <div className="space-y-3 pt-2 max-h-[500px] overflow-y-auto pr-1">
                  {Object.values(allUsers).filter(u => u.role === 'guru_wali').map(u => {
                    if (editingUserUid === u.uid) {
                      return (
                        <div key={u.uid} className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3 text-xs animate-fadeIn">
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                            <div className="space-y-1">
                              <span className="text-[10px] font-bold text-slate-500 block">NAMA GURU WALI:</span>
                              <input
                                type="text"
                                value={editUserName}
                                onChange={(e) => setEditUserName(e.target.value)}
                                className="bg-white border border-slate-200 p-2 rounded-lg w-full font-medium"
                              />
                            </div>
                            <div className="space-y-1">
                              <span className="text-[10px] font-bold text-slate-500 block">EMAIL SEKOLAH:</span>
                              <input
                                type="email"
                                value={editUserEmail}
                                onChange={(e) => setEditUserEmail(e.target.value)}
                                className="bg-white border border-slate-200 p-2 rounded-lg w-full font-medium"
                              />
                            </div>
                            <div className="space-y-1">
                              <span className="text-[10px] font-bold text-slate-500 block">KELAS BINAAN (Pisahkan koma jika lebih dari 1 kelas):</span>
                              <input
                                type="text"
                                value={editUserKelas}
                                placeholder="Contoh: 7B, 7C"
                                onChange={(e) => setEditUserKelas(e.target.value)}
                                className="bg-white border border-slate-200 p-2 rounded-lg w-full font-extrabold text-indigo-700"
                              />
                            </div>
                          </div>
                          <div className="flex justify-end gap-2 text-[10px]">
                            <button
                              type="button"
                              onClick={() => setEditingUserUid(null)}
                              className="px-3.5 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold rounded-lg transition-colors"
                            >
                              Batal
                            </button>
                            <button
                              type="button"
                              onClick={() => handleUpdateUser(u.uid)}
                              className="px-5 py-1.5 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-lg transition-colors"
                            >
                              Simpan Perubahan
                            </button>
                          </div>
                        </div>
                      );
                    }

                    return (
                      <div key={u.uid} className="p-4 bg-slate-50/50 border border-slate-100 rounded-2xl flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs hover:border-indigo-100 transition-colors">
                        <div className="flex items-center gap-3">
                          <img src={u.photoURL} alt={u.displayName} className="w-10 h-10 rounded-full border border-slate-200" />
                          <div>
                            <span className="font-bold text-slate-800 text-sm block">{u.displayName}</span>
                            <span className="text-[10px] text-slate-400 font-mono">{u.email}</span>
                          </div>
                        </div>

                        <div className="flex items-center gap-3">
                          <span className="bg-amber-100 text-amber-900 font-extrabold text-xs px-3 py-1 rounded-xl border border-amber-200">
                            👩‍🏫 Guru Wali {u.kelas}
                          </span>
                          
                          <div className="flex items-center gap-1.5">
                            <button
                              onClick={() => {
                                setEditingUserUid(u.uid);
                                setEditUserName(u.displayName);
                                setEditUserEmail(u.email);
                                setEditUserRole(u.role as any);
                                setEditUserKelas(u.kelas);
                              }}
                              className="text-[10px] font-bold text-teal-700 hover:text-teal-900 px-3 py-2 bg-white hover:bg-teal-50 border border-slate-200 rounded-xl transition-colors shadow-xs"
                            >
                              Edit
                            </button>
                            <button
                              onClick={() => handleDeleteUser(u.uid, u.displayName)}
                              className="text-[10px] font-bold text-rose-700 hover:text-rose-900 px-3 py-2 bg-white hover:bg-rose-50 border border-slate-200 rounded-xl transition-colors shadow-xs"
                            >
                              Hapus
                            </button>
                          </div>
                        </div>
                      </div>
                    );
                  })}

                  {Object.values(allUsers).filter(u => u.role === 'guru_wali').length === 0 && (
                    <p className="text-center text-slate-400 italic py-8 text-sm">Belum ada Guru Wali yang terdaftar di dalam roster.</p>
                  )}
                </div>
              </div>
            )}

            {/* Sub-tab 3: Student Management (Murid) */}
            {adminSubTab === 'murid' && (
              <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-4">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 border-b border-slate-100 pb-4">
                  <div>
                    <h3 className="font-bold text-slate-800 text-base flex items-center gap-1.5">
                      <span>🎒 Data Murid / Siswa</span>
                    </h3>
                    <p className="text-xs text-slate-400 mt-0.5">Daftar siswa yang berhak menulis cerita refleksi 4F dan mencatatkan log pilar kebiasaan harian.</p>
                  </div>
                  
                  <button
                    onClick={handleDeleteAllMurid}
                    disabled={Object.values(allUsers).filter(u => u.role === 'murid').length === 0}
                    className="text-xs font-bold text-rose-600 hover:text-rose-800 bg-rose-50 hover:bg-rose-100 px-4 py-2 rounded-xl border border-rose-100 transition-colors disabled:opacity-50"
                  >
                    ⚠️ Hapus Seluruh Murid
                  </button>
                </div>

                <div className="space-y-3 pt-2 max-h-[500px] overflow-y-auto pr-1">
                  {Object.values(allUsers).filter(u => u.role === 'murid').map(u => {
                    if (editingUserUid === u.uid) {
                      return (
                        <div key={u.uid} className="p-4 bg-slate-50 border border-slate-200 rounded-2xl space-y-3 text-xs animate-fadeIn">
                          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                            <div className="space-y-1">
                              <span className="text-[10px] font-bold text-slate-500 block">NAMA LENGKAP SISWA:</span>
                              <input
                                type="text"
                                value={editUserName}
                                onChange={(e) => setEditUserName(e.target.value)}
                                className="bg-white border border-slate-200 p-2 rounded-lg w-full font-medium"
                              />
                            </div>
                            <div className="space-y-1">
                              <span className="text-[10px] font-bold text-slate-500 block">EMAIL AKUN:</span>
                              <input
                                type="email"
                                value={editUserEmail}
                                onChange={(e) => setEditUserEmail(e.target.value)}
                                className="bg-white border border-slate-200 p-2 rounded-lg w-full font-medium"
                              />
                            </div>
                            <div className="space-y-1">
                              <span className="text-[10px] font-bold text-slate-500 block">KELAS SISWA:</span>
                              <input
                                type="text"
                                value={editUserKelas}
                                onChange={(e) => setEditUserKelas(e.target.value)}
                                className="bg-white border border-slate-200 p-2 rounded-lg w-full font-extrabold text-emerald-700"
                              />
                            </div>
                          </div>
                          <div className="flex justify-end gap-2 text-[10px]">
                            <button
                              type="button"
                              onClick={() => setEditingUserUid(null)}
                              className="px-3.5 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold rounded-lg transition-colors"
                            >
                              Batal
                            </button>
                            <button
                              type="button"
                              onClick={() => handleUpdateUser(u.uid)}
                              className="px-5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg transition-colors"
                            >
                              Simpan Perubahan
                            </button>
                          </div>
                        </div>
                      );
                    }

                    const studentStats = getStudentHabitStats(u.uid);

                    return (
                      <div key={u.uid} className="p-4 bg-slate-50/50 border border-slate-100 rounded-2xl flex flex-col sm:flex-row sm:items-center justify-between gap-3 text-xs hover:border-emerald-100 transition-colors">
                        <div className="flex items-center gap-3">
                          <img src={u.photoURL} alt={u.displayName} className="w-10 h-10 rounded-full border border-slate-200" />
                          <div>
                            <span className="font-bold text-slate-800 text-sm block">{u.displayName}</span>
                            <span className="text-[10px] text-slate-400 font-mono">{u.email}</span>
                          </div>
                        </div>

                        <div className="flex items-center gap-3 flex-wrap sm:flex-nowrap">
                          <span className="bg-emerald-100 text-emerald-900 font-extrabold text-xs px-3 py-1 rounded-xl border border-emerald-200">
                            🎒 Kelas {u.kelas}
                          </span>
                          
                          <div className="text-right px-2 hidden md:block">
                            <span className="text-[10px] font-bold text-slate-500 block">Kepatuhan Karakter:</span>
                            <span className="font-extrabold text-emerald-700 text-xs">{studentStats.score}% ({studentStats.count} hari)</span>
                          </div>

                          <div className="flex items-center gap-1.5">
                            <button
                              onClick={() => {
                                setEditingUserUid(u.uid);
                                setEditUserName(u.displayName);
                                setEditUserEmail(u.email);
                                setEditUserRole(u.role as any);
                                setEditUserKelas(u.kelas);
                              }}
                              className="text-[10px] font-bold text-teal-700 hover:text-teal-900 px-3 py-2 bg-white hover:bg-teal-50 border border-slate-200 rounded-xl transition-colors shadow-xs"
                            >
                              Edit
                            </button>
                            <button
                              onClick={() => handleDeleteUser(u.uid, u.displayName)}
                              className="text-[10px] font-bold text-rose-700 hover:text-rose-900 px-3 py-2 bg-white hover:bg-rose-50 border border-slate-200 rounded-xl transition-colors shadow-xs"
                            >
                              Hapus
                            </button>
                          </div>
                        </div>
                      </div>
                    );
                  })}

                  {Object.values(allUsers).filter(u => u.role === 'murid').length === 0 && (
                    <p className="text-center text-slate-400 italic py-8 text-sm">Belum ada Siswa yang terdaftar di dalam roster.</p>
                  )}
                </div>
              </div>
            )}

            {/* Sub-tab 4: Stories Management (Kelola Cerita) */}
            {adminSubTab === 'stories' && (
              <div className="bg-white p-6 rounded-3xl border border-slate-100 shadow-xs space-y-4">
                <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3 border-b border-slate-100 pb-4">
                  <div>
                    <h3 className="font-bold text-slate-800 text-base flex items-center gap-1.5">
                      <span>📚 Kelola Cerita Kelas (Moderasi)</span>
                    </h3>
                    <p className="text-xs text-slate-400 mt-0.5">Pantau seluruh tulisan anak bimbingan di galeri, moderasi isi cerita, dan edit detail/status jika diperlukan.</p>
                  </div>
                  
                  <button
                    onClick={handleDeleteAllStories}
                    disabled={stories.length === 0}
                    className="text-xs font-bold text-rose-600 hover:text-rose-800 bg-rose-50 hover:bg-rose-100 px-4 py-2 rounded-xl border border-rose-100 transition-colors disabled:opacity-50 animate-pulse"
                  >
                    🔥 Bersihkan Seluruh Cerita Galeri
                  </button>
                </div>

                <div className="space-y-4 pt-2 max-h-[600px] overflow-y-auto pr-1">
                  {stories.map(story => {
                    if (editingStoryId === story.id) {
                      return (
                        <div key={story.id} className="p-5 bg-indigo-50/20 border-2 border-indigo-200 rounded-3xl space-y-4 animate-fadeIn">
                          <h4 className="font-bold text-xs text-indigo-900 uppercase tracking-wider">✏️ Mengedit Cerita: "{story.title}"</h4>
                          
                          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                            <div className="space-y-1">
                              <label className="block text-xs font-bold text-slate-600">Judul Cerita:</label>
                              <input
                                type="text"
                                value={editStoryTitle}
                                onChange={(e) => setEditStoryTitle(e.target.value)}
                                className="bg-white border border-slate-200 p-2.5 rounded-xl w-full text-xs font-bold text-slate-700"
                              />
                            </div>
                            <div className="grid grid-cols-2 gap-2">
                              <div className="space-y-1">
                                <label className="block text-xs font-bold text-slate-600">Kelas Cerita:</label>
                                <input
                                  type="text"
                                  value={editStoryKelas}
                                  onChange={(e) => setEditStoryKelas(e.target.value)}
                                  className="bg-white border border-slate-200 p-2.5 rounded-xl w-full text-xs font-extrabold text-indigo-700"
                                />
                              </div>
                              <div className="space-y-1">
                                <label className="block text-xs font-bold text-slate-600">Status Moderasi:</label>
                                <select
                                  value={editStoryStatus}
                                  onChange={(e: any) => setEditStoryStatus(e.target.value)}
                                  className="bg-white border border-slate-200 p-2.5 rounded-xl w-full text-xs font-bold text-slate-700"
                                >
                                  <option value="draft">📁 Draf / Private</option>
                                  <option value="submitted">⏳ Perlu Ulasan</option>
                                  <option value="reviewed">✅ Selesai Diulas</option>
                                </select>
                              </div>
                            </div>
                          </div>

                          <div className="space-y-1">
                            <label className="block text-xs font-bold text-slate-600">Isi Cerita Utama:</label>
                            <textarea
                              rows={5}
                              value={editStoryContent}
                              onChange={(e) => setEditStoryContent(e.target.value)}
                              className="bg-white border border-slate-200 p-3 rounded-xl w-full text-xs leading-relaxed text-slate-600"
                            />
                          </div>

                          <div className="flex justify-end gap-2 text-[10px]">
                            <button
                              type="button"
                              onClick={() => setEditingStoryId(null)}
                              className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-700 font-bold rounded-xl transition-colors"
                            >
                              Batal
                            </button>
                            <button
                              type="button"
                              onClick={() => handleUpdateStory(story.id)}
                              className="px-6 py-2 bg-indigo-600 hover:bg-indigo-700 text-white font-bold rounded-xl transition-colors shadow-xs"
                            >
                              Simpan Cerita
                            </button>
                          </div>
                        </div>
                      );
                    }

                    return (
                      <div key={story.id} className="p-4 bg-slate-50/40 border border-slate-100 rounded-2xl flex items-start justify-between gap-4 hover:border-indigo-100 transition-colors">
                        <div className="flex gap-3 flex-1 min-w-0">
                          <div className="w-14 h-14 bg-slate-100 rounded-xl overflow-hidden shrink-0 border border-slate-100 flex items-center justify-center">
                            {story.illustrationUrl && story.illustrationUrl.startsWith('data:') ? (
                              <img src={story.illustrationUrl} alt={story.title} className="w-full h-full object-cover" />
                            ) : (
                              <span className="text-2xl">{story.feelingEmoji || '😊'}</span>
                            )}
                          </div>
                          
                          <div className="space-y-1 truncate flex-1 min-w-0">
                            <h4 className="font-bold text-slate-800 text-sm leading-tight truncate">{story.title}</h4>
                            <p className="text-[10px] text-slate-400 flex items-center gap-1.5 flex-wrap">
                              <span>Ditulis oleh: <strong>{story.authorName}</strong></span>
                              <span>·</span>
                              <span className="font-bold bg-slate-100 text-slate-700 px-1.5 py-0.5 rounded-md uppercase text-[9px]">Kelas {story.authorKelas}</span>
                              <span>·</span>
                              {story.status === 'draft' && <span className="bg-slate-100 text-slate-600 font-bold px-1.5 py-0.5 rounded-md uppercase text-[9px]">Draf</span>}
                              {story.status === 'submitted' && <span className="bg-amber-100 text-amber-800 font-bold px-1.5 py-0.5 rounded-md uppercase text-[9px] animate-pulse">Menunggu Ulasan</span>}
                              {story.status === 'reviewed' && <span className="bg-emerald-100 text-emerald-800 font-bold px-1.5 py-0.5 rounded-md uppercase text-[9px]">Sudah Diulas</span>}
                            </p>
                            <p className="text-xs text-slate-500 leading-snug line-clamp-2 pt-1">{story.content}</p>
                          </div>
                        </div>

                        <div className="flex items-center gap-1.5 shrink-0 ml-2">
                          <button
                            onClick={() => {
                              setEditingStoryId(story.id);
                              setEditStoryTitle(story.title);
                              setEditStoryContent(story.content);
                              setEditStoryStatus(story.status);
                              setEditStoryKelas(story.authorKelas);
                            }}
                            className="text-[10px] font-bold text-teal-700 hover:text-teal-900 px-3 py-2 bg-white hover:bg-teal-50 border border-slate-200 rounded-xl transition-colors shadow-xs"
                          >
                            Edit
                          </button>
                          <button
                            onClick={() => handleDeleteStory(story.id)}
                            className="text-[10px] font-bold text-rose-700 hover:text-rose-900 px-3 py-2 bg-white hover:bg-rose-50 border border-slate-200 rounded-xl transition-colors shadow-xs"
                          >
                            Hapus
                          </button>
                        </div>
                      </div>
                    );
                  })}

                  {stories.length === 0 && (
                    <p className="text-center text-slate-400 italic py-8 text-sm">Belum ada cerita siswa se-sekolah yang diterbitkan.</p>
                  )}
                </div>
              </div>
            )}

          </div>
        )}

      </main>

      {/* FOOTER */}
      <footer className="bg-white border-t border-slate-100 mt-12 py-6 text-center text-xs text-slate-400">
        <p>© 2026 CERDAS - Cerita Digital Anak Sempatik. Literasi Hebat, Indonesia Unggul! 🇮🇩✨</p>
        <p className="mt-1">Didukung oleh Google Gemini AI Studio & Firebase</p>
      </footer>

      {/* ======================================= */}
      {/* DIGITAL BOOK READER / STORY DETAIL MODAL */}
      {/* ======================================= */}
      {selectedStory && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto">
          <div className="bg-white rounded-3xl max-w-4xl w-full max-h-[90vh] overflow-y-auto shadow-2xl relative flex flex-col">
            
            <div className="sticky top-0 bg-white border-b border-slate-50 px-6 py-4 flex items-center justify-between z-10">
              <div className="flex items-center gap-2">
                <BookOpen className="w-5 h-5 text-teal-600" />
                <span className="text-sm font-bold text-slate-700">Digital Book Reader CERDAS</span>
              </div>
              <div className="flex items-center gap-3">
                {(selectedStory.authorId === currentUser?.uid || userProfile?.role === 'admin' || userProfile?.role === 'guru_wali' || userProfile?.role === 'guru_bk' || offlineMode) && (
                  <button
                    onClick={() => {
                      const idToDelete = selectedStory.id;
                      setSelectedStory(null);
                      handleDeleteStory(idToDelete);
                    }}
                    className="px-3 py-1.5 bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 rounded-xl font-bold text-xs flex items-center gap-1.5 transition-colors shadow-2xs"
                  >
                    <span>🗑️ Hapus Cerita</span>
                  </button>
                )}
                <button 
                  onClick={() => setSelectedStory(null)}
                  className="text-slate-400 hover:text-slate-700 font-bold text-sm"
                >
                  Tutup (✕)
                </button>
              </div>
            </div>

            <div className="p-6 md:p-8 space-y-8 flex-1">
              
              <div className="relative rounded-2xl h-56 md:h-72 overflow-hidden flex items-center justify-center text-center">
                {selectedStory.illustrationUrl && selectedStory.illustrationUrl.startsWith('data:') ? (
                  <img src={selectedStory.illustrationUrl} alt={selectedStory.title} className="absolute inset-0 w-full h-full object-cover" />
                ) : (
                  <div style={{ backgroundImage: getCuteFallbackBackground(selectedStory.title) }} className="absolute inset-0 animate-pulse" />
                )}
                <div className="absolute inset-0 bg-gradient-to-t from-black/80 via-black/40 to-transparent" />
                
                <div className="relative z-10 text-white p-6 space-y-2 max-w-2xl">
                  <span className="text-4xl block mb-2">{selectedStory.feelingEmoji}</span>
                  <h2 className="text-xl md:text-3xl font-extrabold">{selectedStory.title}</h2>
                  <div className="text-xs md:text-sm text-slate-200 flex items-center justify-center gap-2">
                    <span>Ditulis oleh: <strong>{selectedStory.authorName}</strong></span>
                    <span>·</span>
                    <span>Kelas {selectedStory.authorKelas}</span>
                  </div>
                </div>
              </div>

              <div className="space-y-4">
                <h3 className="text-lg font-bold text-slate-800 border-b border-slate-100 pb-2">📚 Isi Cerita</h3>
                <p className="text-sm text-slate-600 leading-relaxed whitespace-pre-wrap font-medium">
                  {selectedStory.content}
                </p>
              </div>

              {/* 4F Reflection */}
              <div className="space-y-4 pt-4 border-t border-slate-50">
                <h3 className="text-lg font-bold text-slate-800">🧩 Refleksi 4F Cerdas</h3>
                
                <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                  <div className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-1">
                    <span className="text-xs font-bold text-teal-800 block">FACT (Fakta / Peristiwa)</span>
                    <p className="text-xs text-slate-600 leading-relaxed">{selectedStory.fact}</p>
                  </div>

                  <div className="p-4 bg-rose-50/50 rounded-2xl border border-rose-100/50 space-y-1">
                    <span className="text-xs font-bold text-rose-800 block">FEELING (Perasaan {selectedStory.feelingEmoji})</span>
                    <p className="text-xs text-slate-600 leading-relaxed">{selectedStory.feeling || 'Merasa damai dan bersyukur atas pelajaran hari ini.'}</p>
                  </div>

                  <div className="p-4 bg-amber-50/50 rounded-2xl border border-amber-100/50 space-y-1">
                    <span className="text-xs font-bold text-amber-800 block">FINDING (Temuan / Pembelajaran)</span>
                    <p className="text-xs text-slate-600 leading-relaxed">{selectedStory.finding}</p>
                  </div>

                  <div className="p-4 bg-blue-50/50 rounded-2xl border border-blue-100/50 space-y-1">
                    <span className="text-xs font-bold text-blue-800 block">FUTURE (Masa Depan / Penerapan)</span>
                    <p className="text-xs text-slate-600 leading-relaxed">{selectedStory.future}</p>
                  </div>
                </div>
              </div>

              <div className="space-y-2 pt-4 border-t border-slate-50">
                <span className="text-xs font-bold text-slate-400 uppercase block tracking-wider">7 Kebiasaan Anak Indonesia Hebat Terkait:</span>
                <div className="flex flex-wrap gap-2 pt-1 text-xs text-slate-600 font-semibold">
                  {selectedStory.associatedHabits.map(habit => {
                    const found = HABITS_INFO.find(h => h.id === habit);
                    return (
                      <span key={habit} className="bg-emerald-50 text-emerald-800 px-3 py-1 rounded-lg border border-emerald-100">
                        {found?.icon || '🌱'} {found?.label || habit}
                      </span>
                    );
                  })}
                </div>
              </div>

              {/* Si CERDAS AI feedback */}
              {(selectedStory.aiSummary || selectedStory.aiFeedback) && (
                <div className="p-4 bg-teal-50 rounded-2xl border border-teal-100/50 space-y-2">
                  <div className="flex items-center gap-1.5 text-teal-800 font-bold text-xs">
                    <Sparkles className="w-4 h-4 text-teal-600" />
                    <span>Catatan Si CERDAS AI</span>
                  </div>
                  {selectedStory.aiSummary && (
                    <p className="text-[11px] text-slate-500 leading-tight">
                      <strong>Ringkasan AI:</strong> "{selectedStory.aiSummary}"
                    </p>
                  )}
                  {selectedStory.aiFeedback && (
                    <p className="text-xs text-teal-800 leading-relaxed italic">
                      "Masukan AI: {selectedStory.aiFeedback}"
                    </p>
                  )}
                </div>
              )}

              {/* List Feedbacks */}
              <div className="space-y-4 pt-4 border-t border-slate-50">
                <h3 className="text-lg font-bold text-slate-800">💬 Ulasan Guru & BK ({feedbacks[selectedStory.id]?.length || 0})</h3>
                
                <div className="space-y-3">
                  {feedbacks[selectedStory.id]?.map(fb => (
                    <div key={fb.id} className="p-4 bg-slate-50 rounded-2xl border border-slate-100 space-y-2">
                      <div className="flex items-center justify-between">
                        <div className="flex items-center gap-1.5">
                          <span className="text-xs font-bold text-slate-800">{fb.reviewerName}</span>
                          <span className="text-[10px] font-semibold px-2 py-0.5 bg-teal-50 text-teal-700 rounded-full">
                            {fb.reviewerRole === 'guru_wali' ? 'Guru Wali' : fb.reviewerRole === 'guru_bk' ? 'Guru BK' : 'Guru Mapel'}
                          </span>
                        </div>
                        
                        <div className="flex items-center text-amber-500">
                          {Array.from({ length: fb.stars }).map((_, i) => (
                            <Star key={i} className="w-3.5 h-3.5 fill-amber-500" />
                          ))}
                        </div>
                      </div>

                      <p className="text-xs text-slate-600 leading-relaxed font-medium">
                        {fb.feedbackText}
                      </p>
                    </div>
                  ))}

                  {(!feedbacks[selectedStory.id] || feedbacks[selectedStory.id].length === 0) && (
                    <p className="text-slate-400 text-xs italic">Belum ada ulasan dari bapak/ibu guru untuk cerita ini.</p>
                  )}
                </div>
              </div>

              {/* Add Feedback Form */}
              {(userProfile?.role === 'guru_wali' || userProfile?.role === 'guru_non_wali' || userProfile?.role === 'guru_bk') && (
                <form onSubmit={handleSubmitFeedback} className="p-4 bg-amber-50/30 border border-amber-100/50 rounded-2xl space-y-4">
                  <h4 className="font-bold text-sm text-slate-800 flex items-center gap-1">
                    <Award className="w-4 h-4 text-amber-500" /> Berikan Apresiasi Guru
                  </h4>

                  <div className="space-y-2">
                    <label className="block text-xs font-bold text-slate-600">Peringkat Bintang:</label>
                    <div className="flex items-center gap-1">
                      {[1, 2, 3, 4, 5].map(starNum => (
                        <button
                          key={starNum}
                          type="button"
                          onClick={() => setNewFeedbackStars(starNum)}
                          className="text-amber-400 transition-transform active:scale-110"
                        >
                          <Star className={`w-6 h-6 ${newFeedbackStars >= starNum ? 'fill-amber-400' : 'text-slate-300'}`} />
                        </button>
                      ))}
                    </div>
                  </div>

                  <div className="space-y-2">
                    <label className="block text-xs font-bold text-slate-600">Catatan Bimbingan / Apresiasi:</label>
                    <textarea
                      rows={3}
                      required
                      placeholder="Wah, kisah yang sangat inspiratif anak hebat! Bapak sangat bangga kamu mau berbagi apel dengan Dika. Teruskan kebiasaan gemar berbagi ya..."
                      value={newFeedbackText}
                      onChange={(e) => setNewFeedbackText(e.target.value)}
                      className="w-full bg-white border border-slate-200 p-3 rounded-xl focus:outline-none focus:ring-2 focus:ring-amber-400 text-xs"
                    />
                  </div>

                  <button
                    type="submit"
                    disabled={submittingFeedback}
                    className="w-full bg-amber-500 hover:bg-amber-600 text-white font-bold py-2 px-4 rounded-xl text-xs shadow-xs transition-colors active:scale-98"
                  >
                    {submittingFeedback ? 'Mengirim Apresiasi...' : 'Kirim Ulasan & Berikan Bintang! 🌟'}
                  </button>
                </form>
              )}

            </div>

          </div>
        </div>
      )}

      {/* Microphone Permission Guide / Alternative Audio Upload Modal */}
      {micPermissionDeniedModal && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 overflow-y-auto animate-fadeIn">
          <div className="bg-white rounded-3xl max-w-lg w-full p-6 md:p-8 shadow-2xl space-y-6 relative border border-slate-100">
            <button
              onClick={() => setMicPermissionDeniedModal(false)}
              className="absolute top-4 right-4 text-slate-400 hover:text-slate-700 font-bold text-sm p-2 hover:bg-slate-100 rounded-full"
            >
              ✕
            </button>

            <div className="flex items-center gap-3">
              <div className="p-3 bg-amber-100 text-amber-800 rounded-2xl">
                <Mic className="w-6 h-6 text-amber-600 animate-pulse" />
              </div>
              <div>
                <h3 className="text-lg font-bold text-slate-800">Akses Mikrofon Belum Diizinkan 🎙️</h3>
                <p className="text-xs text-slate-500">Izin mikrofon browser perlu diaktifkan atau gunakan alternatif unggah file audio di bawah ini.</p>
              </div>
            </div>

            <div className="space-y-3 bg-amber-50/60 p-4 rounded-2xl border border-amber-100 text-xs text-slate-700 leading-relaxed">
              <span className="font-bold text-amber-900 block">💡 Cara Mengaktifkan Izin Mikrofon:</span>
              <ol className="list-decimal list-inside space-y-1 text-slate-600">
                <li>Klik ikon gembok 🔒 atau Pengaturan Situs di samping alamat URL browser Anda.</li>
                <li>Ubah setelan <strong>"Mikrofon"</strong> menjadi <strong>"Izinkan" (Allow)</strong>.</li>
                <li>Muat ulang (refresh) halaman aplikasi CERDAS ini.</li>
              </ol>
            </div>

            <div className="space-y-3 pt-2 border-t border-slate-100">
              <div className="text-xs font-bold text-slate-800 flex items-center gap-1.5">
                <Sparkles className="w-4 h-4 text-indigo-500" />
                <span>Alternatif: Unggah File Rekaman Suara (.mp3/.wav/.m4a)</span>
              </div>
              <p className="text-[11px] text-slate-500 leading-relaxed">
                Anda dapat merekam suara menggunakan perekam bawaan HP/Laptop, lalu unggah filenya di bawah ini. Gemini AI akan mentranskripsikannya ke dalam teks secara otomatis!
              </p>

              <div className="flex flex-col gap-2">
                <label className="cursor-pointer bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs py-3 px-4 rounded-xl text-center shadow-xs transition-colors flex items-center justify-center gap-2">
                  <span>📁 Pilih File Audio Rekaman Suara</span>
                  <input
                    type="file"
                    accept="audio/*,.mp3,.wav,.m4a,.webm,.aac"
                    onChange={(e) => {
                      setMicPermissionDeniedModal(false);
                      handleAudioFileUpload(e, 'content');
                    }}
                    className="hidden"
                  />
                </label>

                <button
                  type="button"
                  onClick={() => setMicPermissionDeniedModal(false)}
                  className="w-full bg-slate-100 hover:bg-slate-200 text-slate-700 font-bold text-xs py-2.5 rounded-xl transition-colors"
                >
                  Tutup & Ketik Cerita Secara Manual
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

    </div>
  );
}
