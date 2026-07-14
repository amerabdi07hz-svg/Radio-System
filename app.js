import { initializeApp } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-app.js";
import { getAuth, signInWithEmailAndPassword, setPersistence, browserSessionPersistence } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-auth.js";
import { getFirestore } from "https://www.gstatic.com/firebasejs/10.12.2/firebase-firestore.js";

// إعدادات فايربيس الخاصة بك
const firebaseConfig = {
    apiKey: "AIzaSyCAKw7quIPloQtMMdt4pp3aeezAKzx39hA",
    authDomain: "radio-system-6aaad.firebaseapp.com",
    projectId: "radio-system-6aaad",
    storageBucket: "radio-system-6aaad.firebasestorage.app",
    messagingSenderId: "528567280152",
    appId: "1:528567280152:web:724638a00da58fc2c2196d"
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const db = getFirestore(app);

// جلسة مؤقتة - تنتهي عند إغلاق التبويب
setPersistence(auth, browserSessionPersistence);

document.getElementById('login-form').addEventListener('submit', async function(e) {
    e.preventDefault();
    const empId = document.getElementById('empId').value;
    const password = document.getElementById('password').value;
    const errorMessage = document.getElementById('error-message');
    const loginBtn = document.getElementById('login-btn');
    
    errorMessage.style.display = 'none';
    loginBtn.textContent = 'جاري الاتصال بالسحابة...';
    loginBtn.disabled = true;

    // استخدام بريد وهمي مبني على رقم التوظيف ليتوافق مع فايربيس
    const email = `${empId}@radio.local`; 

    try {
        // محاولة تسجيل الدخول
        await signInWithEmailAndPassword(auth, email, password);
        window.location.href = "dashboard.html";
    } catch (error) {
        errorMessage.style.display = 'block';
        errorMessage.textContent = 'رقم التوظيف أو كلمة السر غير صحيحة!';
        
        loginBtn.textContent = 'تسجيل الدخول';
        loginBtn.disabled = false;
    }
});