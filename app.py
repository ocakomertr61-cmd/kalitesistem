import os
import re
import shutil
import logging
from datetime import datetime
from typing import Dict, Any, List, Optional

# Logging Ayarları
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[logging.StreamHandler()]
)

class QualityDocumentSystem:
    """
    Kalite Yönetim Sistemi - Doküman, Revizyon ve Arşiv Yönetim Modülü
    
    Çalışma Prensibi: 'Single Source of Truth' (Panel Verisi Esastır)
    - Yüklenen fiziksel dosyanın adı tamamen yoksayılır (sadece uzantısı alınır).
    - Dosya adı, panelde girilen Doküman No, Doküman Adı ve Revizyon No ile otomatik üretilir.
    - Canlıdaki eski revizyon otomatik olarak tespit edilir ve güvenli şekilde arşivlenir.
    - Eksik/hatalı sütun veya form verilerinde otomatik doğrulama ve önleyici hata yönetimi içerir.
    """

    def __init__(self, active_dir: str = "active_documents", archive_dir: str = "archived_documents"):
        self.active_dir = active_dir
        self.archive_dir = archive_dir
        self._ensure_directories()

    def _ensure_directories(self) -> None:
        """Gerekli klasör yapısını kontrol eder ve yoksa oluşturur."""
        os.makedirs(self.active_dir, exist_ok=True)
        os.makedirs(self.archive_dir, exist_ok=True)
        logging.info(f"Sistem klasörleri hazır: Canlı='{self.active_dir}', Arşiv='{self.archive_dir}'")

    @staticmethod
    def clean_text_for_filename(text: Any) -> str:
        """
        Dosya adlarında uyumsuzluk veya işletim sistemi hatası yaratacak
        karakterleri temizler, Türkçe karakterleri uyarlar ve boşlukları düzenler.
        """
        if text is None:
            return ""
        text_str = str(text).strip()
        # Türkçe karakter dönüştürme ve geçersiz karakterlerin temizlenmesi
        tr_map = str.maketrans("çğıöşüÇĞİÖŞÜ", "cgiosuCGIOSU")
        clean_str = text_str.translate(tr_map)
        # Dosya sisteminde yasaklı karakterleri kaldır
        clean_str = re.sub(r'[\\/*?:"<>|]', "", clean_str)
        # Boşlukları ve birden fazla alt çizgiyi tek alt çizgiye çevir
        clean_str = re.sub(r'\s+', "_", clean_str)
        clean_str = re.sub(r'_+', "_", clean_str)
        return clean_str

    @staticmethod
    def format_revision_number(rev_input: Any) -> str:
        """
        Revizyon numarasını standart 2 haneli formata getirir (Örn: 0 -> '00', 1 -> '01', 2 -> '02').
        """
        try:
            rev_int = int(str(rev_input).strip())
            return f"{rev_int:02d}"
        except (ValueError, TypeError):
            # Sayıya çevrilemiyorsa temizleyip metin olarak döndürür
            return str(rev_input).strip()

    def generate_standard_filename(self, doc_no: str, doc_name: str, rev_no: Any, file_extension: str) -> str:
        """
        Panel verilerini baz alarak standart dosya adını oluşturur.
        Format: [DOKUMAN_NO]_[DOKUMAN_ADI]_R[REV_NO].[uzanti]
        Örnek: FRM-045_Gelen_Malzeme_Kontrol_Formu_R02.pdf
        """
        clean_no = self.clean_text_for_filename(doc_no)
        clean_name = self.clean_text_for_filename(doc_name)
        formatted_rev = self.format_revision_number(rev_no)
        
        ext = file_extension if file_extension.startswith(".") else f".{file_extension}"
        
        standard_filename = f"{clean_no}_{clean_name}_R{formatted_rev}{ext}"
        return standard_filename

    def validate_panel_data(self, panel_data: Dict[str, Any]) -> Dict[str, Any]:
        """
        Panel üzerinden gelen form girdilerini doğrular, eksik alan kontrolü yapar.
        """
        required_fields = ["doc_no", "doc_name", "rev_no"]
        missing_fields = [field for field in required_fields if not panel_data.get(field) and panel_data.get(field) != 0]

        if missing_fields:
            raise ValueError(f"HATA: Form üzerinde eksik zorunlu alanlar var: {', '.join(missing_fields)}")

        validated = {
            "doc_no": str(panel_data["doc_no"]).strip(),
            "doc_name": str(panel_data["doc_name"]).strip(),
            "rev_no": panel_data["rev_no"],
            "author": panel_data.get("author", "Ömer Ocak"),
            "department": panel_data.get("department", "Kalite Güvence"),
            "notes": panel_data.get("notes", ""),
            "process_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }
        return validated

    def find_active_revisions(self, doc_no: str) -> List[str]:
        """
        Canlı klasörde girilen Doküman No ile başlayan mevcut tüm revizyon dosyalarını bulur.
        """
        clean_no = self.clean_text_for_filename(doc_no)
        prefix = f"{clean_no}_"
        
        active_files = []
        if os.path.exists(self.active_dir):
            for fname in os.listdir(self.active_dir):
                if fname.startswith(prefix):
                    active_files.append(fname)
        return active_files

    def archive_existing_file(self, filename: str) -> str:
        """
        Canlı klasördeki eski revizyon dosyasını zaman damgası ekleyerek arşiv klasörüne taşır.
        """
        source_path = os.path.join(self.active_dir, filename)
        if not os.path.exists(source_path):
            raise FileNotFoundError(f"Arşivlenecek dosya canlı klasörde bulunamadı: {source_path}")

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        archived_filename = f"ARCHIVED_{timestamp}_{filename}"
        target_path = os.path.join(self.archive_dir, archived_filename)

        shutil.move(source_path, target_path)
        logging.info(f"[ARŞİVLENDİ] Eski revizyon arşive taşındı: '{filename}' -> '{archived_filename}'")
        return archived_filename

    def process_document_upload(self, panel_data: Dict[str, Any], uploaded_file_path: str) -> Dict[str, Any]:
        """
        Ana İş Akışı:
        1. Form verilerini doğrular (Single Source of Truth).
        2. Yüklenen dosyanın orijinal adını yoksayar, sadece uzantısını alır.
        3. Panel verileriyle standart dosya adını üretir.
        4. Canlıdaki eski revizyonu bulup arşive kaldırır.
        5. Yeni dosyayı panel adıyla canlıya kaydeder.
        6. Veritabanı/Sistem kayıt nesnesi (Metadata) üretip döndürür.
        """
        # 1. Panel Veri Doğrulama
        validated_data = self.validate_panel_data(panel_data)
        
        # Fiziksel Yükleme Dosya Kontrolü
        if not os.path.exists(uploaded_file_path):
            raise FileNotFoundError(f"HATA: Yüklenen kaynak dosya bulunamadı: {uploaded_file_path}")

        # 2. Orijinal Dosya Adı Yoksayılır -> Sadece Uzantı Alınır
        original_basename = os.path.basename(uploaded_file_path)
        _, ext = os.path.splitext(uploaded_file_path)

        # 3. Panel Verilerinden Standart İsim Üretme
        standard_filename = self.generate_standard_filename(
            doc_no=validated_data["doc_no"],
            doc_name=validated_data["doc_name"],
            rev_no=validated_data["rev_no"],
            file_extension=ext
        )
        target_active_path = os.path.join(self.active_dir, standard_filename)

        # 4. Mevcut Revizyon Kontrolü ve Arşivleme
        existing_active_files = self.find_active_revisions(validated_data["doc_no"])
        archived_records = []

        for old_file in existing_active_files:
            if old_file == standard_filename:
                logging.warning(f"Aynı revizyon ({standard_filename}) canlıda mevcut. Üzerine yazılacak.")
            archived_name = self.archive_existing_file(old_file)
            archived_records.append({
                "original_active_name": old_file,
                "archived_as": archived_name,
                "archive_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

        # 5. Yeni Dosyayı Panel Tabanlı Adıyla Canlı Klasöre Kaydetme
        shutil.copy2(uploaded_file_path, target_active_path)
        logging.info(f"[CANLIYA ALINDI] Dosya başarıyla yüklendi: '{standard_filename}'")

        # 6. Sistem Kayıt Metadatasının Oluşturulması
        result_record = {
            "status": "SUCCESS",
            "doc_no": validated_data["doc_no"],
            "doc_name": validated_data["doc_name"],
            "rev_no": validated_data["rev_no"],
            "system_file_name": standard_filename,
            "original_uploaded_file_name": original_basename,
            "active_file_path": target_active_path,
            "archived_previous_revisions": archived_records,
            "processed_by": validated_data["author"],
            "department": validated_data["department"],
            "notes": validated_data["notes"],
            "process_timestamp": validated_data["process_date"]
        }

        return result_record


# ==============================================================================
# ÖRNEK KULLANIM VE SİMÜLASYON TESTİ
# ==============================================================================
if __name__ == "__main__":
    print("--- Kalite Doküman ve Revizyon Yönetim Sistemi Testi ---")
    
    system = QualityDocumentSystem()

    # --- SENARYO 1: REVİZYON 01 YÜKLEMESİ ---
    print("\n>>> Senaryo 1: Revizyon 01 Yükleniyor...")
    user_file_v1 = "temp_taslak_v1_son_hali_final.pdf"
    with open(user_file_v1, "w", encoding="utf-8") as f:
        f.write("Revizyon 01 İçeriği - Gelen Malzeme Kontrol")

    panel_data_r1 = {
        "doc_no": "FRM-045",
        "doc_name": "Gelen Malzeme Kontrol Formu",
        "rev_no": "01",
        "author": "Ömer Ocak",
        "notes": "İlk yayınlanan revizyon."
    }

    res1 = system.process_document_upload(panel_data_r1, user_file_v1)
    print("Sonuç 1 Dosya Adı:", res1["system_file_name"])

    # Geçici yükleme dosyasını temizle
    if os.path.exists(user_file_v1):
        os.remove(user_file_v1)

    # --- SENARYO 2: REVİZYON 02 YÜKLEMESİ (ARŞİVLEME SİMÜLASYONU) ---
    print("\n>>> Senaryo 2: Revizyon 02 Yükleniyor (R01 Otomatik Arşivlenmeli)...")
    user_file_v2 = "SCAN_20260909_987654.pdf"
    with open(user_file_v2, "w", encoding="utf-8") as f:
        f.write("Revizyon 02 İçeriği - Güncellenmiş Malzeme Kontrol")

    panel_data_r2 = {
        "doc_no": "FRM-045",
        "doc_name": "Gelen Malzeme Kontrol Formu",
        "rev_no": "02",
        "author": "Ömer Ocak",
        "notes": "Ölçüm kriterleri güncellendi."
    }

    res2 = system.process_document_upload(panel_data_r2, user_file_v2)
    print("Sonuç 2 Dosya Adı:", res2["system_file_name"])
    print("-> Canlı Klasör İçeriği:", os.listdir(system.active_dir))
    print("-> Arşiv Klasör İçeriği:", os.listdir(system.archive_dir))

    # Geçici yükleme dosyasını temizle
    if os.path.exists(user_file_v2):
        os.remove(user_file_v2)
