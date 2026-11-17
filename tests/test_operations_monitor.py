from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path: sys.path.insert(0,str(ROOT))
import unittest
from datetime import date
import pandas as pd
from operations_monitor import collect_extended_findings


class ExtendedOperationsMonitorTests(unittest.TestCase):
    def test_detects_cross_module_control_gaps(self):
        data = {
            "risk": pd.DataFrame([{"Risk ID":"R-01", "Risk Tanımı":"Tedarik riski", "Olasılık (1-5)":5,
                "Etki (1-5)":4, "Risk Skoru":20, "Önlem / Aksiyon":"", "Sorumlu":"", "Durum":"Açık", "Hedef Tarih":"2026-11-01"}]),
            "contracts": pd.DataFrame([{"Sözleşme No":"C-01", "Durum":"Aktif", "Bitiş Tarihi":"2026-11-10",
                "Orijinal Bedel":100000, "Para Birimi":"TRY"}]),
            "changes": pd.DataFrame([{"Değişiklik No":"VO-01", "Durum":"Bekliyor", "Talep Tarihi":"2026-10-20"}]),
            "hakediş": pd.DataFrame([{"Hakediş No":"H-01", "WBS":"", "Dönem":"2026-10", "Ödenecek Tutar":100,
                "Ödenen":150}]),
            "employer_hakediş": pd.DataFrame([{"Hakediş No":"İH-01", "Kalan Alacak":50000, "Vade Tarihi":"2026-11-01"}]),
            "approvals": pd.DataFrame([{"_id":"A-01", "Durum":"Bekliyor", "Talep Tarihi":"2026-11-01", "Onay Rolü":""}]),
            "cash": pd.DataFrame([{"_id":"N-01", "Durum":"Bekliyor", "Tutar":2000, "Para Birimi":"TRY", "Vade Tarihi":"2026-11-02"}]),
            "daily_reports": pd.DataFrame([{"_id":"D-01", "Rapor Tarihi":"2026-11-15", "Raporlayan":"", "Günlük İmalat Özeti":""}]),
            "labor": pd.DataFrame([{"_id":"L-01", "Çalışan Sayısı":5, "Gün":2, "Kişi Günlük Ücret":1000, "Toplam İşçilik":3000}]),
        }
        findings = collect_extended_findings(data, date(2026, 11, 16))
        modules = {x["Modül"] for x in findings}
        self.assertIn("Risk Yönetimi", modules)
        self.assertIn("Sözleşme", modules)
        self.assertIn("Sözleşme Değişikliği", modules)
        self.assertIn("Taşeron Hakedişi", modules)
        self.assertIn("İşveren Hakedişi", modules)
        self.assertIn("Onay Merkezi", modules)
        self.assertIn("Nakit / Ödeme", modules)
        self.assertIn("Günlük Şantiye Raporu", modules)
        self.assertIn("İşçilik / Puantaj", modules)

    def test_closed_risk_is_not_reopened(self):
        data = {"risk": pd.DataFrame([{"Risk ID":"R-closed", "Risk Skoru":25, "Durum":"Kapalı"}])}
        self.assertEqual(collect_extended_findings(data, date(2026,11,16)), [])

    def test_no_data_returns_no_false_positive(self):
        self.assertEqual(collect_extended_findings({}, date(2026,11,16)), [])

    def test_duplicate_claim_and_cash_documents_are_detected(self):
        data = {
            "hakediş": pd.DataFrame([{"Hakediş No":"H-1"},{"Hakediş No":"H-1"}]),
            "cash": pd.DataFrame([{"Belge No":"F-1","Durum":"Ödendi"},{"Belge No":"F-1","Durum":"Ödendi"}]),
        }
        findings = collect_extended_findings(data, date(2026,11,16))
        self.assertTrue(any("mükerrer" in x["Tespit"] for x in findings))


if __name__ == "__main__":
    unittest.main()
