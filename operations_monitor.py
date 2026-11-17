# -*- coding: utf-8 -*-
"""Deterministic cross-module exception checks for Emir Asel ERP.

The monitor is read-only. Findings are work items for authorized human review;
it never edits, approves, posts, deletes, or settles source records.
"""
from __future__ import annotations
from datetime import date, datetime
from typing import Any
import math
import pandas as pd


def _frame(data: dict[str, Any], key: str) -> pd.DataFrame:
    value = data.get(key)
    return value.copy() if isinstance(value, pd.DataFrame) else pd.DataFrame()


def _s(value: Any) -> str:
    return "" if value is None or pd.isna(value) else str(value).strip()


def _n(value: Any, default: float = 0.0) -> float:
    try:
        x = float(str(value).replace(" ", "").replace(",", "."))
        return x if math.isfinite(x) else default
    except (TypeError, ValueError):
        return default


def _d(value: Any):
    # ISO dates are parsed explicitly; Turkish day-first inputs use a separate pass.
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return pd.NaT
    raw = str(value).strip()
    if not raw:
        return pd.NaT
    iso = pd.to_datetime(raw, format="%Y-%m-%d", errors="coerce")
    if not pd.isna(iso):
        return iso
    return pd.to_datetime(raw, errors="coerce", dayfirst=True)


def _closed(value: Any) -> bool:
    return _s(value).casefold().replace("ı", "i") in {
        "kapali", "closed", "tamamlandi", "tamamlandı", "iptal", "reddedildi",
        "odendi", "ödendi", "tahsil edildi", "sonuclandi", "sonuçlandı", "feshedildi",
    }


def collect_extended_findings(data: dict[str, Any], report_date: date) -> list[dict[str, str]]:
    """Return deterministic findings for risk, contracts, changes, claims, cash and approvals."""
    findings: list[dict[str, str]] = []

    def add(level: str, module: str, record: Any, detail: str, action: str) -> None:
        findings.append({"Önem": level, "Modül": module, "Kayıt": _s(record)[:120],
                        "Tespit": detail[:600], "Önerilen İşlem": action[:600]})

    # Risk register: a risk without owner, action or target date is not controllable.
    risks = _frame(data, "risk")
    for _, row in risks.iterrows():
        if _closed(row.get("Durum")):
            continue
        risk_ref = row.get("Risk ID") or row.get("_id") or row.get("Risk Tanımı")
        score = _n(row.get("Risk Skoru"), _n(row.get("Olasılık (1-5)")) * _n(row.get("Etki (1-5)")))
        target = _d(row.get("Hedef Tarih"))
        if score >= 15:
            add("YÜKSEK", "Risk Yönetimi", risk_ref, f"Açık risk skoru {score:g}; yüksek/çok yüksek risk.",
                "Risk sahibini, azaltıcı önlemi, hedef tarihini ve maliyet/süre etkisini yönetim kararıyla takip edin.")
        if not _s(row.get("Sorumlu")):
            add("ORTA", "Risk Yönetimi", risk_ref, "Açık risk için sorumlu atanmamış.", "Yetkili bir risk sahibi atayın.")
        if not _s(row.get("Önlem / Aksiyon")):
            add("ORTA", "Risk Yönetimi", risk_ref, "Açık risk için azaltıcı aksiyon tanımlanmamış.", "Ölçülebilir önlem, sorumlu ve kapanış kanıtı tanımlayın.")
        if pd.isna(target):
            add("ORTA", "Risk Yönetimi", risk_ref, "Açık riskte hedef tarih eksik.", "Risk aksiyonu için hedef tarih girin.")
        elif target.date() < report_date:
            add("YÜKSEK", "Risk Yönetimi", risk_ref, f"Risk aksiyon hedef tarihi geçmiş ({target.date().isoformat()}).",
                "Aksiyonun gerçekleşmesini doğrulayın; gerekirse yeni hedef tarih ve eskalasyon kaydı açın.")

    # Contracts: expiration and financial terms are monitored separately from change orders.
    contracts = _frame(data, "contracts")
    if not contracts.empty:
        numbers = contracts[contracts.get("Sözleşme No", pd.Series(index=contracts.index, dtype=object)).fillna("").astype(str).str.strip().ne("")]
        duplicate_numbers = numbers[numbers["Sözleşme No"].astype(str).str.strip().duplicated(keep=False)] if not numbers.empty else pd.DataFrame()
        for number in sorted(duplicate_numbers["Sözleşme No"].astype(str).unique()) if not duplicate_numbers.empty else []:
            add("KRİTİK", "Sözleşme", number, "Sözleşme numarası birden fazla kayıtta kullanılmış.",
                "Belge ve sözleşme kayıtlarını karşılaştırın; mükerrer veya revizyon kaydını ayırın.")
        for _, row in contracts.iterrows():
            status = _s(row.get("Durum"))
            if _closed(status) or status.casefold() in {"taslak", "draft"}:
                continue
            ref = row.get("Sözleşme No") or row.get("_id")
            end = _d(row.get("Bitiş Tarihi"))
            if pd.isna(end):
                add("ORTA", "Sözleşme", ref, "Aktif sözleşmede bitiş tarihi girilmemiş.",
                    "İmzalı sözleşmedeki yürürlük/bitiş koşulunu kaydedin.")
            elif end.date() < report_date:
                add("YÜKSEK", "Sözleşme", ref, f"Sözleşme bitiş tarihi geçmiş ({end.date().isoformat()}); durum açık.",
                    "Uzatma, fesih veya kapanış belgesini doğrulayın; satın alma ve hakediş etkisini gözden geçirin.")
            elif (end.date() - report_date).days <= 30:
                add("ORTA", "Sözleşme", ref, f"Sözleşmenin bitişine {(end.date() - report_date).days} gün kaldı.",
                    "Termin, uzatma opsiyonu, teminat ve açık hakedişleri kontrol edin.")
            if _n(row.get("Orijinal Bedel"), -1) < 0 or _n(row.get("Orijinal Bedel"), 0) == 0:
                add("ORTA", "Sözleşme", ref, "Sözleşme bedeli sıfır veya eksik.", "İmzalı sözleşmedeki bedeli, para birimini ve kur bazını kontrol edin.")
            if not _s(row.get("Para Birimi")):
                add("ORTA", "Sözleşme", ref, "Sözleşme para birimi belirtilmemiş.", "Sözleşme para birimini ve varsa kur kaynağını tanımlayın.")

    # Change orders: unresolved variation requests age; approved changes need approval trace.
    changes = _frame(data, "changes")
    for _, row in changes.iterrows():
        status = _s(row.get("Durum"))
        if _closed(status):
            continue
        ref = row.get("Değişiklik No") or row.get("_id")
        requested = _d(row.get("Talep Tarihi"))
        if status.casefold() in {"bekliyor", "talep edildi", "incelemede", "kontrol bekliyor", "onay bekliyor"}:
            age = (report_date - requested.date()).days if pd.notna(requested) else None
            if age is None:
                add("ORTA", "Sözleşme Değişikliği", ref, "Değişiklik talebinin tarihi eksik.", "Talep tarihini ve RFI/ek belge bağlantısını tamamlayın.")
            elif age > 7:
                add("YÜKSEK", "Sözleşme Değişikliği", ref, f"Değişiklik talebi {age} gündür sonuçlanmamış.",
                    "Maliyet, süre ve kapsam etkisini değerlendirip yetkili karara taşıyın.")
        if status.casefold() in {"onaylandi", "onaylandı"} and (not _s(row.get("Onay Tarihi")) or not _s(row.get("Onaylayan"))):
            add("YÜKSEK", "Sözleşme Değişikliği", ref, "Onaylı değişiklikte onay tarihi veya onaylayan izi eksik.",
                "Onay kararını belge/kişi/tarih bilgisiyle tamamlayın; tutarı ve süre etkisini baseline revizyonuna bağlayın.")

    # Subcontractor progress payments: duplicate IDs and impossible paid totals.
    subclaims = _frame(data, "hakediş")
    if not subclaims.empty:
        ids = subclaims[subclaims.get("Hakediş No", pd.Series(index=subclaims.index, dtype=object)).fillna("").astype(str).str.strip().ne("")]
        duplicates = ids[ids["Hakediş No"].astype(str).str.strip().duplicated(keep=False)] if not ids.empty else pd.DataFrame()
        for number in sorted(duplicates["Hakediş No"].astype(str).unique()) if not duplicates.empty else []:
            add("KRİTİK", "Taşeron Hakedişi", number, "Hakediş numarası mükerrer.", "Dönem, taşeron, WBS ve hakediş evrakını mutabık kılın.")
        for _, row in subclaims.iterrows():
            ref = row.get("Hakediş No") or row.get("_id")
            payable = _n(row.get("Ödenecek Tutar"), _n(row.get("Net Hakediş")))
            paid = _n(row.get("Ödenen"))
            if paid > payable + 0.01 and payable >= 0:
                add("KRİTİK", "Taşeron Hakedişi", ref, f"Ödenen tutar ({paid:.2f}) ödenecek tutarı ({payable:.2f}) aşıyor.",
                    "Banka dekontu, mahsup ve hakediş revizyonlarını kontrol edin; muhasebe düzeltmesi olmadan kaydı değiştirmeyin.")
            if not _s(row.get("WBS")):
                add("ORTA", "Taşeron Hakedişi", ref, "Hakediş WBS bağlantısı taşımıyor.", "Hakediş kapsamını onaylı imalat WBS kodlarına bağlayın.")
            if not _s(row.get("Dönem")):
                add("ORTA", "Taşeron Hakedişi", ref, "Hakediş dönemi eksik.", "Dönem başlangıç/bitişini ve ekli metraj ataşmanını kaydedin.")

    # Employer receivables: overdue outstanding balances and over-collection.
    employer_claims = _frame(data, "employer_hakediş")
    for _, row in employer_claims.iterrows():
        ref = row.get("Hakediş No") or row.get("_id")
        balance = _n(row.get("Kalan Alacak"))
        over = _n(row.get("Fazla Tahsilat"))
        due = _d(row.get("Vade Tarihi"))
        if over > 0.01:
            add("KRİTİK", "İşveren Hakedişi", ref, f"Fazla tahsilat {over:.2f} olarak hesaplanmış.",
                "Banka hareketi, mahsup ve hakediş toplamını mutabık kılın.")
        if balance > 0.01 and pd.notna(due) and due.date() < report_date:
            add("YÜKSEK", "İşveren Hakedişi", ref, f"Vadesi geçmiş tahsil edilmemiş alacak {balance:.2f}.",
                "İşveren ödeme takvimini takip edin; gecikme nedenini ve tahsilat aksiyonunu kaydedin.")
        if balance > 0.01 and pd.isna(due):
            add("ORTA", "İşveren Hakedişi", ref, "Açık alacağın vade tarihi eksik.", "Sözleşme ödeme vadesine göre tahsilat hedef tarihi belirleyin.")

    # Approval queue aging and routing completeness.
    approvals = _frame(data, "approvals")
    for _, row in approvals.iterrows():
        if _s(row.get("Durum")).casefold() not in {"bekliyor", "onay bekliyor", "pending"}:
            continue
        ref = row.get("_id") or row.get("Kayıt ID")
        created = _d(row.get("Talep Tarihi"))
        if pd.notna(created) and (report_date - created.date()).days > 5:
            age = (report_date - created.date()).days
            add("YÜKSEK", "Onay Merkezi", ref, f"Onay talebi {age} gündür bekliyor.",
                "Onay rolü sahibine eskalasyon yapın; karar gerekçesini ve tarihini kaydedin.")
        if not _s(row.get("Onay Rolü")):
            add("ORTA", "Onay Merkezi", ref, "Bekleyen onay talebine rol atanmamış.", "Talebi doğru yetki rolüne yönlendirin.")

    # Cash commitments/payment plan: no automatic cashflow netting across currencies.
    cash = _frame(data, "cash")
    if not cash.empty:
        docs = cash[cash.get("Belge No", pd.Series(index=cash.index, dtype=object)).fillna("").astype(str).str.strip().ne("")]
        duplicates = docs[docs["Belge No"].astype(str).str.strip().duplicated(keep=False)] if not docs.empty else pd.DataFrame()
        for doc in sorted(duplicates["Belge No"].astype(str).unique()) if not duplicates.empty else []:
            add("YÜKSEK", "Nakit / Ödeme", doc, "Nakit hareketlerinde belge numarası tekrarlanıyor.", "Banka dekontu ve cari hesap ekstresiyle mükerrer kayıt olup olmadığını doğrulayın.")
        for _, row in cash.iterrows():
            status = _s(row.get("Durum")).casefold()
            if status in {"odendi", "ödendi", "tahsil edildi", "kapali", "closed", "iptal"}:
                continue
            due = _d(row.get("Vade Tarihi"))
            amount = _n(row.get("Tutar"))
            if amount > 0 and pd.notna(due) and due.date() < report_date:
                direction = _s(row.get("Tür")) or "Nakit"
                add("YÜKSEK", "Nakit / Ödeme", row.get("Belge No") or row.get("_id"),
                    f"Vadesi geçmiş açık {direction.lower()} kaydı ({amount:.2f}).",
                    "Ödeme/tahsilat durumunu banka ve cari kayıtlarıyla doğrulayın; döviz bazını koruyun.")
            if amount > 0 and not _s(row.get("Para Birimi")):
                add("ORTA", "Nakit / Ödeme", row.get("Belge No") or row.get("_id"), "Nakit kaydında para birimi eksik.",
                    "Tutarın para birimini girin; farklı para birimlerini kur olmadan toplamaya dahil etmeyin.")

    # Daily reports are checked for completeness only when a report exists; missing days are not assumed.
    daily = _frame(data, "daily_reports")
    for _, row in daily.iterrows():
        ref = row.get("Rapor Tarihi") or row.get("_id")
        if not _s(row.get("Rapor Tarihi")):
            add("ORTA", "Günlük Şantiye Raporu", ref, "Şantiye günlük rapor tarihi eksik.", "Gerçek rapor tarihini ve raporlayan kişiyi kaydedin.")
        if not _s(row.get("Raporlayan")):
            add("ORTA", "Günlük Şantiye Raporu", ref, "Günlük raporda raporlayan kişi eksik.", "Rapor sorumlusunu ve doğrulama bilgisini ekleyin.")
        if not _s(row.get("Günlük İmalat Özeti")):
            add("DÜŞÜK", "Günlük Şantiye Raporu", ref, "Günlük imalat özeti boş.", "Yapılan imalatı, mahal/katı, miktarı ve engelleri yazın.")

    # Labour amount arithmetic sanity check; flags only material discrepancies.
    labor = _frame(data, "labor")
    for _, row in labor.iterrows():
        ref = row.get("_id") or row.get("Tarih")
        expected = _n(row.get("Çalışan Sayısı")) * _n(row.get("Gün")) * _n(row.get("Kişi Günlük Ücret"))
        actual = _n(row.get("Toplam İşçilik"))
        if expected > 0 and actual > 0 and abs(expected - actual) > max(1.0, expected * 0.01):
            add("YÜKSEK", "İşçilik / Puantaj", ref,
                f"Puantaj matematik kontrolü uyuşmuyor: beklenen {expected:.2f}, kayıtlı {actual:.2f}.",
                "Kişi sayısı, gün, günlük ücret, fazla mesai ve toplamı bordro/puantaj ekiyle karşılaştırın.")
        if _n(row.get("Çalışan Sayısı")) < 0 or _n(row.get("Gün")) < 0 or _n(row.get("Kişi Günlük Ücret")) < 0:
            add("KRİTİK", "İşçilik / Puantaj", ref, "Puantaj satırında negatif kişi/gün/ücret girdisi var.", "Kaynak puantaj belgesine göre giriş hatasını düzeltin.")

    return findings
