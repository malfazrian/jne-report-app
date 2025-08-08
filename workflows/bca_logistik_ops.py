import os
from workflows.email_ops import find_latest_matching_email, save_attachments_danamon, extract_date_from_subject


def process_bca_logistik_report(tracker, task_paths):
    """
    Proses laporan BCA Logistik.
    """
    print("=== Mulai Proses Laporan BCA Logistik ===")
    
    print("Mencari email dengan subjek 'OS CARD JNE'...")
    email_os = find_latest_matching_email(
        file_path="D:/email/Email TB/outlook.office365.com/Inbox.sbd/DANAMON",
        subject_prefix="OS CARD JNE",
        max_emails=10)
    if email_os:
        save_attachments_danamon(email_os, save_dir="d:/RYAN/1. References/Danamon")
        print("Email OS ditemukan dan attachment disimpan.")