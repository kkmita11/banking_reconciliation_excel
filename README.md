# banking_reconciliation_excel
🇵🇱 Automatyzacja Uzgodnień Finansowych (Excel + VBA + Power Query)
### Opis Projektu

Zaawansowane narzędzie do automatyzacji procesów uzgadniania transakcji bankowych z zapisami w systemie ERP. Projekt eliminuje manualną pracę przy parowaniu setek operacji finansowych, minimalizując błędy ludzkie.
### Kluczowe Funkcje

    Automatyczne Pobieranie Danych: Integracja plików CSV z ewidencji i wyciągów bankowych za pomocą Power Query.

    Inteligentne Uzgadnianie: Automatyczne łączenie transakcji na podstawie unikalnych identyfikatorów (np. ERP/2026/001).

    Manual Matching UI: Dedykowany moduł VBA do ręcznego parowania trudnych przypadków (np. prowizji lub wpłat zbiorczych).

    Bezpieczeństwo Danych: System ochrony arkuszy oparty na haśle „admin”.

    Interaktywny Dashboard: Wizualizacja wyników uzgodnień przy użyciu fragmentatorów i osi czasu.

### Struktura Repozytorium

    /Data: Przykładowe pliki źródłowe CSV.

    projekt.xlsm: Główny plik aplikacji.

    /Source: VBA i PowerQuery kod.

### Instrukcja Uruchomienia

    Pobierz plik projekt.xlsm i pliki CSV.

    Upewnij się, że w opcjach Power Query odznaczone jest „Odświeżanie w tle”.

    Zaktualizuj ścieżki do plików źródłowych w zapytaniach.

    Domyślne hasło do arkuszy to: admin.

🇬🇧 Financial Reconciliation Automation (Excel + VBA + Power Query)
### Project Overview

A professional tool designed to automate the reconciliation process between ERP records and bank statements. The project eliminates manual labor in pairing financial operations, significantly reducing the risk of human error.
### Key Features

    Automated Data Sourcing: Seamless CSV integration for ERP records and bank statements via Power Query.

    Smart Matching: Automatic pairing of transactions based on unique internal IDs (e.g., ERP/2026/001).

    Manual Matching UI: Custom VBA module for manual pairing of complex cases like bank fees or partial payments.

    Data Security: Sheet protection system utilizing "admin" password.

    Interactive Dashboard: Visual results overview using Slicers and Timelines for quick analysis.

### Repository Structure

    /Data: Sample CSV source files.

    projekt.xlsm: Main application file.
    
    /Source: VBA and PowerQuery code.

### Getting Started

    Download projekt.xlsm and the corresponding CSV files.

    Ensure that "Background Refresh" is disabled in the Power Query connection properties.

    Update the source file paths in the queries.

    Default sheet password: admin.
    
