# TEST.md - cli-anything-puantaj

## Test plani

Bu harness, calisan bir uretim veritabanina (gercek personel/puantaj verisi)
baglandigi icin testler **veriyi degistirmez**. Kapsam:

- `test_core.py`
  - `calc.calc_day_hours` saf hesaplama testleri (DB gerektirmez):
    - hafta ici fazla mesai
    - mola dusumu
    - gece saatleri
    - ozel gun
  - CLI subprocess: `--help`, `info --json`, `employee list --json`
    (salt-okunur, veri degistirmez)

Veri ekleme/silme (employee add, timesheet add, delete) komutlari kasitli
olarak otomatik test edilmez; uretim DB'sini kirletmemek icin manuel/izole
DB ile dogrulanmalidir.

## Test sonuclari

```
platform win32 -- Python 3.14.3, pytest-8.4.1
collected 8 items

TestCalc::test_weekday_overtime PASSED
TestCalc::test_no_overtime PASSED
TestCalc::test_break_deduction PASSED
TestCalc::test_night_hours PASSED
TestCalc::test_special_day PASSED
TestCLISubprocess::test_help PASSED
TestCLISubprocess::test_info_json PASSED
TestCLISubprocess::test_employee_list_json PASSED

8 passed in 1.74s
```

Ozet: 8/8 gecti (%100). CLI_ANYTHING_FORCE_INSTALLED=1 ile kurulu komut
uzerinden dogrulandi. Yazma komutlari (add/update/delete) uretim DB'sini
korumak icin otomatik test disinda birakildi.
