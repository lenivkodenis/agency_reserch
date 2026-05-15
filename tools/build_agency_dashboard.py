from pathlib import Path
import json
import math
import re

import pandas as pd


ROOT = Path(__file__).resolve().parent
SOURCE = Path("/Users/denislenivko/Downloads/DataExport_2026-04-27_19-54-22.xlsx")
PRESENTATION = Path("/Users/denislenivko/Downloads/agency_classifieds_research_presentation.html")
OUTPUT = ROOT / "agency_revenue_dashboard_q1_2026.html"


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).replace("\xa0", " ").strip()


def safe_float(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return 0.0
    return float(value)


def pct(part, whole):
    return 0 if not whole else float(part) / float(whole)


def top_name(series):
    if series.empty:
        return ""
    return str(series.idxmax())


def monthly_columns(df, label):
    columns = [c for c in df.columns if label in str(c)]

    def key(col):
        match = re.search(r"(\d{2})\.(\d{2})\.(\d{4})", str(col))
        if not match:
            return str(col)
        day, month, year = match.groups()
        return f"{year}-{month}-{day}"

    return sorted(columns, key=key)


def records_from_grouped(grouped):
    return grouped.reset_index().to_dict(orient="records")


df = pd.read_excel(SOURCE, sheet_name="Chart data")

text_cols = ["Регион", "Агентство", "ГК", "Дилер"]
for col in text_cols:
    df[col] = df[col].map(clean_text)

rev_cols = monthly_columns(df, "Выручка")
svc_cols = monthly_columns(df, "Кол-во платных услуг")
months = ["Январь", "Февраль", "Март"]

for col in rev_cols + svc_cols:
    df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

df["revenue_q1"] = df[rev_cols].sum(axis=1)
df["services_q1"] = df[svc_cols].sum(axis=1)

total_revenue = float(df["revenue_q1"].sum())
total_services = float(df["services_q1"].sum())
managed_mask = df["Агентство"].astype(str).str.strip().ne("")
managed_revenue = float(df.loc[managed_mask, "revenue_q1"].sum())
unassigned_revenue = total_revenue - managed_revenue

agency_group = df.groupby("Агентство", dropna=False)
agency_summary = agency_group.agg(
    revenue=("revenue_q1", "sum"),
    services=("services_q1", "sum"),
    dealers=("Дилер", "nunique"),
    groups=("ГК", "nunique"),
    regions=("Регион", "nunique"),
).reset_index()
for idx, col in enumerate(rev_cols):
    agency_summary[f"revenue_m{idx}"] = agency_group[col].sum().values
for idx, col in enumerate(svc_cols):
    agency_summary[f"services_m{idx}"] = agency_group[col].sum().values
agency_summary["revenueShare"] = agency_summary["revenue"].map(lambda x: pct(x, total_revenue))
agency_summary["servicesShare"] = agency_summary["services"].map(lambda x: pct(x, total_services))
agency_summary["revenuePerDealer"] = agency_summary.apply(lambda r: pct(r["revenue"], r["dealers"]), axis=1)
agency_summary["revenuePerService"] = agency_summary.apply(lambda r: pct(r["revenue"], r["services"]), axis=1)
agency_summary = agency_summary.sort_values("revenue", ascending=False)

agencies = []
for _, row in agency_summary.iterrows():
    agency = row["Агентство"]
    sub = df[df["Агентство"] == agency]
    top_groups = (
        sub.groupby("ГК")
        .agg(revenue=("revenue_q1", "sum"), services=("services_q1", "sum"), dealers=("Дилер", "nunique"))
        .sort_values("revenue", ascending=False)
        .head(5)
        .reset_index()
    )
    top_regions = (
        sub.groupby("Регион")
        .agg(revenue=("revenue_q1", "sum"), services=("services_q1", "sum"), dealers=("Дилер", "nunique"))
        .sort_values("revenue", ascending=False)
        .head(6)
        .reset_index()
    )
    agencies.append(
        {
            "name": agency,
            "revenue": safe_float(row["revenue"]),
            "services": safe_float(row["services"]),
            "dealers": int(row["dealers"]),
            "groups": int(row["groups"]),
            "regions": int(row["regions"]),
            "revenueShare": safe_float(row["revenueShare"]),
            "servicesShare": safe_float(row["servicesShare"]),
            "revenuePerDealer": safe_float(row["revenuePerDealer"]),
            "revenuePerService": safe_float(row["revenuePerService"]),
            "revenueByMonth": [safe_float(row[f"revenue_m{i}"]) for i in range(len(rev_cols))],
            "servicesByMonth": [safe_float(row[f"services_m{i}"]) for i in range(len(svc_cols))],
            "topGroups": [
                {
                    "name": r["ГК"],
                    "revenue": safe_float(r["revenue"]),
                    "services": safe_float(r["services"]),
                    "dealers": int(r["dealers"]),
                    "shareInsideAgency": pct(r["revenue"], row["revenue"]),
                }
                for _, r in top_groups.iterrows()
            ],
            "topRegions": [
                {
                    "name": r["Регион"],
                    "revenue": safe_float(r["revenue"]),
                    "services": safe_float(r["services"]),
                    "dealers": int(r["dealers"]),
                    "shareInsideAgency": pct(r["revenue"], row["revenue"]),
                }
                for _, r in top_regions.iterrows()
            ],
        }
    )

region_rows = []
for region, sub in df.groupby("Регион"):
    agency_rev = sub.groupby("Агентство")["revenue_q1"].sum().sort_values(ascending=False)
    agency_svc = sub.groupby("Агентство")["services_q1"].sum().sort_values(ascending=False)
    revenue = float(sub["revenue_q1"].sum())
    services = float(sub["services_q1"].sum())
    region_rows.append(
        {
            "name": region,
            "revenue": revenue,
            "services": services,
            "dealers": int(sub["Дилер"].nunique()),
            "groups": int(sub["ГК"].nunique()),
            "agencies": int(sub["Агентство"].nunique()),
            "agencyNames": sorted(sub["Агентство"].unique().tolist()),
            "dominantAgency": top_name(agency_rev),
            "dominantAgencyShare": pct(float(agency_rev.iloc[0]) if len(agency_rev) else 0, revenue),
            "revenueByAgency": {str(k): safe_float(v) for k, v in agency_rev.items()},
            "servicesByAgency": {str(k): safe_float(v) for k, v in agency_svc.items()},
        }
    )
regions = sorted(region_rows, key=lambda x: x["revenue"], reverse=True)

group_rows = []
for group, sub in df.groupby("ГК"):
    agency_rev = sub.groupby("Агентство")["revenue_q1"].sum().sort_values(ascending=False)
    revenue = float(sub["revenue_q1"].sum())
    services = float(sub["services_q1"].sum())
    group_rows.append(
        {
            "name": group,
            "revenue": revenue,
            "services": services,
            "dealers": int(sub["Дилер"].nunique()),
            "regions": int(sub["Регион"].nunique()),
            "agencies": int(sub["Агентство"].nunique()),
            "topAgency": top_name(agency_rev),
            "topAgencyShare": pct(float(agency_rev.iloc[0]) if len(agency_rev) else 0, revenue),
            "agencySplit": {str(k): safe_float(v) for k, v in agency_rev.items()},
        }
    )
groups = sorted(group_rows, key=lambda x: x["revenue"], reverse=True)

dealers = []
for _, row in df.sort_values("revenue_q1", ascending=False).iterrows():
    dealers.append(
        {
            "region": row["Регион"],
            "agency": row["Агентство"],
            "group": row["ГК"],
            "dealer": row["Дилер"],
            "revenue": safe_float(row["revenue_q1"]),
            "services": safe_float(row["services_q1"]),
            "revenueByMonth": [safe_float(row[col]) for col in rev_cols],
            "servicesByMonth": [safe_float(row[col]) for col in svc_cols],
        }
    )

month_revenue = [safe_float(df[col].sum()) for col in rev_cols]
month_services = [safe_float(df[col].sum()) for col in svc_cols]
capitals_revenue = sum(r["revenue"] for r in regions if r["name"] in {"Москва и Московская область", "Санкт-Петербург и Ленинградская область"})
capitals_services = sum(r["services"] for r in regions if r["name"] in {"Москва и Московская область", "Санкт-Петербург и Ленинградская область"})
agency_presence_counts = {}
for r in regions:
    agency_presence_counts[str(r["agencies"])] = agency_presence_counts.get(str(r["agencies"]), 0) + 1

least_regions = sorted(regions, key=lambda x: x["revenue"])[:5]
multi_agency_regions = [r for r in regions if r["agencies"] > 1]

data = {
    "meta": {
        "source": "DataLens",
        "sourceFile": SOURCE.name,
        "exportedAt": "27.04.2026 19:54:22",
        "period": "Q1 2026",
        "months": months,
        "totalRevenue": total_revenue,
        "totalServices": total_services,
        "managedRevenue": managed_revenue,
        "unassignedRevenue": unassigned_revenue,
        "totalDealers": int(df["Дилер"].nunique()),
        "totalGroups": int(df["ГК"].nunique()),
        "totalRegions": int(df["Регион"].nunique()),
        "totalAgencies": int(df["Агентство"].nunique()),
        "monthRevenue": month_revenue,
        "monthServices": month_services,
        "capitalsRevenueShare": pct(capitals_revenue, total_revenue),
        "capitalsServicesShare": pct(capitals_services, total_services),
        "agencyPresenceCounts": agency_presence_counts,
        "leastRegions": least_regions,
        "multiAgencyRegions": multi_agency_regions,
    },
    "agencies": agencies,
    "regions": regions,
    "groups": groups,
    "dealers": dealers,
}

autoru_potential = {
    "title": "Расчет потенциала агентства Авто.ру",
    "subtitle": "Первичная SWOT-гипотеза: есть ли смысл запускать собственное агентство сопровождения кабинетов и продвижения дилеров.",
    "assumptions": [
        "Расчёт ресурса: 1 сотрудник операционного сопровождения = 20 салонов.",
        "Стартовая команда: 4-6 человек на пилот, включая руководителя/аккаунта, 2-4 сотрудников агентства и shared-ресурс аналитика/техинтегратора.",
        "Команда Auto.ru имеет доступ к Авто.ру Бизнес и может управлять размещением и аналитикой в более расширенном формате, чем внешние агентства.",
        "Базовая цена сопровождения: 25-35 тыс. ₽ за салон; расширенный контур с продвижением и контентом может быть выше.",
        "Пилотный перехват у текущих агентств: 5-10% подключенных салонов из выгрузки, если предложить прозрачную связку данных Auto.ru, контроля качества и SLA.",
        "Новый региональный потенциал: 40-80 салонов в городах-миллионниках на первом этапе при фокусе на регионы с низкой представленностью агентств.",
    ],
    "market": {
        "currentSalons": int(sum(agency["dealers"] for agency in agencies)),
        "captureLow": int(round(sum(agency["dealers"] for agency in agencies) * 0.05)),
        "captureHigh": int(round(sum(agency["dealers"] for agency in agencies) * 0.10)),
        "newRegionalLow": 40,
        "newRegionalHigh": 80,
        "priceLow": 25000,
        "priceHigh": 35000,
    },
    "swot": [
        {
            "type": "S",
            "title": "Сильные стороны",
            "items": [
                "Прямой доступ к продуктовой экспертизе Auto.ru и данным по качеству размещений.",
                "Доступ к Авто.ру Бизнес позволяет управлять и анализировать кабинеты глубже, чем это обычно доступно конкурентам.",
                "Можно упаковать сопровождение как прозрачный SLA: выгрузка, ошибки, блокировки, рекомендации, эффективность.",
                "Высокое доверие дилеров к площадке снижает барьер первого пилота.",
            ],
        },
        {
            "type": "W",
            "title": "Слабые стороны",
            "items": [
                "Нужна операционная команда, а не только продукт: примерно 1 сотрудник на 20 салонов плюс shared-аналитика и техподдержка.",
                "Риск конфликта восприятия: площадка одновременно продаёт размещение и услугу сопровождения.",
                "Без сильного сервиса агентство быстро станет очередным постинг-подрядчиком.",
            ],
        },
        {
            "type": "O",
            "title": "Возможности",
            "items": [
                "Перехватить 5-10% текущих салонов за счёт прозрачной экономики и качества данных.",
                "Добрать 40-80 новых салонов в городах-миллионниках, где сейчас меньше агентского покрытия.",
                "Вести за клиента бортжурнал в Авто.ру как дополнительный контентный и удерживающий сервис.",
                "Собрать стандарт рынка: единая аналитика, отчётность, контроль ошибок и понятная модель оплаты.",
            ],
        },
        {
            "type": "T",
            "title": "Риски",
            "items": [
                "Агентства могут быстро снизить цену или усилить персональный сервис.",
                "Дилеры могут не захотеть менять подрядчика без доказанного прироста лидов.",
                "Для полноценного агентского продукта нужен свой фотобанк и регулярное создание контента для дилеров.",
                "Операционные ошибки будут бить по бренду Auto.ru сильнее, чем по независимому агентству.",
            ],
        },
    ],
}
autoru_potential["market"]["totalPotentialLow"] = autoru_potential["market"]["captureLow"] + autoru_potential["market"]["newRegionalLow"]
autoru_potential["market"]["totalPotentialHigh"] = autoru_potential["market"]["captureHigh"] + autoru_potential["market"]["newRegionalHigh"]
autoru_potential["market"]["monthlyRevenueLow"] = autoru_potential["market"]["totalPotentialLow"] * autoru_potential["market"]["priceLow"]
autoru_potential["market"]["monthlyRevenueHigh"] = autoru_potential["market"]["totalPotentialHigh"] * autoru_potential["market"]["priceHigh"]
autoru_potential["market"]["staffRatio"] = 20
autoru_potential["market"]["staffLow"] = math.ceil(autoru_potential["market"]["totalPotentialLow"] / autoru_potential["market"]["staffRatio"])
autoru_potential["market"]["staffHigh"] = math.ceil(autoru_potential["market"]["totalPotentialHigh"] / autoru_potential["market"]["staffRatio"])

employee_net_salary = 120000
tax_rate = 0.13
contribution_rate = 0.40
employee_full_cost = employee_net_salary / (1 - tax_rate) * (1 + contribution_rate)
manager_net_salary = employee_net_salary * 2
manager_full_cost = manager_net_salary / (1 - tax_rate) * (1 + contribution_rate)
team_low = 3
team_high = 6
team_cost_low = team_low * employee_full_cost + manager_full_cost
team_cost_high = team_high * employee_full_cost + manager_full_cost
price_low = autoru_potential["market"]["priceLow"]
price_high = autoru_potential["market"]["priceHigh"]
medium_buffer = 1.25
autoru_potential["economics"] = {
    "salaryNet": employee_net_salary,
    "taxRate": tax_rate,
    "contributionRate": contribution_rate,
    "employeeFullCost": float(employee_full_cost),
    "managerNetSalary": manager_net_salary,
    "managerFullCost": float(manager_full_cost),
    "teamLow": team_low,
    "teamHigh": team_high,
    "teamCostLow": float(team_cost_low),
    "teamCostHigh": float(team_cost_high),
    "targetRevenuePerEmployee": 1000000,
    "distribution": "Продажи агентства возможны совместно с постингом либо как отдельный продуктовый фокус.",
    "unitScenarios": [
        {
            "name": "Текущая норма загрузки",
            "salonsPerEmployee": autoru_potential["market"]["staffRatio"],
            "priceLow": price_low,
            "priceHigh": price_high,
            "revenueLow": float(autoru_potential["market"]["staffRatio"] * price_low),
            "revenueHigh": float(autoru_potential["market"]["staffRatio"] * price_high),
            "requirement": "20 салонов на сотрудника дают 0,5-0,7 млн ₽/мес. Для 1 млн нужна либо более высокая загрузка, либо чек выше базового.",
        },
        {
            "name": "Цель через загрузку",
            "salonsPerEmployeeLow": int(math.ceil(1000000 / price_high)),
            "salonsPerEmployeeHigh": int(math.ceil(1000000 / price_low)),
            "priceLow": price_low,
            "priceHigh": price_high,
            "revenueLow": float(math.ceil(1000000 / price_high) * price_high),
            "revenueHigh": float(math.ceil(1000000 / price_low) * price_low),
            "requirement": "Нужно 29-40 салонов на сотрудника при цене 25-35 тыс. ₽. Это возможно только при сильной автоматизации рутины и стандартизированном SLA.",
        },
        {
            "name": "Цель через чек",
            "salonsPerEmployee": autoru_potential["market"]["staffRatio"],
            "priceLow": 50000,
            "priceHigh": 50000,
            "revenueLow": 1000000,
            "revenueHigh": 1000000,
            "requirement": "При 20 салонах на сотрудника нужен средний чек 50 тыс. ₽/мес: сопровождение + контент + аналитика Авто.ру Бизнес + контроль качества.",
        },
        {
            "name": "Сбалансированный сценарий",
            "salonsPerEmployeeLow": 25,
            "salonsPerEmployeeHigh": 30,
            "priceLow": 34000,
            "priceHigh": 40000,
            "revenueLow": 850000,
            "revenueHigh": 1200000,
            "requirement": "25-30 салонов на сотрудника агентства и чек 34-40 тыс. ₽ дают выход к 1 млн+ без экстремальной нагрузки.",
        },
    ],
    "scenarios": [
        {
            "name": "Старт: окупаемость персонала",
            "salonsLow": int(math.ceil(team_cost_low / price_high)),
            "salonsHigh": int(math.ceil(team_cost_high / price_low)),
            "revenueLow": float(math.ceil(team_cost_low / price_high) * price_high),
            "revenueHigh": float(math.ceil(team_cost_high / price_low) * price_low),
            "logic": "Минимально закрыть ФОТ команды и руководителя без учёта прочих накладных расходов.",
        },
        {
            "name": "Среднее плечо: продажи совместно с постингом",
            "salonsLow": int(math.ceil(team_cost_low * medium_buffer / price_low)),
            "salonsHigh": int(math.ceil(team_cost_high * medium_buffer / price_low)),
            "revenueLow": float(math.ceil(team_cost_low * medium_buffer / price_low) * price_low),
            "revenueHigh": float(math.ceil(team_cost_high * medium_buffer / price_low) * price_low),
            "logic": "Продажи агентства идут совместно с постингом либо как отдельный продуктовый фокус; закладываем 25% буфер на churn, простои и операционные накладные расходы.",
        },
        {
            "name": "Длинное плечо: хорошие объёмы",
            "salonsLow": int(team_low * autoru_potential["market"]["staffRatio"]),
            "salonsHigh": int(team_high * autoru_potential["market"]["staffRatio"]),
            "revenueLow": float(team_low * autoru_potential["market"]["staffRatio"] * price_low),
            "revenueHigh": float(team_high * autoru_potential["market"]["staffRatio"] * price_high),
            "logic": "Выход на нормальную загрузку по правилу 1 сотрудник агентства = 20 салонов.",
        },
    ],
}

account_manager_salary = 200000
sales_bonus_rate = 0.05
sales_bonus_per_connection = account_manager_salary * sales_bonus_rate
sales_plan_months = [
    {"month": 1, "conservative": 5, "target": 10},
    {"month": 2, "conservative": 7, "target": 15},
    {"month": 3, "conservative": 10, "target": 20},
    {"month": 4, "conservative": 12, "target": 22},
    {"month": 5, "conservative": 13, "target": 25},
    {"month": 6, "conservative": 13, "target": 28},
]
conservative_cumulative = 0
target_cumulative = 0
for row in sales_plan_months:
    conservative_cumulative += row["conservative"]
    target_cumulative += row["target"]
    row["conservativeCumulative"] = conservative_cumulative
    row["targetCumulative"] = target_cumulative
    row["conservativeBonus"] = row["conservative"] * sales_bonus_per_connection
    row["targetBonus"] = row["target"] * sales_bonus_per_connection
    row["conservativeRevenue"] = conservative_cumulative * price_low
    row["targetRevenue"] = target_cumulative * price_high

autoru_potential["economics"]["salesCompensation"] = {
    "accountManagerSalary": account_manager_salary,
    "bonusRate": sales_bonus_rate,
    "bonusPerConnection": float(sales_bonus_per_connection),
    "plan": sales_plan_months,
    "conservativeTotal": conservative_cumulative,
    "targetTotal": target_cumulative,
    "conservativeBonusTotal": float(sum(row["conservativeBonus"] for row in sales_plan_months)),
    "targetBonusTotal": float(sum(row["targetBonus"] for row in sales_plan_months)),
}

agency_by_name = {agency["name"]: agency for agency in agencies}


def market_row(name, source_name, price_label, low_monthly, high_monthly=None, note="", source=""):
    high_monthly = low_monthly if high_monthly is None else high_monthly
    agency = agency_by_name[source_name]
    return {
        "name": name,
        "sourceName": source_name,
        "salons": int(agency["dealers"]),
        "groups": int(agency["groups"]),
        "priceLabel": price_label,
        "monthlyLow": float(low_monthly),
        "monthlyHigh": float(high_monthly),
        "annualLow": float(low_monthly * 12),
        "annualHigh": float(high_monthly * 12),
        "note": note,
        "source": source,
    }


four_pixels_salons = agency_by_name["4 пикселя +"]["dealers"]
karkopi_salons = agency_by_name["Каркопи"]["dealers"]
tandem_salons = agency_by_name["ТАК"]["dealers"]
karkopi_package_count = math.ceil(karkopi_salons / 15)

market_analysis = {
    "definition": "Клиент в расчёте = салон/дилерская точка из выгрузки. ГК показаны отдельно, потому что одна ГК может включать много салонов.",
    "rows": [
        market_row(
            "4 Пикселя",
            "4 пикселя +",
            "от 28 000 ₽/мес за классифайды",
            four_pixels_salons * 28000,
            note="Минимальный публичный тариф; расширенный scope и индивидуальные условия не включены.",
            source="https://4px.ru/services/classifieds/",
        ),
        market_row(
            "Каркопи / CarCopy",
            "Каркопи",
            "прокси: 200 000 ₽/мес за 15 аккаунтов",
            karkopi_package_count * 200000,
            note=f"Открытый тариф CarCopy указан для CRM-аккаунтов, не для салонов. Для сопоставимости принято {karkopi_package_count} пакетов по 15 аккаунтов.",
            source="https://carcopy.ru/crm",
        ),
        market_row(
            "ТАК / Tandem",
            "ТАК",
            "8–10 тыс. ₽ без продвижения; 50 тыс. ₽ с продвижением + постинг",
            tandem_salons * 8000,
            tandem_salons * 50000,
            note="Нижняя граница — сопровождение без продвижения; верхняя — автоматизированное продвижение + постинг по вводным пользователя.",
            source="https://tandemdigital.ru/uslugi/classifieds-reklama/",
        ),
    ],
    "sources": [
        {"label": "4 Пикселя: классифайды от 28 000 ₽", "url": "https://4px.ru/services/classifieds/"},
        {"label": "CarCopy CRM: 15 аккаунтов = 200 000 ₽/мес", "url": "https://carcopy.ru/crm"},
        {"label": "Tandem: Autocloud / автоматизированное продвижение", "url": "https://tandemdigital.ru/uslugi/classifieds-reklama/"},
        {"label": "Tandem: тарифные вводные пользователя", "url": ""},
    ],
}
market_analysis["summary"] = {
    "salons": int(sum(row["salons"] for row in market_analysis["rows"])),
    "monthlyLow": float(sum(row["monthlyLow"] for row in market_analysis["rows"])),
    "monthlyHigh": float(sum(row["monthlyHigh"] for row in market_analysis["rows"])),
    "annualLow": float(sum(row["annualLow"] for row in market_analysis["rows"])),
    "annualHigh": float(sum(row["annualHigh"] for row in market_analysis["rows"])),
}
data["marketAnalysis"] = market_analysis
data["autoruPotential"] = autoru_potential

html_template = r"""<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Работа агентств с клиентами Авто.ру</title>
  <style>
    :root {
      --paper: #d7ddd9;
      --sheet: #fdfdfb;
      --ink: #171716;
      --muted: #656560;
      --line: #cfd6d3;
      --line-dark: #272724;
      --red: #dd3f2d;
      --petrol: #006f76;
      --amber: #c88b13;
      --blue: #395f9f;
      --green: #497a45;
      --soft-red: rgba(221, 63, 45, .11);
      --soft-petrol: rgba(0, 111, 118, .11);
      --soft-amber: rgba(200, 139, 19, .14);
      --shadow: 0 18px 36px rgba(23, 23, 22, .09);
      color-scheme: light;
    }

    * {
      box-sizing: border-box;
    }

    html {
      scroll-behavior: smooth;
    }

    body {
      margin: 0;
      background:
        linear-gradient(90deg, rgba(23, 23, 22, .075) 1px, transparent 1px),
        linear-gradient(180deg, rgba(23, 23, 22, .06) 1px, transparent 1px),
        var(--paper);
      background-size: 72px 72px, 72px 72px, auto;
      color: var(--ink);
      font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
      letter-spacing: 0;
    }

    a {
      color: inherit;
    }

    .page {
      width: min(1480px, calc(100% - 32px));
      margin: 0 auto;
      padding: 28px 0 56px;
    }

    .study-header {
      margin: 4px 0 14px;
      text-align: center;
    }

    .study-title {
      max-width: none;
      margin: 0;
      font-size: clamp(2.3rem, 4.8vw, 4.7rem);
      line-height: .98;
    }

    .study-switch {
      position: relative;
      display: flex;
      justify-content: center;
      gap: 8px;
      width: fit-content;
      max-width: 100%;
      margin: 0 auto 22px;
      padding: 6px;
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      background: rgba(253, 253, 251, .88);
      box-shadow: 0 14px 30px rgba(23, 23, 22, .12);
      backdrop-filter: blur(12px);
    }

    .study-tab {
      min-height: 38px;
      padding: 8px 14px;
      border-color: transparent;
      background: transparent;
      color: var(--muted);
      font-size: .88rem;
      font-weight: 800;
      cursor: pointer;
      transition: background .16s ease, color .16s ease, transform .16s ease;
    }

    .study-tab:hover {
      color: var(--ink);
      transform: translateY(-1px);
    }

    .study-tab.is-active {
      background: #171716;
      color: #fdfdfb;
      border-color: #171716;
    }

    .tab-panel[hidden] {
      display: none;
    }

    .research-shell {
      display: grid;
      gap: 14px;
    }

    .research-intro {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 16px;
      align-items: end;
      padding: 22px;
      background: #171716;
      color: #fdfdfb;
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      box-shadow: var(--shadow);
    }

    .research-intro h2 {
      margin: 0;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 2.05rem;
      line-height: 1.08;
    }

    .research-intro p {
      max-width: 780px;
      margin: 8px 0 0;
      color: rgba(253, 253, 251, .72);
      line-height: 1.5;
    }

    .research-intro span {
      color: rgba(253, 253, 251, .58);
      font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
      font-size: .78rem;
      white-space: nowrap;
    }

    .research-frame {
      width: 100%;
      min-height: 760px;
      height: calc(100vh - 145px);
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      background: #0b1020;
      box-shadow: var(--shadow);
    }

    .tab-panel[hidden].market-panel,
    .tab-panel[hidden].research-shell {
      display: none;
    }

    .market-panel {
      display: grid;
      gap: 14px;
    }

    .market-hero {
      display: grid;
      grid-template-columns: minmax(0, 1fr) minmax(280px, .42fr);
      gap: 14px;
      align-items: stretch;
    }

    .market-hero-main,
    .market-summary-card {
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      background: var(--sheet);
      box-shadow: var(--shadow);
      padding: 22px;
    }

    .market-hero-main h2 {
      margin: 0;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 2.35rem;
      line-height: 1.05;
    }

    .market-hero-main p {
      max-width: 820px;
      margin: 10px 0 0;
      color: var(--muted);
      line-height: 1.5;
    }

    .market-summary-card {
      background: #171716;
      color: #fdfdfb;
    }

    .market-summary-card span {
      display: block;
      color: rgba(253, 253, 251, .62);
      font-size: .78rem;
      text-transform: uppercase;
    }

    .market-summary-card strong {
      display: block;
      margin-top: 8px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 2.35rem;
      line-height: 1;
    }

    .market-cards {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
    }

    .market-card {
      position: relative;
      overflow: hidden;
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      background: var(--sheet);
      box-shadow: var(--shadow);
      padding: 18px;
    }

    .market-card::before {
      content: "";
      position: absolute;
      inset: 0 0 auto;
      height: 7px;
      background: var(--agency-color, var(--red));
    }

    .market-card h3 {
      margin: 10px 0 10px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 1.55rem;
    }

    .market-metric {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 8px;
      margin: 12px 0;
    }

    .market-note {
      min-height: 72px;
      color: var(--muted);
      font-size: .88rem;
      line-height: 1.42;
    }

    .market-table td:first-child {
      font-weight: 800;
    }

    .source-list {
      margin: 0;
      padding-left: 20px;
      color: var(--muted);
      line-height: 1.6;
      font-size: .9rem;
    }

    .source-list a {
      color: var(--petrol);
      font-weight: 800;
      text-decoration: none;
    }

    .potential-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }

    .swot-card {
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      background: var(--sheet);
      box-shadow: var(--shadow);
      padding: 18px;
      min-height: 270px;
    }

    .swot-card h3 {
      margin: 0 0 12px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 1.35rem;
    }

    .swot-mark {
      display: inline-grid;
      place-items: center;
      width: 34px;
      height: 34px;
      margin-bottom: 12px;
      border-radius: 4px;
      background: #171716;
      color: #fdfdfb;
      font-weight: 900;
    }

    .swot-card ul,
    .assumption-list {
      margin: 0;
      padding-left: 18px;
      color: var(--muted);
      line-height: 1.48;
      font-size: .9rem;
    }

    .potential-summary {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 10px;
      margin-top: 14px;
    }

    .masthead {
      display: grid;
      grid-template-columns: minmax(0, 1.15fr) minmax(300px, .85fr);
      gap: 22px;
      align-items: stretch;
      min-height: 400px;
      margin-bottom: 18px;
    }

    .title-panel,
    .source-panel,
    .card,
    .insight,
    .chart-card,
    .agency-card {
      background: var(--sheet);
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      box-shadow: var(--shadow);
    }

    .title-panel {
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      padding: 30px;
      position: relative;
      overflow: hidden;
    }

    .title-panel::before {
      content: "";
      position: absolute;
      inset: 0 0 auto;
      height: 9px;
      background: linear-gradient(90deg, var(--red), var(--petrol) 45%, var(--amber));
    }

    .eyebrow {
      color: var(--red);
      font-size: .76rem;
      font-weight: 800;
      text-transform: uppercase;
      letter-spacing: 0;
    }

    h1,
    h2,
    h3,
    p {
      margin-top: 0;
    }

    h1 {
      max-width: 980px;
      margin-bottom: 16px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 4.5rem;
      line-height: .98;
      font-weight: 700;
      letter-spacing: 0;
    }

    .lead {
      max-width: 860px;
      color: #31312f;
      font-size: 1.08rem;
      line-height: 1.58;
    }

    .headline-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
      margin-top: 26px;
    }

    .headline-stat {
      border-top: 2px solid var(--ink);
      padding-top: 10px;
    }

    .headline-stat strong {
      display: block;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 2rem;
      line-height: 1.05;
    }

    .headline-stat span {
      display: block;
      margin-top: 5px;
      color: var(--muted);
      font-size: .8rem;
      line-height: 1.35;
    }

    .source-panel {
      display: grid;
      grid-template-rows: auto 1fr auto;
      padding: 22px;
      gap: 18px;
      background:
        linear-gradient(135deg, rgba(221, 63, 45, .08), transparent 36%),
        linear-gradient(225deg, rgba(0, 111, 118, .09), transparent 48%),
        var(--sheet);
    }

    .source-top {
      display: flex;
      justify-content: space-between;
      gap: 14px;
      align-items: flex-start;
      color: var(--muted);
      font-size: .85rem;
      line-height: 1.35;
    }

    .control-total {
      align-self: center;
      border: 1px solid #8f1f16;
      background:
        linear-gradient(135deg, #e44632 0%, #b82418 58%, #7f1711 100%);
      color: #fdfdfb;
      border-radius: 6px;
      padding: 22px;
      box-shadow: 0 20px 40px rgba(221, 63, 45, .28);
    }

    .control-total .label {
      color: rgba(253, 253, 251, .78);
      font-size: .8rem;
      text-transform: uppercase;
      letter-spacing: 0;
    }

    .control-total .value {
      display: block;
      margin-top: 10px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 3.25rem;
      line-height: 1;
    }

    .control-total .sub {
      display: none;
    }

    .chapter-tabs {
      display: flex;
      flex-wrap: wrap;
      gap: 8px;
    }

    .chapter-tabs a,
    .chip,
    button,
    select,
    input {
      border: 1px solid var(--line-dark);
      border-radius: 4px;
      background: var(--sheet);
      color: var(--ink);
      font: inherit;
    }

    .chapter-tabs a {
      display: inline-flex;
      align-items: center;
      min-height: 34px;
      padding: 7px 11px;
      text-decoration: none;
      font-size: .84rem;
      font-weight: 700;
      transition: background .18s ease, color .18s ease, transform .18s ease;
    }

    .chapter-tabs a:hover {
      background: var(--ink);
      color: var(--sheet);
      transform: translateY(-1px);
    }

    .section {
      margin-top: 22px;
    }

    .section-heading {
      display: grid;
      grid-template-columns: minmax(0, 1fr) auto;
      gap: 18px;
      align-items: end;
      margin: 0 0 12px;
      padding: 0 2px;
    }

    .section-heading h2 {
      margin: 0;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 2.05rem;
      line-height: 1.08;
      letter-spacing: 0;
    }

    .section-heading p {
      max-width: 760px;
      margin: 6px 0 0;
      color: var(--muted);
      line-height: 1.5;
    }

    .section-tag {
      font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
      font-size: .78rem;
      color: var(--muted);
    }

    .insight-grid {
      display: grid;
      grid-template-columns: repeat(4, minmax(0, 1fr));
      gap: 12px;
    }

    .insight {
      min-height: 168px;
      padding: 18px;
      display: flex;
      flex-direction: column;
      justify-content: space-between;
      box-shadow: none;
    }

    .insight:nth-child(1) {
      background: linear-gradient(180deg, var(--soft-red), var(--sheet) 52%);
    }

    .insight:nth-child(2) {
      background: linear-gradient(180deg, var(--soft-petrol), var(--sheet) 52%);
    }

    .insight:nth-child(3) {
      background: linear-gradient(180deg, var(--soft-amber), var(--sheet) 52%);
    }

    .insight b {
      display: block;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 2rem;
      line-height: 1.05;
    }

    .insight span {
      color: var(--muted);
      font-size: .92rem;
      line-height: 1.45;
    }

    .dashboard-grid {
      display: grid;
      grid-template-columns: repeat(12, minmax(0, 1fr));
      gap: 14px;
      align-items: stretch;
    }

    .chart-card {
      padding: 18px;
      overflow: hidden;
    }

    .span-4 {
      grid-column: span 4;
    }

    .span-5 {
      grid-column: span 5;
    }

    .span-6 {
      grid-column: span 6;
    }

    .span-7 {
      grid-column: span 7;
    }

    .span-8 {
      grid-column: span 8;
    }

    .span-12 {
      grid-column: span 12;
    }

    .chart-title {
      display: flex;
      justify-content: space-between;
      align-items: baseline;
      gap: 12px;
      margin-bottom: 14px;
      border-bottom: 1px solid var(--line);
      padding-bottom: 11px;
    }

    .chart-title h3 {
      margin: 0;
      font-size: 1.02rem;
      line-height: 1.2;
    }

    .chart-title span {
      color: var(--muted);
      font-size: .8rem;
      white-space: nowrap;
    }

    .donut-wrap {
      display: grid;
      grid-template-columns: 220px minmax(0, 1fr);
      gap: 16px;
      align-items: center;
      min-height: 230px;
    }

    .donut svg,
    .line svg,
    .monthly svg {
      width: 100%;
      height: auto;
      display: block;
    }

    .legend {
      display: grid;
      gap: 10px;
      align-content: center;
    }

    .legend-item {
      display: grid;
      grid-template-columns: 12px minmax(0, 1fr) auto;
      gap: 9px;
      align-items: center;
      font-size: .9rem;
    }

    .swatch {
      width: 12px;
      height: 12px;
      border-radius: 2px;
      border: 1px solid rgba(23, 23, 22, .35);
    }

    .legend small {
      display: block;
      color: var(--muted);
      font-size: .76rem;
      line-height: 1.35;
    }

    .agency-grid {
      display: grid;
      grid-template-columns: repeat(3, minmax(0, 1fr));
      gap: 14px;
    }

    .agency-card {
      padding: 18px;
      position: relative;
      overflow: hidden;
    }

    .agency-card::before {
      content: "";
      position: absolute;
      inset: 0 0 auto;
      height: 7px;
      background: var(--agency-color, var(--red));
    }

    .agency-card h3 {
      margin: 10px 0 12px;
      font-family: Georgia, "Times New Roman", serif;
      font-size: 1.7rem;
      letter-spacing: 0;
    }

    .agency-main {
      display: grid;
      grid-template-columns: 1fr 1fr;
      gap: 10px;
      margin-bottom: 16px;
    }

    .mini-stat {
      border: 1px solid var(--line);
      border-radius: 4px;
      padding: 10px;
      min-width: 0;
    }

    .mini-stat strong {
      display: block;
      font-size: 1.25rem;
      line-height: 1.1;
      word-break: break-word;
    }

    .mini-stat span {
      display: block;
      margin-top: 4px;
      color: var(--muted);
      font-size: .76rem;
      line-height: 1.3;
    }

    .agency-note {
      margin: 0 0 14px;
      color: #30302d;
      line-height: 1.45;
    }

    .spark {
      display: grid;
      gap: 7px;
      margin-top: 10px;
    }

    .spark-row {
      display: grid;
      grid-template-columns: 74px 1fr auto;
      align-items: center;
      gap: 9px;
      color: var(--muted);
      font-size: .77rem;
    }

    .track {
      height: 8px;
      background: rgba(23, 23, 22, .09);
      border-radius: 2px;
      overflow: hidden;
    }

    .fill {
      height: 100%;
      width: var(--w, 0%);
      background: var(--agency-color, var(--red));
    }

    .bar-list {
      display: grid;
      gap: 9px;
    }

    .bar-row {
      display: grid;
      grid-template-columns: minmax(190px, 1fr) minmax(220px, 2fr) minmax(95px, auto);
      gap: 12px;
      align-items: center;
      min-height: 30px;
      font-size: .9rem;
    }

    .bar-label {
      min-width: 0;
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
    }

    .bar-track {
      height: 17px;
      border: 1px solid var(--line-dark);
      background: rgba(255, 255, 255, .6);
      border-radius: 3px;
      overflow: hidden;
    }

    .bar-fill {
      height: 100%;
      min-width: 2px;
      background: linear-gradient(90deg, var(--red), var(--petrol));
    }

    .bar-value {
      text-align: right;
      font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
      font-size: .82rem;
      color: var(--muted);
    }

    .matrix {
      display: grid;
      gap: 8px;
    }

    .matrix-head,
    .matrix-row {
      display: grid;
      grid-template-columns: minmax(190px, 1.5fr) repeat(3, minmax(92px, 1fr)) 78px;
      gap: 8px;
      align-items: center;
    }

    .matrix-head {
      color: var(--muted);
      font-size: .74rem;
      text-transform: uppercase;
      letter-spacing: 0;
      padding-bottom: 3px;
      border-bottom: 1px solid var(--line);
    }

    .matrix-name {
      overflow: hidden;
      text-overflow: ellipsis;
      white-space: nowrap;
      font-size: .88rem;
    }

    .matrix-cell {
      height: 30px;
      border: 1px solid var(--line);
      border-radius: 3px;
      position: relative;
      overflow: hidden;
      background: #f7f4ed;
    }

    .matrix-cell span {
      position: relative;
      z-index: 1;
      display: block;
      padding: 6px 7px;
      font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
      font-size: .72rem;
      color: #222;
    }

    .matrix-cell::before {
      content: "";
      position: absolute;
      inset: 0 auto 0 0;
      width: var(--w, 0%);
      background: var(--agency-color, var(--red));
      opacity: .28;
    }

    .agency-count {
      justify-self: end;
      min-width: 32px;
      text-align: center;
      padding: 5px 7px;
      border: 1px solid var(--line-dark);
      border-radius: 4px;
      background: #171716;
      color: #fdfdfb;
      font-size: .78rem;
    }

    .monthly-chart {
      min-height: 300px;
    }

    .table-tools {
      display: grid;
      grid-template-columns: minmax(240px, 1fr) 190px 120px;
      gap: 10px;
      margin-bottom: 12px;
    }

    input,
    select {
      width: 100%;
      min-height: 38px;
      padding: 8px 10px;
      background: #fdfdfb;
      outline: none;
    }

    input:focus,
    select:focus {
      border-color: var(--red);
      box-shadow: 0 0 0 3px rgba(221, 63, 45, .13);
    }

    .table-wrap {
      max-height: 620px;
      overflow: auto;
      border: 1px solid var(--line-dark);
      border-radius: 6px;
      background: var(--sheet);
    }

    table {
      width: 100%;
      border-collapse: collapse;
      min-width: 980px;
      font-size: .86rem;
    }

    th,
    td {
      padding: 9px 10px;
      border-bottom: 1px solid var(--line);
      vertical-align: top;
      text-align: left;
    }

    th {
      position: sticky;
      top: 0;
      z-index: 2;
      background: #171716;
      color: #fdfdfb;
      font-size: .72rem;
      text-transform: uppercase;
      letter-spacing: 0;
      cursor: pointer;
      user-select: none;
    }

    td.num,
    th.num {
      text-align: right;
      font-variant-numeric: tabular-nums;
      white-space: nowrap;
    }

    tbody tr:hover {
      background: rgba(221, 63, 45, .065);
    }

    .dealer-name {
      max-width: 340px;
      font-weight: 700;
      line-height: 1.35;
    }

    .muted {
      color: var(--muted);
    }

    .badge {
      display: inline-flex;
      align-items: center;
      min-height: 24px;
      padding: 4px 7px;
      border-radius: 4px;
      background: rgba(23, 23, 22, .06);
      font-size: .78rem;
      line-height: 1.2;
      white-space: nowrap;
    }

    .group-table .agency-split {
      display: flex;
      min-width: 160px;
      height: 18px;
      border: 1px solid var(--line-dark);
      border-radius: 3px;
      overflow: hidden;
    }

    .split-segment {
      height: 100%;
      width: var(--w);
      min-width: 2px;
      background: var(--agency-color);
    }

    .group-table th:first-child,
    .group-table td:first-child {
      width: 44px;
      text-align: center;
    }

    .expand-btn {
      width: 28px;
      height: 28px;
      display: inline-grid;
      place-items: center;
      padding: 0;
      border-radius: 4px;
      background: #171716;
      color: #fdfdfb;
      cursor: pointer;
      transition: transform .16s ease, background .16s ease;
    }

    .expand-btn:hover {
      background: var(--red);
      transform: translateY(-1px);
    }

    .group-summary-row.is-open {
      background: rgba(221, 63, 45, .07);
    }

    .group-name-cell {
      max-width: 360px;
      font-weight: 700;
      line-height: 1.35;
    }

    .dealer-detail-row td {
      padding: 0;
      background: #f6f8f6;
    }

    .dealer-detail {
      padding: 14px 14px 18px;
      border-top: 1px solid var(--line-dark);
    }

    .dealer-detail-head {
      display: flex;
      justify-content: space-between;
      gap: 12px;
      align-items: baseline;
      margin-bottom: 10px;
      color: var(--muted);
      font-size: .82rem;
    }

    .dealer-detail-table {
      min-width: 820px;
      background: var(--sheet);
      border: 1px solid var(--line);
    }

    .dealer-detail-table th {
      position: static;
      background: #2a2a27;
    }

    .dealer-detail-table td {
      padding: 8px 9px;
    }

    .dealer-share {
      min-width: 130px;
    }

    .dealer-share-track {
      height: 8px;
      margin-top: 5px;
      background: rgba(23, 23, 22, .09);
      border-radius: 2px;
      overflow: hidden;
    }

    .dealer-share-fill {
      height: 100%;
      width: var(--w);
      background: var(--agency-color, var(--red));
    }

    .footnote {
      margin-top: 20px;
      padding-top: 14px;
      border-top: 1px solid var(--line-dark);
      color: var(--muted);
      line-height: 1.5;
      font-size: .86rem;
    }

    @media (max-width: 1180px) {
      .masthead,
      .market-hero,
      .dashboard-grid,
      .market-cards,
      .potential-grid,
      .potential-summary,
      .agency-grid,
      .insight-grid {
        grid-template-columns: 1fr;
      }

      .span-4,
      .span-5,
      .span-6,
      .span-7,
      .span-8,
      .span-12 {
        grid-column: auto;
      }

      .headline-grid {
        grid-template-columns: repeat(2, minmax(0, 1fr));
      }
    }

    @media (max-width: 760px) {
      .page {
        width: min(100% - 18px, 1480px);
        padding-top: 10px;
      }

      .title-panel,
      .source-panel,
      .chart-card,
      .agency-card,
      .insight {
        padding: 14px;
      }

      h1 {
        font-size: 2.55rem;
      }

      .headline-grid,
      .agency-main,
      .donut-wrap,
      .table-tools,
      .research-intro,
      .section-heading {
        grid-template-columns: 1fr;
      }

      .study-switch {
        width: 100%;
      }

      .study-tab {
        flex: 1;
        padding-inline: 8px;
      }

      .research-frame {
        min-height: 520px;
        height: 72vh;
      }

      .control-total .value {
        font-size: 2.35rem;
      }

      .bar-row {
        grid-template-columns: 1fr;
        gap: 5px;
      }

      .bar-value {
        text-align: left;
      }

      .matrix {
        overflow-x: auto;
        padding-bottom: 2px;
      }

      .matrix-head,
      .matrix-row {
        min-width: 720px;
      }
    }

    @media print {
      body {
        background: #fff;
      }

      .page {
        width: 100%;
        padding: 0;
      }

      .title-panel,
      .source-panel,
      .card,
      .insight,
      .chart-card,
      .agency-card {
        box-shadow: none;
      }

      .chapter-tabs,
      .study-switch,
      .table-tools {
        display: none;
      }

      .table-wrap {
        max-height: none;
        overflow: visible;
      }
    }
  </style>
</head>
<body>
  <main class="page">
    <header class="study-header">
      <h1 class="study-title">Работа агентств с клиентами Авто.ру</h1>
    </header>
    <nav class="study-switch" aria-label="Вкладки исследования">
      <button class="study-tab is-active" type="button" data-study-tab="dashboard">Дашборд</button>
      <button class="study-tab" type="button" data-study-tab="research">Агентства</button>
      <button class="study-tab" type="button" data-study-tab="market">Анализ выручки агентств</button>
      <button class="study-tab" type="button" data-study-tab="potential">Расчет потенциала агентства Авто.ру</button>
      <button class="study-tab" type="button" data-study-tab="economics">Экономическая модель</button>
    </nav>

    <div class="tab-panel is-active" id="dashboardPanel" data-tab-panel="dashboard">
    <section class="masthead" id="top">
      <div class="title-panel">
        <div>
          <div class="eyebrow">Дилеры • группы компаний • агентства</div>
          <h1>Кто управляет выручкой дилеров в Q1 2026</h1>
          <p class="lead">Сводный дашборд по агентствам постинга: выручка, платные услуги, дилерские группы, регионы присутствия и концентрация управления внутри сети.</p>
        </div>
        <div class="headline-grid" id="headlineStats"></div>
      </div>
      <aside class="source-panel">
        <div class="source-top">
          <div>
            <b>Источник выгрузки</b><br>
            <span id="sourceName"></span>
          </div>
          <div>
            <b>Период</b><br>
            <span id="periodName"></span>
          </div>
        </div>
        <div class="control-total">
          <span class="label">Выручка под управлением агентств</span>
          <span class="value" id="managedRevenue"></span>
        </div>
        <nav class="chapter-tabs" aria-label="Разделы дашборда">
          <a href="#summary">Итоги</a>
          <a href="#agencies">Агентства</a>
          <a href="#regions">Регионы</a>
          <a href="#groups">ГК и дилеры</a>
        </nav>
      </aside>
    </section>

    <section class="section" id="summary">
      <div class="section-heading">
        <div>
          <h2>Главные выводы</h2>
          <p>Разделение по деньгам и платным услугам показывает разные типы контроля: одно агентство забирает основную выручку, другое — основной операционный объём.</p>
        </div>
        <span class="section-tag">Executive layer</span>
      </div>
      <div class="insight-grid" id="insightGrid"></div>
    </section>

    <section class="section" id="agencies">
      <div class="section-heading">
        <div>
          <h2>Доли агентств</h2>
          <p>Круговые диаграммы сравнивают финансовую долю и долю платных услуг: это две разные картины управления.</p>
        </div>
        <span class="section-tag">Agency split</span>
      </div>
      <div class="dashboard-grid">
        <article class="chart-card span-6">
          <div class="chart-title">
            <h3>Доля выручки Q1</h3>
            <span>по агентствам</span>
          </div>
          <div class="donut-wrap">
            <div class="donut" id="revenueDonut"></div>
            <div class="legend" id="revenueLegend"></div>
          </div>
        </article>
        <article class="chart-card span-6">
          <div class="chart-title">
            <h3>Доля платных услуг Q1</h3>
            <span>по агентствам</span>
          </div>
          <div class="donut-wrap">
            <div class="donut" id="servicesDonut"></div>
            <div class="legend" id="servicesLegend"></div>
          </div>
        </article>
        <div class="span-12 agency-grid" id="agencyCards"></div>
        <article class="chart-card span-12">
          <div class="chart-title">
            <h3>Помесячная динамика выручки</h3>
            <span>январь — март</span>
          </div>
          <div class="monthly monthly-chart" id="monthlyRevenue"></div>
        </article>
      </div>
    </section>

    <section class="section" id="regions">
      <div class="section-heading">
        <div>
          <h2>Региональная концентрация</h2>
          <p>Сверху — регионы с крупнейшей управляемой выручкой, ниже — матрица распределения агентств по ключевым территориям.</p>
        </div>
        <span class="section-tag">Russia footprint</span>
      </div>
      <div class="dashboard-grid">
        <article class="chart-card span-7">
          <div class="chart-title">
            <h3>Топ регионов по выручке</h3>
            <span>Q1 2026</span>
          </div>
          <div class="bar-list" id="regionBars"></div>
        </article>
        <article class="chart-card span-5">
          <div class="chart-title">
            <h3>Где агентств больше и меньше</h3>
            <span>присутствие</span>
          </div>
          <div id="geoNotes"></div>
        </article>
        <article class="chart-card span-12">
          <div class="chart-title">
            <h3>Матрица выручки по регионам и агентствам</h3>
            <span>топ-20 регионов</span>
          </div>
          <div class="matrix" id="regionMatrix"></div>
        </article>
      </div>
    </section>

    <section class="section" id="groups">
      <div class="section-heading">
        <div>
          <h2>Декомпозиция по ГК и дилерам</h2>
          <p>Нажмите на группу компаний, чтобы раскрыть дилеров внутри неё: видна выручка каждого дилера, его агентство и доля в выручке выбранной ГК.</p>
        </div>
        <span class="section-tag">Company groups + dealer drilldown</span>
      </div>
      <article class="chart-card span-12">
        <div class="chart-title">
          <h3>Группы компаний</h3>
          <span>раскрывающийся список дилеров</span>
        </div>
        <div class="table-wrap group-table">
          <table>
            <thead>
              <tr>
                <th></th>
                <th>ГК</th>
                <th>Ведущее агентство</th>
                <th class="num">Выручка</th>
                <th class="num">Услуги</th>
                <th class="num">Дилеры</th>
                <th>Доля агентств внутри ГК</th>
              </tr>
            </thead>
            <tbody id="groupTable"></tbody>
          </table>
        </div>
      </article>
    </section>

    <p class="footnote">Данные рассчитаны по строкам листа «Chart data». Пустые значения месячных показателей интерпретированы как 0; все доли считаются от квартального итога по выгрузке.</p>
    </div>

    <div class="tab-panel research-shell" id="researchPanel" data-tab-panel="research" hidden>
      <section class="research-intro">
        <div>
          <h2>Исследование агентств</h2>
          <p>Презентационное сравнение CarCopy, Четыре пикселя и ТАК / Tandem Group по модели работы, публичным ценам, позиционированию и критериям выбора подрядчика.</p>
        </div>
        <span>Presentation mode</span>
      </section>
      <iframe class="research-frame" id="researchFrame" title="Исследование агентств по классифайдам" sandbox="allow-scripts allow-popups allow-popups-to-escape-sandbox"></iframe>
    </div>

    <div class="tab-panel market-panel" id="marketPanel" data-tab-panel="market" hidden>
      <section class="market-hero">
        <div class="market-hero-main">
          <div class="eyebrow">Расчётная модель</div>
          <h2>Выручка агентств</h2>
          <p id="marketDefinition"></p>
        </div>
        <div class="market-summary-card">
          <span>Оценка месячной выручки рынка</span>
          <strong id="marketTotal"></strong>
          <p id="marketTotalSub"></p>
        </div>
      </section>
      <section class="market-cards" id="marketCards"></section>
      <article class="chart-card">
        <div class="chart-title">
          <h3>Короткая сводка по агентствам</h3>
          <span>салоны × тариф сопровождения</span>
        </div>
        <div class="table-wrap market-table">
          <table>
            <thead>
              <tr>
                <th>Агентство</th>
                <th class="num">Салонов</th>
                <th class="num">ГК</th>
                <th>Тариф / допущение</th>
                <th class="num">Оценка в месяц</th>
                <th class="num">Оценка в год</th>
              </tr>
            </thead>
            <tbody id="marketTable"></tbody>
          </table>
        </div>
      </article>
      <article class="chart-card">
        <div class="chart-title">
          <h3>Источники и ограничения</h3>
          <span>для расчёта</span>
        </div>
        <ul class="source-list" id="marketSources"></ul>
      </article>
    </div>

    <div class="tab-panel market-panel" id="potentialPanel" data-tab-panel="potential" hidden>
      <section class="market-hero">
        <div class="market-hero-main">
          <div class="eyebrow">Сценарная гипотеза</div>
          <h2>Расчет потенциала агентства Авто.ру</h2>
          <p id="potentialSubtitle"></p>
        </div>
        <div class="market-summary-card">
          <span>Потенциальная месячная выручка</span>
          <strong id="potentialRevenue"></strong>
          <p id="potentialRevenueSub"></p>
        </div>
      </section>
      <section class="potential-summary" id="potentialSummary"></section>
      <section class="potential-grid" id="swotGrid"></section>
      <article class="chart-card">
        <div class="chart-title">
          <h3>Ключевые допущения</h3>
          <span>для первого пилота</span>
        </div>
        <ul class="assumption-list" id="potentialAssumptions"></ul>
      </article>
    </div>

    <div class="tab-panel market-panel" id="economicsPanel" data-tab-panel="economics" hidden>
      <section class="market-hero">
        <div class="market-hero-main">
          <div class="eyebrow">Unit-экономика</div>
          <h2>Экономическая модель</h2>
          <p id="economicsSubtitle"></p>
        </div>
        <div class="market-summary-card">
          <span>Цель на сотрудника</span>
          <strong id="employeeRevenueTarget"></strong>
          <p id="employeeRevenueTargetSub"></p>
        </div>
      </section>
      <article class="chart-card">
        <div class="chart-title">
          <h3>ФОТ пилотной команды</h3>
          <span>зарплата на руки + НДФЛ + отчисления</span>
        </div>
        <div class="potential-summary" id="economicsSummary"></div>
      </article>
      <article class="chart-card">
        <div class="chart-title">
          <h3>Сколько салонов нужно вести</h3>
          <span>3 сценария отдела</span>
        </div>
        <div class="table-wrap market-table">
          <table>
            <thead>
              <tr>
                <th>Сценарий</th>
                <th class="num">Салонов нужно вести</th>
                <th class="num">Выручка в месяц</th>
                <th>Логика расчёта</th>
              </tr>
            </thead>
            <tbody id="economicsTable"></tbody>
          </table>
        </div>
      </article>
      <article class="chart-card">
        <div class="chart-title">
          <h3>Как сотруднику приносить от 1 млн ₽/мес</h3>
          <span>чек × загрузка</span>
        </div>
        <div class="table-wrap market-table">
          <table>
            <thead>
              <tr>
                <th>Модель</th>
                <th class="num">Салонов на сотрудника</th>
                <th class="num">Средний чек</th>
                <th class="num">Выручка / сотрудник</th>
                <th>Что для этого нужно</th>
              </tr>
            </thead>
            <tbody id="unitEconomicsTable"></tbody>
          </table>
        </div>
      </article>
      <article class="chart-card">
        <div class="chart-title">
          <h3>План продаж на 6 месяцев</h3>
          <span>аккаунт-менеджеры и разовая выплата</span>
        </div>
        <div class="potential-summary" id="salesCompSummary"></div>
        <div class="table-wrap market-table" style="margin-top: 12px;">
          <table>
            <thead>
              <tr>
                <th>Месяц</th>
                <th class="num">Новые салоны: база</th>
                <th class="num">Накоплено: база</th>
                <th class="num">Новые салоны: цель</th>
                <th class="num">Накоплено: цель</th>
                <th class="num">Выплата за подключения</th>
                <th class="num">MRR к концу месяца</th>
              </tr>
            </thead>
            <tbody id="salesPlanTable"></tbody>
          </table>
        </div>
      </article>
    </div>
  </main>

  <script>
    const DASHBOARD_DATA = __DASHBOARD_DATA__;
    const RESEARCH_PRESENTATION = __RESEARCH_PRESENTATION__;
    const COLORS = {
      "4 пикселя +": "#dd3f2d",
      "Каркопи": "#006f76",
      "ТАК": "#c88b13"
    };
    const FALLBACK_COLORS = ["#395f9f", "#497a45", "#7b4a8a"];
    const state = { openGroups: new Set() };

    const $ = (id) => document.getElementById(id);
    const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (m) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[m]));
    const colorFor = (name, idx = 0) => COLORS[name] || FALLBACK_COLORS[idx % FALLBACK_COLORS.length];
    const nf = new Intl.NumberFormat("ru-RU");
    const moneyFull = (n) => `${nf.format(Math.round(n))} ₽`;
    const money = (n) => {
      if (Math.abs(n) >= 1_000_000) return `${(n / 1_000_000).toLocaleString("ru-RU", { maximumFractionDigits: 1 })} млн ₽`;
      if (Math.abs(n) >= 1_000) return `${(n / 1_000).toLocaleString("ru-RU", { maximumFractionDigits: 1 })} тыс. ₽`;
      return `${n.toLocaleString("ru-RU", { maximumFractionDigits: 0 })} ₽`;
    };
    const compactNum = (n) => {
      if (Math.abs(n) >= 1_000_000) return `${(n / 1_000_000).toLocaleString("ru-RU", { maximumFractionDigits: 1 })} млн`;
      if (Math.abs(n) >= 1_000) return `${(n / 1_000).toLocaleString("ru-RU", { maximumFractionDigits: 1 })} тыс.`;
      return nf.format(Math.round(n));
    };
    const percent = (n, digits = 1) => `${(n * 100).toLocaleString("ru-RU", { maximumFractionDigits: digits })}%`;

    function init() {
      setupStudyTabs();
      renderHeader();
      renderInsights();
      renderDonut("revenueDonut", "revenueLegend", DASHBOARD_DATA.agencies, "revenue", DASHBOARD_DATA.meta.totalRevenue);
      renderDonut("servicesDonut", "servicesLegend", DASHBOARD_DATA.agencies, "services", DASHBOARD_DATA.meta.totalServices);
      renderAgencyCards();
      renderMonthlyRevenue();
      renderRegionBars();
      renderGeoNotes();
      renderRegionMatrix();
      renderGroupTable();
      renderMarketAnalysis();
      renderPotentialAnalysis();
      renderEconomicsModel();
    }

    function setupStudyTabs() {
      const tabs = Array.from(document.querySelectorAll("[data-study-tab]"));
      const panels = Array.from(document.querySelectorAll("[data-tab-panel]"));
      const frame = $("researchFrame");
      if (frame && !frame.srcdoc) frame.srcdoc = RESEARCH_PRESENTATION;

      const activate = (name) => {
        tabs.forEach(tab => tab.classList.toggle("is-active", tab.dataset.studyTab === name));
        panels.forEach(panel => {
          panel.hidden = panel.dataset.tabPanel !== name;
        });
      };

      tabs.forEach(tab => tab.addEventListener("click", () => activate(tab.dataset.studyTab)));
      if (location.hash === "#research") activate("research");
      if (location.hash === "#market") activate("market");
      if (location.hash === "#potential") activate("potential");
      if (location.hash === "#economics") activate("economics");
    }

    function renderHeader() {
      const meta = DASHBOARD_DATA.meta;
      $("sourceName").textContent = meta.source;
      $("periodName").textContent = `${meta.period} • ${meta.months.join(", ")}`;
      $("managedRevenue").textContent = money(meta.managedRevenue);

      const stats = [
        [money(meta.totalRevenue), "общая квартальная выручка"],
        [compactNum(meta.totalServices), "платных услуг за квартал"],
        [nf.format(meta.totalDealers), "дилера в детализации"],
        [nf.format(meta.totalRegions), "региона присутствия"]
      ];
      $("headlineStats").innerHTML = stats.map(([value, label]) => `
        <div class="headline-stat"><strong>${esc(value)}</strong><span>${esc(label)}</span></div>
      `).join("");
    }

    function renderInsights() {
      const meta = DASHBOARD_DATA.meta;
      const topRevenueAgency = [...DASHBOARD_DATA.agencies].sort((a, b) => b.revenue - a.revenue)[0];
      const topServicesAgency = [...DASHBOARD_DATA.agencies].sort((a, b) => b.services - a.services)[0];
      const marchGrowth = (meta.monthRevenue[2] - meta.monthRevenue[1]) / meta.monthRevenue[1];
      const singleAgencyRegions = Number(meta.agencyPresenceCounts["1"] || 0);
      const leastNames = meta.leastRegions.slice(0, 3).map(r => r.name).join(", ");
      const insights = [
        [
          percent(topRevenueAgency.revenueShare),
          `${topRevenueAgency.name} управляет основной долей выручки: ${money(topRevenueAgency.revenue)} через ${topRevenueAgency.dealers} дилера.`
        ],
        [
          percent(topServicesAgency.servicesShare),
          `${topServicesAgency.name} лидирует по платным услугам: ${compactNum(topServicesAgency.services)} услуг и самая широкая география.`
        ],
        [
          percent(meta.capitalsRevenueShare),
          `Москва/МО и Санкт-Петербург/ЛО дают почти всю квартальную выручку, хотя по услугам их доля ${percent(meta.capitalsServicesShare)}.`
        ]
      ];
      $("insightGrid").innerHTML = insights.map(([value, text]) => `
        <article class="insight"><b>${esc(value)}</b><span>${esc(text)}</span></article>
      `).join("");
    }

    function renderDonut(svgId, legendId, rows, key, total) {
      const radius = 76;
      const circumference = 2 * Math.PI * radius;
      let offset = 0;
      const rings = rows.map((row, idx) => {
        const portion = total ? row[key] / total : 0;
        const dash = portion * circumference;
        const stroke = colorFor(row.name, idx);
        const circle = `<circle cx="120" cy="120" r="${radius}" fill="none" stroke="${stroke}" stroke-width="34" stroke-dasharray="${dash} ${circumference - dash}" stroke-dashoffset="${-offset}" transform="rotate(-90 120 120)"></circle>`;
        offset += dash;
        return circle;
      }).join("");
      const top = [...rows].sort((a, b) => b[key] - a[key])[0];
      $(svgId).innerHTML = `
        <svg viewBox="0 0 240 240" role="img" aria-label="Круговая диаграмма">
          <circle cx="120" cy="120" r="${radius}" fill="none" stroke="#e7e2d7" stroke-width="34"></circle>
          ${rings}
          <circle cx="120" cy="120" r="50" fill="#fdfdfb" stroke="#171716" stroke-width="1"></circle>
          <text x="120" y="113" text-anchor="middle" font-size="24" font-family="Georgia, serif" fill="#171716">${percent(top[key] / total)}</text>
          <text x="120" y="136" text-anchor="middle" font-size="11" fill="#656560">${esc(top.name)}</text>
        </svg>`;
      $(legendId).innerHTML = rows.map((row, idx) => `
        <div class="legend-item">
          <span class="swatch" style="background:${colorFor(row.name, idx)}"></span>
          <div><b>${esc(row.name)}</b><small>${key === "revenue" ? money(row[key]) : compactNum(row[key]) + " услуг"}</small></div>
          <strong>${percent(row[key] / total)}</strong>
        </div>
      `).join("");
    }

    function renderAgencyCards() {
      const maxRevenue = Math.max(...DASHBOARD_DATA.agencies.map(a => a.revenue));
      const maxServices = Math.max(...DASHBOARD_DATA.agencies.map(a => a.services));
      $("agencyCards").innerHTML = DASHBOARD_DATA.agencies.map((agency, idx) => {
        const color = colorFor(agency.name, idx);
        const topGroup = agency.topGroups[0];
        const topRegion = agency.topRegions[0];
        return `
          <article class="agency-card" style="--agency-color:${color}">
            <h3>${esc(agency.name)}</h3>
            <div class="agency-main">
              <div class="mini-stat"><strong>${money(agency.revenue)}</strong><span>${percent(agency.revenueShare)} выручки</span></div>
              <div class="mini-stat"><strong>${compactNum(agency.services)}</strong><span>${percent(agency.servicesShare)} платных услуг</span></div>
              <div class="mini-stat"><strong>${nf.format(agency.dealers)}</strong><span>дилеров</span></div>
              <div class="mini-stat"><strong>${nf.format(agency.regions)}</strong><span>регионов</span></div>
            </div>
            <p class="agency-note">Главный источник выручки: <b>${esc(topGroup.name)}</b> (${percent(topGroup.shareInsideAgency)} внутри агентства). Региональный центр: <b>${esc(topRegion.name)}</b>.</p>
            <div class="spark">
              <div class="spark-row"><span>Выручка</span><div class="track"><div class="fill" style="--w:${agency.revenue / maxRevenue * 100}%"></div></div><b>${money(agency.revenue)}</b></div>
              <div class="spark-row"><span>Услуги</span><div class="track"><div class="fill" style="--w:${agency.services / maxServices * 100}%"></div></div><b>${compactNum(agency.services)}</b></div>
              <div class="spark-row"><span>₽ / усл.</span><div class="track"><div class="fill" style="--w:${Math.min(100, agency.revenuePerService / 650 * 100)}%"></div></div><b>${money(agency.revenuePerService)}</b></div>
            </div>
          </article>
        `;
      }).join("");
    }

    function renderMonthlyRevenue() {
      const meta = DASHBOARD_DATA.meta;
      const months = meta.months;
      const width = 940;
      const height = 300;
      const pad = { left: 58, top: 24, right: 22, bottom: 48 };
      const chartW = width - pad.left - pad.right;
      const chartH = height - pad.top - pad.bottom;
      const maxTotal = Math.max(...meta.monthRevenue);
      const barW = 116;
      const gap = chartW / months.length;
      let svg = `<svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Помесячная динамика выручки">`;
      svg += `<line x1="${pad.left}" y1="${pad.top + chartH}" x2="${width - pad.right}" y2="${pad.top + chartH}" stroke="#171716" stroke-width="1"/>`;
      [0, .25, .5, .75, 1].forEach(t => {
        const y = pad.top + chartH - t * chartH;
        svg += `<line x1="${pad.left}" y1="${y}" x2="${width - pad.right}" y2="${y}" stroke="#cfd6d3" stroke-width="1"/>`;
        svg += `<text x="${pad.left - 8}" y="${y + 4}" text-anchor="end" font-size="11" fill="#656560">${money(maxTotal * t)}</text>`;
      });
      months.forEach((month, monthIdx) => {
        const x = pad.left + gap * monthIdx + gap / 2 - barW / 2;
        let yCursor = pad.top + chartH;
        DASHBOARD_DATA.agencies.forEach((agency, agencyIdx) => {
          const value = agency.revenueByMonth[monthIdx] || 0;
          const h = maxTotal ? value / maxTotal * chartH : 0;
          yCursor -= h;
          svg += `<rect x="${x}" y="${yCursor}" width="${barW}" height="${h}" fill="${colorFor(agency.name, agencyIdx)}"></rect>`;
        });
        svg += `<text x="${x + barW / 2}" y="${height - 24}" text-anchor="middle" font-size="13" fill="#171716">${month}</text>`;
        svg += `<text x="${x + barW / 2}" y="${yCursor - 8}" text-anchor="middle" font-size="12" font-weight="700" fill="#171716">${money(meta.monthRevenue[monthIdx])}</text>`;
      });
      svg += DASHBOARD_DATA.agencies.map((agency, idx) => {
        const x = pad.left + idx * 175;
        return `<rect x="${x}" y="${height - 18}" width="12" height="12" fill="${colorFor(agency.name, idx)}"></rect><text x="${x + 18}" y="${height - 8}" font-size="12" fill="#656560">${esc(agency.name)}</text>`;
      }).join("");
      svg += `</svg>`;
      $("monthlyRevenue").innerHTML = svg;
    }

    function renderRegionBars() {
      const rows = DASHBOARD_DATA.regions.slice(0, 14);
      const max = Math.max(...rows.map(r => r.revenue));
      $("regionBars").innerHTML = rows.map((region) => `
        <div class="bar-row">
          <div class="bar-label" title="${esc(region.name)}"><b>${esc(region.name)}</b><br><span class="muted">${region.agencies} агентств • ${region.dealers} дилеров</span></div>
          <div class="bar-track"><div class="bar-fill" style="width:${region.revenue / max * 100}%"></div></div>
          <div class="bar-value">${money(region.revenue)}<br>${percent(region.revenue / DASHBOARD_DATA.meta.totalRevenue)}</div>
        </div>
      `).join("");
    }

    function renderGeoNotes() {
      const meta = DASHBOARD_DATA.meta;
      const threeAgency = meta.multiAgencyRegions.filter(r => r.agencies === 3).map(r => r.name);
      const twoAgency = meta.multiAgencyRegions.filter(r => r.agencies === 2).map(r => r.name);
      const oneCount = Number(meta.agencyPresenceCounts["1"] || 0);
      const least = meta.leastRegions.map(r => `${r.name} — ${money(r.revenue)}`).join("; ");
      $("geoNotes").innerHTML = `
        <div class="bar-list">
          <div class="mini-stat"><strong>${threeAgency.length}</strong><span>региона с максимальным покрытием: ${esc(threeAgency.join(", "))}</span></div>
          <div class="mini-stat"><strong>${twoAgency.length}</strong><span>регионов с двумя агентствами: ${esc(twoAgency.join(", "))}</span></div>
          <div class="mini-stat"><strong>${oneCount}</strong><span>регионов с одним агентством; это длинный хвост присутствия.</span></div>
          <div class="mini-stat"><strong>${percent(meta.capitalsRevenueShare)}</strong><span>выручки сосредоточено в двух столичных макрорегионах.</span></div>
          <p class="agency-note"><b>Минимальная концентрация выручки:</b> ${esc(least)}.</p>
        </div>
      `;
    }

    function renderRegionMatrix() {
      const agencies = DASHBOARD_DATA.agencies.map(a => a.name);
      const rows = DASHBOARD_DATA.regions.slice(0, 20);
      const maxCell = Math.max(...rows.flatMap(r => agencies.map(a => r.revenueByAgency[a] || 0)));
      const head = `<div class="matrix-head"><div>Регион</div>${agencies.map(a => `<div>${esc(a)}</div>`).join("")}<div>Аг.</div></div>`;
      const body = rows.map((region) => `
        <div class="matrix-row">
          <div class="matrix-name" title="${esc(region.name)}"><b>${esc(region.name)}</b><br><span class="muted">${money(region.revenue)}</span></div>
          ${agencies.map((agency, idx) => {
            const value = region.revenueByAgency[agency] || 0;
            const width = maxCell ? value / maxCell * 100 : 0;
            return `<div class="matrix-cell" style="--w:${width}%; --agency-color:${colorFor(agency, idx)}"><span>${value ? money(value) : "—"}</span></div>`;
          }).join("")}
          <div class="agency-count">${region.agencies}</div>
        </div>
      `).join("");
      $("regionMatrix").innerHTML = head + body;
    }

    function renderGroupTable() {
      const table = $("groupTable");
      table.innerHTML = DASHBOARD_DATA.groups.map((group, groupIndex) => {
        const isOpen = state.openGroups.has(groupIndex);
        const split = DASHBOARD_DATA.agencies.map((agency, idx) => {
          const value = group.agencySplit[agency.name] || 0;
          if (!value) return "";
          return `<span class="split-segment" title="${esc(agency.name)}: ${money(value)}" style="--w:${group.revenue ? value / group.revenue * 100 : 0}%; --agency-color:${colorFor(agency.name, idx)}"></span>`;
        }).join("");
        const groupRow = `
          <tr class="group-summary-row ${isOpen ? "is-open" : ""}">
            <td><button class="expand-btn" type="button" data-group-index="${groupIndex}" title="${isOpen ? "Свернуть дилеров" : "Раскрыть дилеров"}" aria-label="${isOpen ? "Свернуть дилеров" : "Раскрыть дилеров"}">${isOpen ? "−" : "+"}</button></td>
            <td class="group-name-cell">${esc(group.name)}<br><span class="muted">${group.regions} регионов • ${group.agencies} агентств</span></td>
            <td><span class="badge">${esc(group.topAgency)}</span><br><span class="muted">${percent(group.topAgencyShare)} внутри ГК</span></td>
            <td class="num">${moneyFull(group.revenue)}</td>
            <td class="num">${nf.format(Math.round(group.services))}</td>
            <td class="num">${nf.format(group.dealers)}</td>
            <td><div class="agency-split">${split}</div></td>
          </tr>
        `;
        return isOpen ? groupRow + renderDealerDrilldown(group) : groupRow;
      }).join("");
      table.onclick = (event) => {
        const button = event.target.closest(".expand-btn");
        if (!button) return;
        const index = Number(button.dataset.groupIndex);
        if (state.openGroups.has(index)) state.openGroups.delete(index);
        else state.openGroups.add(index);
        renderGroupTable();
      };
    }

    function renderDealerDrilldown(group) {
      const rows = DASHBOARD_DATA.dealers
        .filter(row => row.group === group.name)
        .sort((a, b) => b.revenue - a.revenue);
      const body = rows.map(row => {
        const agencyIndex = DASHBOARD_DATA.agencies.findIndex(agency => agency.name === row.agency);
        const share = group.revenue ? row.revenue / group.revenue : 0;
        return `
          <tr>
            <td class="dealer-name">${esc(row.dealer)}</td>
            <td>${esc(row.region)}</td>
            <td><span class="badge">${esc(row.agency)}</span></td>
            <td class="num"><b>${moneyFull(row.revenue)}</b></td>
            <td class="num">${nf.format(Math.round(row.services))}</td>
            <td class="dealer-share">
              <b>${percent(share)}</b>
              <div class="dealer-share-track"><div class="dealer-share-fill" style="--w:${share * 100}%; --agency-color:${colorFor(row.agency, agencyIndex)}"></div></div>
            </td>
          </tr>
        `;
      }).join("");
      return `
        <tr class="dealer-detail-row">
          <td colspan="7">
            <div class="dealer-detail">
              <div class="dealer-detail-head">
                <b>${nf.format(rows.length)} дилеров внутри «${esc(group.name)}»</b>
                <span>Ведущее агентство ГК: ${esc(group.topAgency)} • ${percent(group.topAgencyShare)} выручки группы</span>
              </div>
              <table class="dealer-detail-table">
                <thead>
                  <tr>
                    <th>Дилер</th>
                    <th>Регион</th>
                    <th>Агентство</th>
                    <th class="num">Выручка Q1</th>
                    <th class="num">Услуги Q1</th>
                    <th>Доля в ГК</th>
                  </tr>
                </thead>
                <tbody>${body}</tbody>
              </table>
            </div>
          </td>
        </tr>
      `;
    }

    function moneyRange(low, high) {
      return Math.round(low) === Math.round(high) ? money(low) : `${money(low)} — ${money(high)}`;
    }

    function renderMarketAnalysis() {
      const market = DASHBOARD_DATA.marketAnalysis;
      $("marketDefinition").textContent = `${market.definition} Оценка показывает агентскую выручку от сопровождения и не включает индивидуальные расширения договора.`;
      $("marketTotal").textContent = moneyRange(market.summary.monthlyLow, market.summary.monthlyHigh);
      $("marketTotalSub").textContent = `В год: ${moneyRange(market.summary.annualLow, market.summary.annualHigh)} при ${nf.format(market.summary.salons)} салонах в модели.`;

      $("marketCards").innerHTML = market.rows.map((row, idx) => {
        const color = colorFor(row.sourceName, idx);
        return `
          <article class="market-card" style="--agency-color:${color}">
            <h3>${esc(row.name)}</h3>
            <div class="market-metric">
              <div class="mini-stat"><strong>${nf.format(row.salons)}</strong><span>салонов на сопровождении</span></div>
              <div class="mini-stat"><strong>${nf.format(row.groups)}</strong><span>групп компаний</span></div>
              <div class="mini-stat"><strong>${moneyRange(row.monthlyLow, row.monthlyHigh)}</strong><span>примерно в месяц</span></div>
              <div class="mini-stat"><strong>${moneyRange(row.annualLow, row.annualHigh)}</strong><span>примерно в год</span></div>
            </div>
            <p class="market-note"><b>Тариф:</b> ${esc(row.priceLabel)}. ${esc(row.note)}</p>
          </article>
        `;
      }).join("");

      $("marketTable").innerHTML = market.rows.map(row => `
        <tr>
          <td>${esc(row.name)}</td>
          <td class="num">${nf.format(row.salons)}</td>
          <td class="num">${nf.format(row.groups)}</td>
          <td>${esc(row.priceLabel)}<br><span class="muted">${esc(row.note)}</span></td>
          <td class="num"><b>${moneyRange(row.monthlyLow, row.monthlyHigh)}</b></td>
          <td class="num">${moneyRange(row.annualLow, row.annualHigh)}</td>
        </tr>
      `).join("");

      $("marketSources").innerHTML = market.sources.map(source => {
        const label = esc(source.label);
        return source.url
          ? `<li><a href="${esc(source.url)}" target="_blank" rel="noreferrer">${label}</a></li>`
          : `<li>${label}</li>`;
      }).join("");
    }

    function renderPotentialAnalysis() {
      const potential = DASHBOARD_DATA.autoruPotential;
      const market = potential.market;
      $("potentialSubtitle").textContent = potential.subtitle;
      $("potentialRevenue").textContent = moneyRange(market.monthlyRevenueLow, market.monthlyRevenueHigh);
      $("potentialRevenueSub").textContent = `${nf.format(market.totalPotentialLow)}–${nf.format(market.totalPotentialHigh)} салонов в сценарии: перехват текущих клиентов + новые города-миллионники.`;

      const summary = [
        [nf.format(market.captureLow) + "–" + nf.format(market.captureHigh), "салонов можно перехватить у текущих агентств"],
        [nf.format(market.newRegionalLow) + "–" + nf.format(market.newRegionalHigh), "новых салонов в городах-миллионниках"],
        [moneyRange(market.priceLow, market.priceHigh), "цена сопровождения за салон в месяц"],
        [`${nf.format(market.staffLow)}–${nf.format(market.staffHigh)}`, `сотрудников по норме 1 человек = ${nf.format(market.staffRatio)} салонов`]
      ];
      $("potentialSummary").innerHTML = summary.map(([value, label]) => `
        <div class="market-summary-card"><span>${esc(label)}</span><strong>${esc(value)}</strong></div>
      `).join("");

      $("swotGrid").innerHTML = potential.swot.map(block => `
        <article class="swot-card">
          <span class="swot-mark">${esc(block.type)}</span>
          <h3>${esc(block.title)}</h3>
          <ul>${block.items.map(item => `<li>${esc(item)}</li>`).join("")}</ul>
        </article>
      `).join("");

      $("potentialAssumptions").innerHTML = potential.assumptions.map(item => `<li>${esc(item)}</li>`).join("");
    }

    function rangeText(low, high, suffix = "") {
      return low === high ? `${nf.format(low)}${suffix}` : `${nf.format(low)}–${nf.format(high)}${suffix}`;
    }

    function renderEconomicsModel() {
      const potential = DASHBOARD_DATA.autoruPotential;
      const market = potential.market;
      const economics = potential.economics;
      $("economicsSubtitle").textContent = `${economics.distribution} Модель показывает, какой объём салонов и какой чек нужны, чтобы экономика сходилась и один сотрудник приносил от 1 млн ₽ выручки в месяц.`;
      $("employeeRevenueTarget").textContent = money(economics.targetRevenuePerEmployee);
      $("employeeRevenueTargetSub").textContent = `Текущая норма ${nf.format(market.staffRatio)} салонов × ${moneyRange(market.priceLow, market.priceHigh)} = ${moneyRange(economics.unitScenarios[0].revenueLow, economics.unitScenarios[0].revenueHigh)} на сотрудника.`;

      const economicsSummary = [
        [money(economics.employeeFullCost), "полная стоимость сотрудника: 120 тыс. ₽ на руки + НДФЛ + отчисления"],
        [money(economics.managerFullCost), "полная стоимость руководителя: x2 от сотрудника"],
        [moneyRange(economics.teamCostLow, economics.teamCostHigh), `ФОТ отдела: ${economics.teamLow}–${economics.teamHigh} сотрудников + руководитель`],
        [`${nf.format(economics.teamLow * market.staffRatio)}–${nf.format(economics.teamHigh * market.staffRatio)}`, "салонов при полной загрузке команды"]
      ];
      $("economicsSummary").innerHTML = economicsSummary.map(([value, label]) => `
        <div class="mini-stat"><strong>${esc(value)}</strong><span>${esc(label)}</span></div>
      `).join("");
      $("economicsTable").innerHTML = economics.scenarios.map(scenario => `
        <tr>
          <td>${esc(scenario.name)}</td>
          <td class="num"><b>${nf.format(scenario.salonsLow)}–${nf.format(scenario.salonsHigh)}</b></td>
          <td class="num">${moneyRange(scenario.revenueLow, scenario.revenueHigh)}</td>
          <td>${esc(scenario.logic)}</td>
        </tr>
      `).join("");
      $("unitEconomicsTable").innerHTML = economics.unitScenarios.map(scenario => {
        const salons = "salonsPerEmployee" in scenario
          ? nf.format(scenario.salonsPerEmployee)
          : rangeText(scenario.salonsPerEmployeeLow, scenario.salonsPerEmployeeHigh);
        return `
          <tr>
            <td>${esc(scenario.name)}</td>
            <td class="num"><b>${salons}</b></td>
            <td class="num">${moneyRange(scenario.priceLow, scenario.priceHigh)}</td>
            <td class="num">${moneyRange(scenario.revenueLow, scenario.revenueHigh)}</td>
            <td>${esc(scenario.requirement)}</td>
          </tr>
        `;
      }).join("");

      const sales = economics.salesCompensation;
      const salesSummary = [
        [money(sales.accountManagerSalary), "средний оклад аккаунт-менеджера"],
        [money(sales.bonusPerConnection), "разовая выплата за подключение: 5% от оклада"],
        [`${nf.format(sales.conservativeTotal)}–${nf.format(sales.targetTotal)}`, "салонов к концу 6-го месяца"],
        [moneyRange(sales.conservativeBonusTotal, sales.targetBonusTotal), "суммарная выплата за подключения за 6 месяцев"]
      ];
      $("salesCompSummary").innerHTML = salesSummary.map(([value, label]) => `
        <div class="mini-stat"><strong>${esc(value)}</strong><span>${esc(label)}</span></div>
      `).join("");
      $("salesPlanTable").innerHTML = sales.plan.map(row => `
        <tr>
          <td>${nf.format(row.month)}</td>
          <td class="num">${nf.format(row.conservative)}</td>
          <td class="num"><b>${nf.format(row.conservativeCumulative)}</b></td>
          <td class="num">${nf.format(row.target)}</td>
          <td class="num"><b>${nf.format(row.targetCumulative)}</b></td>
          <td class="num">${moneyRange(row.conservativeBonus, row.targetBonus)}</td>
          <td class="num">${moneyRange(row.conservativeRevenue, row.targetRevenue)}</td>
        </tr>
      `).join("");
    }

    init();
  </script>
</body>
</html>
"""

def script_json(value):
    return json.dumps(value, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")


def build_research_presentation(source_html):
    sections = re.findall(r'    <section class="slide(?: active)?">[\s\S]*?    </section>', source_html)
    if len(sections) < 9:
        return source_html

    title_slide = """    <section class="slide active">
      <div style="height:100%;display:grid;place-items:center;text-align:center;">
        <h1>анализ агентств</h1>
      </div>
    </section>"""

    pricing_slide = """    <section class="slide">
      <div class="eyebrow">Pricing snapshot</div>
      <h2>Тарифы и логика оплаты: сравнивать нужно состав услуги, а не только цену входа</h2>
      <table class="price-table">
        <thead><tr><th>Агентство</th><th>Ориентир по тарифу</th><th>Что входит / как читать цену</th></tr></thead>
        <tbody>
          <tr><td>CarCopy</td><td>400-800 ₽ за ТС; новые ТС - 5 900 ₽ за ТС; CRM - 200 000 ₽/мес за 15 аккаунтов</td><td>Модель ближе к SaaS/DMS: стоимость масштабируется от стока и подключенных модулей.</td></tr>
          <tr><td>4 Пикселя</td><td>Публично: классифайды от 28 000 ₽; в приложенном перечне тариф не раскрыт</td><td>Состав услуги широкий: фиды, публикация, склад, отчёты, контроль выгрузки 24/7 и персональный менеджер.</td></tr>
          <tr><td>ТАК / Tandem</td><td>Автопродвижение + постинг - 50 000 ₽; без продвижения - 8 000 / 9 000 / 10 000 ₽</td><td>Финальная цена зависит от договорённостей с дилером, объёма и того, нужна ли автоматизация продвижения.</td></tr>
        </tbody>
      </table>
      <p style="margin-top: 22px; max-width: 940px;">Вывод: у 4 Пикселей и Тандем цена должна сверяться с фактическим scope: только постинг, постинг + контроль, или продвижение под целевой объём лидов.</p>
      <div class="footer"><span>Тарифная логика</span><span>2 / 8</span></div>
    </section>"""

    four_pixels_slide = """    <section class="slide">
      <div class="eyebrow">Agency profile 02</div>
      <h2>4 Пикселя: ведение кабинетов как операционный контур дилера</h2>
      <div class="grid grid-2" style="margin-top: 30px;">
        <div class="card">
          <div class="tag">Фиды и публикация</div>
          <p class="label">Генерация фидов под форматы автомобильных площадок, кастомные фиды по параметрам авто, правила выбора площадок, автостратегии и планирование услуг продвижения.</p>
        </div>
        <div class="card">
          <div class="tag">Склад и контент</div>
          <p class="label">Добавление авто из API/XML/JSON/XLS, комплектации и опции, обогащение справочниками, видео/360, фотобанки, автоматические тексты объявлений.</p>
        </div>
        <div class="card">
          <div class="tag">Отчётность и аналитика</div>
          <p class="label">Расходы, просмотры, избранное, звонки, VIN-ошибки, заявки trade-in на Auto.ru, дни в стоке, уникальные авто на площадках, ежедневная аналитика и кастомные отчёты.</p>
        </div>
        <div class="card">
          <div class="tag">Контроль 24/7 и менеджер</div>
          <p class="label">Исправление ошибок выгрузки, возврат заблокированных объявлений, коммуникация с техподдержкой площадок, бюджетирование, аудиты, рекомендации и контроль акций.</p>
        </div>
      </div>
      <p style="margin-top: 22px; max-width: 1020px;">Агрегированный вывод: 4 Пикселя закрывают не только публикацию, но и ежедневную операционную дисциплину вокруг стока, качества объявлений, бюджетов и эффективности размещения.</p>
      <div class="footer"><span>4 Пикселя</span><span>5 / 8</span></div>
    </section>"""

    tandem_slide = """    <section class="slide">
      <div class="eyebrow">Agency profile 03</div>
      <h2>ТАК / Tandem: продвижение под нужный модел-микс и объём лидов</h2>
      <div class="split">
        <div class="card">
          <div class="tag">Ключевая модель</div>
          <p class="big-quote" style="margin-top: 22px;">Автоматизированное продвижение новых и б/у авто на классифайдах в соответствии с модел-миксом, который нужен дилеру для целевого количества лидов.</p>
        </div>
        <div class="stack">
          <div class="mini"><strong>Операционный контур</strong><span>Контроль выгрузки объявлений, обработка ошибок и блокировок, поддержание корректного статуса публикаций.</span></div>
          <div class="mini"><strong>Тариф с продвижением</strong><span>Автоматизированное продвижение + постинг: 50 000 ₽.</span></div>
          <div class="mini"><strong>Тариф без продвижения</strong><span>Постинг без продвижения: 8 000 / 9 000 / 10 000 ₽; финальные условия зависят от договорённостей с конкретным дилером.</span></div>
          <div class="mini"><strong>Когда выбирать</strong><span>Если нужна не просто выгрузка, а управление размещением под лиды, модел-микс и эффективность продвижения.</span></div>
        </div>
      </div>
      <div class="footer"><span>ТАК / Tandem Group</span><span>6 / 8</span></div>
    </section>"""

    sources_slide = """    <section class="slide">
      <div class="eyebrow">Sources & limits</div>
      <h2>Источники и ограничения</h2>
      <div class="card" style="margin-top: 30px;">
        <ul class="sources">
          <li><a href="https://carcopy.ru/" target="_blank">carcopy.ru</a> - описание платформы CarCopy.</li>
          <li><a href="https://carcopy.ru/tariffs" target="_blank">carcopy.ru/tariffs</a> - тарифы CarCopy по ТС и звонкам.</li>
          <li><a href="https://carcopy.ru/crm" target="_blank">carcopy.ru/crm</a> - CRM-пакеты CarCopy.</li>
          <li><a href="https://4px.ru/" target="_blank">4px.ru</a> и <a href="https://4px.ru/services/classifieds/" target="_blank">4px.ru/services/classifieds</a> - услуги 4 Пикселя по классифайдам.</li>
          <li><a href="https://tandemdigital.ru/o-programme-autocloud/" target="_blank">tandemdigital.ru/o-programme-autocloud</a> и <a href="https://tandemdigital.ru/uslugi/classifieds-reklama/" target="_blank">classifieds-реклама</a> - Autocloud и услуги Tandem.</li>
          <li>Приложенный файл «Услуги 4 Пикселеи.docx» - детализация услуг 4 Пикселей.</li>
          <li>Вводные по ТАК / Tandem от пользователя - модель продвижения, контроль выгрузок и тарифы.</li>
        </ul>
        <p style="margin-top: 16px;">Ограничение: публичные цены и переданные ориентиры могут не включать интеграции, дополнительные кабинеты, расширенную аналитику и индивидуальные условия.</p>
      </div>
      <div class="footer"><span>Проверить перед коммерческим сравнением</span><span>Appendix</span></div>
    </section>"""

    sources_slide = sources_slide.replace("<span>Appendix</span>", "<span>7 / 7</span>")

    kept = [
        pricing_slide,
        sections[3].replace("<span>4 / 8</span>", "<span>3 / 8</span>"),
        sections[4].replace("<span>5 / 8</span>", "<span>4 / 8</span>"),
        four_pixels_slide,
        tandem_slide,
        sources_slide,
    ]

    body = "\n\n".join([title_slide, *kept])
    source_html = re.sub(r"<title>.*?</title>", "<title>Анализ агентств</title>", source_html, count=1, flags=re.S)
    source_html = re.sub(
        r'  <main class="deck">\n[\s\S]*?\n  </main>',
        f"  <main class=\"deck\">\n{body}\n  </main>",
        source_html,
        count=1,
    )
    return source_html


payload = script_json(data)
presentation_html = build_research_presentation(PRESENTATION.read_text(encoding="utf-8"))
presentation_payload = script_json(presentation_html)
html = (
    html_template
    .replace("__DASHBOARD_DATA__", payload)
    .replace("__RESEARCH_PRESENTATION__", presentation_payload)
)
OUTPUT.write_text(html, encoding="utf-8")
print(OUTPUT)
