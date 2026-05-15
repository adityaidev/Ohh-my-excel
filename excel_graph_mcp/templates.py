TEMPLATES: dict[str, dict] = {
    "expense_tracker": {
        "name": "Expense Tracker",
        "category": "Finance",
        "description": "Track monthly expenses with categories, auto-totals, and summary dashboard",
        "sheets": [
            {"name": "Expenses", "columns": ["Date", "Category", "Description", "Amount", "Payment Method"]},
            {"name": "Summary", "columns": ["Category", "Total Spent", "% of Budget", "vs Last Month"]},
        ],
    },
    "budget_planner": {
        "name": "Budget Planner",
        "category": "Finance",
        "description": "Plan monthly budgets with variance tracking and alerts",
        "sheets": [
            {"name": "Budget", "columns": ["Category", "Budgeted", "Actual", "Difference", "Status"]},
            {"name": "Overview", "columns": ["Month", "Total Budgeted", "Total Spent", "Remaining"]},
        ],
    },
    "invoice_generator": {
        "name": "Invoice Generator",
        "category": "Finance",
        "description": "Create professional invoices with line items, tax, and totals",
        "sheets": [
            {"name": "Invoice", "columns": ["Item", "Quantity", "Unit Price", "Total", "Notes"]},
            {"name": "Settings", "columns": ["Company Name", "Tax Rate", "Currency", "Invoice Prefix"]},
        ],
    },
    "pnl_statement": {
        "name": "P&L Statement",
        "category": "Finance",
        "description": "Monthly profit and loss statement with YTD comparisons",
        "sheets": [
            {"name": "Income", "columns": ["Category", "Current Month", "Previous Month", "YTD", "Budget"]},
            {"name": "Expenses", "columns": ["Category", "Current Month", "Previous Month", "YTD", "Budget"]},
            {"name": "Summary", "columns": ["Metric", "Value", "vs Budget", "vs Last Year"]},
        ],
    },
    "cash_flow": {
        "name": "Cash Flow Statement",
        "category": "Finance",
        "description": "Track operating, investing, and financing cash flows",
        "sheets": [
            {"name": "Operating", "columns": ["Item", "Amount", "Notes"]},
            {"name": "Investing", "columns": ["Item", "Amount", "Notes"]},
            {"name": "Financing", "columns": ["Item", "Amount", "Notes"]},
            {"name": "Summary", "columns": ["Period", "Net Operating", "Net Investing", "Net Financing", "Net Change"]},
        ],
    },
    "balance_sheet": {
        "name": "Balance Sheet",
        "category": "Finance",
        "description": "Assets, liabilities, and equity with auto-calculated totals",
        "sheets": [
            {"name": "Assets", "columns": ["Category", "Current", "Long-term", "Total"]},
            {"name": "Liabilities", "columns": ["Category", "Current", "Long-term", "Total"]},
            {"name": "Equity", "columns": ["Category", "Amount", "Notes"]},
            {"name": "Summary", "columns": ["Section", "Total"]},
        ],
    },
    "sales_pipeline": {
        "name": "Sales Pipeline",
        "category": "Sales",
        "description": "Track deals through stages with probability weighting and forecast",
        "sheets": [
            {"name": "Pipeline", "columns": ["Deal Name", "Company", "Value", "Stage", "Probability", "Expected Close", "Owner"]},
            {"name": "Forecast", "columns": ["Quarter", "Pipeline Value", "Weighted Forecast", "Closed Won"]},
        ],
    },
    "crm_tracker": {
        "name": "CRM Tracker",
        "category": "Sales",
        "description": "Customer relationship management with contact history",
        "sheets": [
            {"name": "Contacts", "columns": ["Name", "Company", "Email", "Phone", "Last Contact", "Status", "Notes"]},
            {"name": "Activities", "columns": ["Date", "Type", "Contact", "Notes", "Follow-up"]},
        ],
    },
    "lead_scoring": {
        "name": "Lead Scoring",
        "category": "Sales",
        "description": "Score leads based on engagement, fit, and intent signals",
        "sheets": [
            {"name": "Leads", "columns": ["Lead Name", "Company", "Score", "Engagement", "Fit", "Intent", "Priority"]},
        ],
    },
    "revenue_forecast": {
        "name": "Revenue Forecast",
        "category": "Sales",
        "description": "Monthly/quarterly revenue projections with scenario analysis",
        "sheets": [
            {"name": "Forecast", "columns": ["Month", "Conservative", "Expected", "Optimistic", "Actual"]},
            {"name": "Assumptions", "columns": ["Parameter", "Conservative", "Expected", "Optimistic"]},
        ],
    },
    "commission_calc": {
        "name": "Commission Calculator",
        "category": "Sales",
        "description": "Calculate sales commissions with tiered rates and targets",
        "sheets": [
            {"name": "Deals", "columns": ["Rep", "Deal Value", "Commission Rate", "Commission", "Target Met"]},
            {"name": "Rates", "columns": ["Tier", "Min", "Max", "Rate"]},
        ],
    },
    "project_timeline": {
        "name": "Project Timeline",
        "category": "Project",
        "description": "Track project milestones, tasks, and dependencies",
        "sheets": [
            {"name": "Tasks", "columns": ["Task", "Owner", "Start", "End", "Status", "Priority", "Depends On"]},
            {"name": "Milestones", "columns": ["Milestone", "Date", "Deliverable", "Status"]},
        ],
    },
    "gantt_chart": {
        "name": "Gantt Chart",
        "category": "Project",
        "description": "Visual project schedule with timeline bars",
        "sheets": [
            {"name": "Schedule", "columns": ["Task", "Start", "End", "Duration", "% Complete", "Assigned To"]},
        ],
    },
    "sprint_planner": {
        "name": "Sprint Planner",
        "category": "Project",
        "description": "Agile sprint planning with story points and velocity tracking",
        "sheets": [
            {"name": "Backlog", "columns": ["Story", "Points", "Priority", "Status", "Assignee"]},
            {"name": "Sprint", "columns": ["Story", "Points", "Status", "Assignee", "Hours Spent"]},
            {"name": "Velocity", "columns": ["Sprint", "Committed", "Completed", "Velocity"]},
        ],
    },
    "resource_allocation": {
        "name": "Resource Allocation",
        "category": "Project",
        "description": "Track team resource allocation across projects",
        "sheets": [
            {"name": "Resources", "columns": ["Name", "Role", "Project", "Allocation %", "Start", "End"]},
            {"name": "Summary", "columns": ["Project", "Total FTEs", "Budget", "Actual"]},
        ],
    },
    "risk_register": {
        "name": "Risk Register",
        "category": "Project",
        "description": "Identify, assess, and track project risks",
        "sheets": [
            {"name": "Risks", "columns": ["Risk", "Category", "Probability", "Impact", "Score", "Mitigation", "Owner", "Status"]},
        ],
    },
    "employee_directory": {
        "name": "Employee Directory",
        "category": "HR",
        "description": "Company employee directory with contact information",
        "sheets": [
            {"name": "Employees", "columns": ["Name", "Department", "Title", "Email", "Phone", "Location", "Manager"]},
        ],
    },
    "attendance_tracker": {
        "name": "Attendance Tracker",
        "category": "HR",
        "description": "Track employee attendance, leave, and absences",
        "sheets": [
            {"name": "Attendance", "columns": ["Employee", "Date", "Type", "Status", "Hours", "Notes"]},
            {"name": "Summary", "columns": ["Employee", "Present", "Absent", "Leave", "Late"]},
        ],
    },
    "performance_review": {
        "name": "Performance Review",
        "category": "HR",
        "description": "Employee performance evaluations with scoring",
        "sheets": [
            {"name": "Review", "columns": ["Employee", "Reviewer", "Category", "Score", "Comments", "Date"]},
            {"name": "Goals", "columns": ["Employee", "Goal", "Target Date", "Status", "Notes"]},
        ],
    },
    "leave_calendar": {
        "name": "Leave Calendar",
        "category": "HR",
        "description": "Team leave and vacation calendar",
        "sheets": [
            {"name": "Leaves", "columns": ["Employee", "Type", "Start", "End", "Days", "Status", "Approver"]},
        ],
    },
    "campaign_tracker": {
        "name": "Campaign Tracker",
        "category": "Marketing",
        "description": "Track marketing campaigns with KPIs and ROI",
        "sheets": [
            {"name": "Campaigns",             "columns": ["Campaign", "Channel", "Budget", "Spent", "Impressions", "Clicks", "Conversions", "ROI"],
        },
        ],
    },
    "content_calendar": {
        "name": "Content Calendar",
        "category": "Marketing",
        "description": "Plan and schedule content across channels",
        "sheets": [
            {"name": "Content", "columns": ["Title", "Channel", "Type", "Author", "Publish Date", "Status", "Notes"]},
        ],
    },
    "social_media_planner": {
        "name": "Social Media Planner",
        "category": "Marketing",
        "description": "Plan and track social media posts across platforms",
        "sheets": [
            {"name": "Posts", "columns": ["Platform", "Date", "Content", "Image", "Link", "Status", "Engagement"]},
        ],
    },
    "roi_calculator": {
        "name": "ROI Calculator",
        "category": "Marketing",
        "description": "Calculate return on investment for marketing activities",
        "sheets": [
            {"name": "Investment", "columns": ["Activity", "Cost", "Date", "Category"]},
            {"name": "Returns", "columns": ["Activity", "Revenue", "Leads", "Conversions", "Date"]},
            {"name": "ROI", "columns": ["Activity", "Total Cost", "Total Return", "ROI %", "Payback Period"]},
        ],
    },
    "habit_tracker": {
        "name": "Habit Tracker",
        "category": "Personal",
        "description": "Track daily habits with streaks and progress",
        "sheets": [
            {"name": "Habits", "columns": ["Habit", "Category", "Frequency", "Current Streak", "Best Streak", "Total Days"]},
            {"name": "Log", "columns": ["Date", "Habit", "Completed", "Notes"]},
        ],
    },
    "meal_planner": {
        "name": "Meal Planner",
        "category": "Personal",
        "description": "Plan weekly meals with grocery lists",
        "sheets": [
            {"name": "Meals", "columns": ["Day", "Breakfast", "Lunch", "Dinner", "Snacks", "Calories"]},
            {"name": "Grocery List", "columns": ["Item", "Category", "Quantity", "Purchased"]},
        ],
    },
    "workout_log": {
        "name": "Workout Log",
        "category": "Personal",
        "description": "Track workouts, sets, reps, and progress",
        "sheets": [
            {"name": "Workouts", "columns": ["Date", "Exercise", "Sets", "Reps", "Weight", "Duration", "Notes"]},
        ],
    },
    "travel_budget": {
        "name": "Travel Budget",
        "category": "Personal",
        "description": "Plan and track travel expenses",
        "sheets": [
            {"name": "Budget", "columns": ["Category", "Estimated", "Actual", "Difference", "Notes"]},
            {"name": "Itinerary", "columns": ["Date", "Time", "Activity", "Location", "Cost", "Notes"]},
        ],
    },
    "grade_calculator": {
        "name": "Grade Calculator",
        "category": "Personal",
        "description": "Calculate grades with weighted assignments",
        "sheets": [
            {"name": "Grades", "columns": ["Assignment", "Category", "Weight %", "Score", "Max Score", "Weighted Score"]},
            {"name": "Summary", "columns": ["Category", "Average", "Weight", "Contribution"]},
        ],
    },
}


def list_templates(category: str = "") -> list[dict]:
    result = []
    for name, tmpl in TEMPLATES.items():
        if category and tmpl["category"].lower() != category.lower():
            continue
        result.append({
            "id": name,
            "name": tmpl["name"],
            "category": tmpl["category"],
            "description": tmpl["description"],
            "sheets": len(tmpl["sheets"]),
        })
    return result


def get_template(template_name: str) -> dict:
    tmpl = TEMPLATES.get(template_name)
    if not tmpl:
        return {"error": f"Template '{template_name}' not found"}
    return tmpl


def get_template_categories() -> list[dict]:
    cats = {}
    for name, tmpl in TEMPLATES.items():
        cat = tmpl["category"]
        if cat not in cats:
            cats[cat] = {"name": cat, "count": 0, "templates": []}
        cats[cat]["count"] += 1
        cats[cat]["templates"].append(name)
    return list(cats.values())
