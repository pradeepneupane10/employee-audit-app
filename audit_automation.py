# /// script
# dependencies = [
#     "playwright>=1.40.0",
#     "pandas>=2.0.0",
#     "openpyxl>=3.1.0",
#     "python-dateutil>=2.8.2",
# ]
# ///

import os
import re
import sys
import argparse
from datetime import datetime, timedelta
import pandas as pd
from dateutil import parser as date_parser
from playwright.sync_api import sync_playwright

# Credentials
LOGIN_USER = "laxman.koirala"
LOGIN_PASS = "koirala...laxman"

# Reconfigure stdout/stderr to utf-8 if supported on Windows
try:
    if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    if sys.stderr and hasattr(sys.stderr, 'reconfigure'):
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# Setup simple, clean logging
def log(msg, level="INFO"):
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted_msg = f"[{timestamp}] [{level}] {msg}"
    try:
        print(formatted_msg, flush=True)
    except UnicodeEncodeError:
        safe_msg = formatted_msg.encode('ascii', errors='replace').decode('ascii')
        print(safe_msg, flush=True)

def wait_for_postback(page, timeout_ms=8000):
    try:
        page.evaluate("""
        () => {
            return new Promise((resolve) => {
                if (typeof Sys === 'undefined' || !Sys.WebForms || !Sys.WebForms.PageRequestManager) {
                    resolve();
                    return;
                }
                const prm = Sys.WebForms.PageRequestManager.getInstance();
                if (!prm.get_isInAsyncPostBack()) {
                    resolve();
                    return;
                }
                const handler = () => {
                    if (!prm.get_isInAsyncPostBack()) {
                        prm.remove_endRequest(handler);
                        resolve();
                    }
                };
                prm.add_endRequest(handler);
                setTimeout(() => {
                    prm.remove_endRequest(handler);
                    resolve();
                }, 7500);
            });
        }
        """)
    except Exception:
        page.wait_for_timeout(500)

def safe_wait_for_networkidle(page, timeout_ms=5000):
    try:
        page.wait_for_load_state("networkidle", timeout=timeout_ms)
    except Exception:
        pass

def safe_eval(page, script, arg=None, max_retries=5):
    """Executes page.evaluate safely, retrying if an ASP.NET postback destroyed the execution context."""
    for attempt in range(max_retries):
        try:
            if arg is not None:
                return page.evaluate(script, arg)
            return page.evaluate(script)
        except Exception as e:
            err_str = str(e).lower()
            if "execution context was destroyed" in err_str or "navigation" in err_str or "target closed" in err_str:
                page.wait_for_timeout(700)
                try:
                    page.wait_for_load_state("domcontentloaded", timeout=5000)
                except Exception:
                    pass
                continue
            if attempt == max_retries - 1:
                log(f"safe_eval error after {max_retries} attempts: {e}", "WARNING")
                return None
            page.wait_for_timeout(300)
    return None

def format_duration(td):
    if pd.isna(td) or not isinstance(td, timedelta):
        return "N/A"
    total_seconds = int(td.total_seconds())
    if total_seconds < 0:
        return "Negative Time"
    days = total_seconds // 86400
    hours = (total_seconds % 86400) // 3600
    minutes = (total_seconds % 3600) // 60
    
    parts = []
    if days > 0:
        parts.append(f"{days}d")
    if hours > 0:
        parts.append(f"{hours}h")
    if minutes > 0 or not parts:
        parts.append(f"{minutes}m")
    return " ".join(parts)

def parse_date_robustly(date_str):
    if not date_str or pd.isna(date_str):
        return None
    # Normalize whitespaces
    date_str = re.sub(r'\s+', ' ', str(date_str).strip())
    if not date_str:
        return None
    
    # Try dateutil parser first as it is very smart
    try:
        return date_parser.parse(date_str)
    except Exception:
        pass
    
    # Fallback to standard formats
    formats = [
        "%d %b %Y %I:%M:%S %p",  # 05 Aug 2026 05:04:12 PM
        "%d %b %Y %I:%M %p",     # 05 Aug 2026 05:04 PM
        "%d-%b-%Y %I:%M:%S %p",
        "%d-%b-%Y %I:%M %p",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %I:%M:%S %p",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(date_str, fmt)
        except ValueError:
            continue
    return None

def handle_login(page):
    try:
        page.locator('input[type="password"]').first.wait_for(state="visible", timeout=6000)
    except Exception as e:
        log(f"Password field wait failed or timed out: {e}. Checking if logged in.")
        # Take a screenshot to help diagnose page state
        try:
            page.screenshot(path="login_error_debug.png")
            log("Saved debug screenshot to login_error_debug.png")
        except Exception as screenshot_err:
            log(f"Failed to capture login screenshot: {screenshot_err}")
        return
        
    log("Login page detected. Performing automatic login...")
    
    # Find username input
    username_field = None
    for selector in ['input[id*="user" i]', 'input[name*="user" i]', 'input[placeholder*="user" i]', 'input[type="text"]']:
        loc = page.locator(selector)
        if loc.count() > 0:
            username_field = loc.first
            break
            
    # Find password input
    password_field = page.locator('input[type="password"]').first
    
    # Fill details
    if username_field:
        username_field.fill(LOGIN_USER)
    password_field.fill(LOGIN_PASS)
    
    # Find submit button
    submit_btn = None
    for selector in ['input[type="submit"]', 'button', 'input[value*="Login" i]', 'input[value*="Log In" i]', 'input[id*="btn" i]']:
        loc = page.locator(selector)
        if loc.count() > 0:
            submit_btn = loc.first
            break
            
    if submit_btn:
        submit_btn.click()
        page.wait_for_load_state("networkidle")
        page.wait_for_timeout(1500)
        log("Logged in successfully.")
    else:
        log("Submit button not found!", "ERROR")


automate_selects_js = r"""
([auditFor, moduleVal, operationVal, employeeVal]) => {
    const selects = Array.from(document.querySelectorAll('select'));
    
    // 1. Audit For
    const auditForSelect = document.getElementById('ContentPlaceHolder1_ddlOF') || selects.find(s => 
        Array.from(s.options).some(o => o.text.trim().toLowerCase() === auditFor.toLowerCase() || o.value.trim().toLowerCase() === auditFor.toLowerCase())
    );
    if (auditForSelect) {
        const opt = Array.from(auditForSelect.options).find(o => o.text.trim().toLowerCase() === auditFor.toLowerCase() || o.value.trim().toLowerCase() === auditFor.toLowerCase());
        if (opt && auditForSelect.value !== opt.value) {
            auditForSelect.value = opt.value;
            auditForSelect.dispatchEvent(new Event('change', { bubbles: true }));
            return "auditFor";
        }
    }
    
    // 2. Module
    const moduleSelect = document.getElementById('ContentPlaceHolder1_combomodul') || selects.find(s => 
        Array.from(s.options).some(o => o.text.trim().toLowerCase() === moduleVal.toLowerCase() || o.value.trim().toLowerCase() === moduleVal.toLowerCase())
    );
    if (moduleSelect) {
        const opt = Array.from(moduleSelect.options).find(o => o.text.trim().toLowerCase() === moduleVal.toLowerCase() || o.value.trim().toLowerCase() === moduleVal.toLowerCase());
        if (opt && moduleSelect.value !== opt.value) {
            moduleSelect.value = opt.value;
            moduleSelect.dispatchEvent(new Event('change', { bubbles: true }));
            return "module";
        }
    }
    
    // 3. Operation
    const operationSelect = document.getElementById('ContentPlaceHolder1_comboactivity') || selects.find(s => 
        Array.from(s.options).some(o => o.text.trim().toLowerCase() === operationVal.toLowerCase() || o.value.trim().toLowerCase() === operationVal.toLowerCase())
    );
    if (operationSelect) {
        const opt = Array.from(operationSelect.options).find(o => o.text.trim().toLowerCase() === operationVal.toLowerCase() || o.value.trim().toLowerCase() === operationVal.toLowerCase());
        if (opt && operationSelect.value !== opt.value) {
            operationSelect.value = opt.value;
            operationSelect.dispatchEvent(new Event('change', { bubbles: true }));
            return "operation";
        }
    }
    
    // 4. Employee (Target specific combouserddl dropdown with exact name matching)
    const cleanTarget = employeeVal.toLowerCase().replace(/\s+/g, ' ').trim();
    const employeeSelect = document.getElementById('ContentPlaceHolder1_combouserddl') || selects.find(s => 
        s.id.includes('user') || Array.from(s.options).some(o => o.text.toLowerCase().replace(/\s+/g, ' ').trim() === cleanTarget)
    );
    if (employeeSelect) {
        // 1. Exact match first
        let opt = Array.from(employeeSelect.options).find(o => o.text.toLowerCase().replace(/\s+/g, ' ').trim() === cleanTarget);
        
        // 2. Strict multi-word matching: ALL words (first name AND last name) must be present in the option
        if (!opt) {
            const targetWords = cleanTarget.split(' ').filter(w => w.length > 1);
            if (targetWords.length > 1) {
                opt = Array.from(employeeSelect.options).find(o => {
                    const optText = o.text.toLowerCase().replace(/\s+/g, ' ').trim();
                    return targetWords.every(w => optText.includes(w));
                });
            }
        }
        
        if (opt && employeeSelect.value !== opt.value) {
            employeeSelect.value = opt.value;
            employeeSelect.dispatchEvent(new Event('change', { bubbles: true }));
            return "employee";
        }
    }
    
    // 5. Select DateWise radio button if present
    const dateWiseRadio = document.getElementById('ContentPlaceHolder1_rdbdat_0') || 
                          document.querySelector('input[type="radio"][value*="Date" i], input[type="radio"][id*="dat" i]');
    if (dateWiseRadio && !dateWiseRadio.checked) {
        dateWiseRadio.click();
        dateWiseRadio.checked = true;
        dateWiseRadio.dispatchEvent(new Event('change', { bubbles: true }));
    }

    // 6. Page Size dropdown - select 500 Records / highest batch size
    const pageSizeSelect = document.getElementById('ContentPlaceHolder1_ddlPageSize') || selects.find(s => 
        Array.from(s.options).some(o => o.text.includes('Records') || o.text.includes('500') || o.text.includes('200'))
    );
    if (pageSizeSelect) {
        const optMax = Array.from(pageSizeSelect.options).find(o => o.text.includes('500') || o.text.includes('200') || o.text.includes('100'));
        if (optMax && pageSizeSelect.value !== optMax.value) {
            pageSizeSelect.value = optMax.value;
            pageSizeSelect.dispatchEvent(new Event('change', { bubbles: true }));
            return "pageSize";
        }
    }
    
    return "done";
}
"""

set_date_inputs_js = r"""
([fromDateStr, toDateStr]) => {
    const fromInput = document.getElementById('ContentPlaceHolder1_txtfrmdate') || document.querySelector('input[id*="frm" i], input[name*="frm" i]');
    const toInput = document.getElementById('ContentPlaceHolder1_txttodate') || document.querySelector('input[id*="to" i], input[name*="to" i]');
    
    if (fromInput) {
        fromInput.value = fromDateStr;
        fromInput.dispatchEvent(new Event('input', { bubbles: true }));
        fromInput.dispatchEvent(new Event('change', { bubbles: true }));
        fromInput.dispatchEvent(new Event('blur', { bubbles: true }));
    }
    if (toInput) {
        toInput.value = toDateStr;
        toInput.dispatchEvent(new Event('input', { bubbles: true }));
        toInput.dispatchEvent(new Event('change', { bubbles: true }));
        toInput.dispatchEvent(new Event('blur', { bubbles: true }));
    }
    return !!(fromInput && toInput);
}
"""

click_search_js = r"""
() => {
    const btn = document.getElementById('ContentPlaceHolder1_btnserch');
    if (btn) {
        btn.click();
        return true;
    }
    const buttons = Array.from(document.querySelectorAll('input[type="submit"], input[type="button"], button'));
    const searchBtn = buttons.find(b => {
        const val = (b.value || b.textContent || '').trim().toLowerCase();
        return val === 'search' || val.includes('search');
    });
    if (searchBtn) {
        searchBtn.click();
        return true;
    }
    return false;
}
"""

# Extraction loop injection scripts
find_grid_table_js = r"""
() => {
    const grid = document.getElementById('ContentPlaceHolder1_gdhistory') || 
                 document.querySelector('table[id*="gdhistory" i]') || 
                 document.querySelector('table.table2');
    return !!grid;
}
"""

get_rows_count_js = r"""
() => {
    const gridTable = document.getElementById('ContentPlaceHolder1_gdhistory') || 
                      document.querySelector('table[id*="gdhistory" i]') || 
                      document.querySelector('table.table2');
    if (!gridTable) return 0;
    const rows = Array.from(gridTable.querySelectorAll('tr'));
    let count = 0;
    for (const row of rows) {
        const cells = Array.from(row.querySelectorAll('td'));
        if (cells.length < 5) continue;
        const firstCellText = cells[0].textContent.trim().toLowerCase();
        if (firstCellText === 'date') continue;
        if (row.querySelector('table') || cells.some(c => c.textContent.trim().match(/^\d+$/) && cells.length <= 2)) {
            continue;
        }
        count++;
    }
    return count;
}
"""

get_row_data_js = r"""
(rowIndex) => {
    const gridTable = document.getElementById('ContentPlaceHolder1_gdhistory') || 
                      document.querySelector('table[id*="gdhistory" i]') || 
                      document.querySelector('table.table2');
    if (!gridTable) return null;
    const rows = Array.from(gridTable.querySelectorAll('tr'));
    
    const dataRows = [];
    for (const row of rows) {
        const cells = Array.from(row.querySelectorAll('td'));
        if (cells.length < 5) continue;
        const firstCellText = cells[0].textContent.trim().toLowerCase();
        if (firstCellText === 'date') continue;
        if (row.querySelector('table') || cells.some(c => c.textContent.trim().match(/^\d+$/) && cells.length <= 2)) {
            continue;
        }
        dataRows.push(row);
    }
    
    if (rowIndex >= dataRows.length) return null;
    const targetRow = dataRows[rowIndex];
    const cells = Array.from(targetRow.querySelectorAll('td'));
    
    return {
        date: cells[0].textContent.trim(),
        userName: cells[1].textContent.trim(),
        employeeName: cells[2].textContent.trim(),
        moduleName: cells[3].textContent.trim(),
        operation: cells[4].textContent.trim(),
        ipAddress: cells.length > 5 ? cells[5].textContent.trim() : '',
        remark: cells.length > 6 ? cells[6].textContent.trim() : ''
    };
}
"""

click_row_date_js = click_row_link_js = r"""
(rowIndex) => {
    const gridTable = document.getElementById('ContentPlaceHolder1_gdhistory') || 
                      document.querySelector('table[id*="gdhistory" i]') || 
                      document.querySelector('table.table2');
    if (!gridTable) return false;
    const rows = Array.from(gridTable.querySelectorAll('tr'));
    
    const dataRows = [];
    for (const row of rows) {
        const cells = Array.from(row.querySelectorAll('td'));
        if (cells.length < 5) continue;
        const firstCellText = cells[0].textContent.trim().toLowerCase();
        if (firstCellText === 'date') continue;
        if (row.querySelector('table') || cells.some(c => c.textContent.trim().match(/^\d+$/) && cells.length <= 2)) {
            continue;
        }
        dataRows.push(row);
    }
    
    if (rowIndex >= dataRows.length) return false;
    const cell = dataRows[rowIndex].querySelectorAll('td')[0];
    const clickTarget = cell.querySelector('a') || cell;
    clickTarget.click();
    return true;
}
"""

scroll_modal_js = r"""
() => {
    const headers = Array.from(document.querySelectorAll('*')).filter(el => 
        el.textContent && el.textContent.trim() === 'Case Information' && el.offsetWidth > 0
    );
    if (headers.length === 0) return false;
    
    let container = null;
    let parent = headers[0].parentElement;
    while (parent && parent.tagName !== 'BODY') {
        if (parent.querySelector('input[type="button"][value="Close"]') || parent.innerText.includes('Ticket Number')) {
            container = parent;
            break;
        }
        parent = parent.parentElement;
    }
    
    if (!container) container = document;
    
    let scrolled = false;
    const allElements = container.querySelectorAll('*');
    for (const el of allElements) {
        if (el.scrollHeight > el.clientHeight && 
            (window.getComputedStyle(el).overflowY === 'auto' || 
             window.getComputedStyle(el).overflowY === 'scroll' ||
             el.style.overflow === 'auto' ||
             el.style.overflow === 'scroll')) {
            el.scrollTop = el.scrollHeight;
            scrolled = true;
        }
    }
    return scrolled;
}
"""

extract_modal_data_js = r"""
() => {
    const targets = [
        "Title", "Ticket Number", "Account Name", "Status", "Case Origin", 
        "Customer Name", "Mobile No", "Category", "Sub Sub Category", 
        "Assigned Date", "Estimated closed time based on SLA", 
        "Estimated closed time based on TAT", "User Id", "Created Date", 
        "Last Modified Date", "Associated Location Status", "Latitude", 
        "Case Reason", "Assign Team", "Assign User", "Priority", "Type", 
        "Customer Address", "Email", "Sub Category", "Escalated Layer", 
        "SLA", "SLA Left/Total", "TAT Left/Total", "Created By", 
        "Last Modified By", "Associated Location Level", 
        "Associated Location Name", "Longitude"
    ];
    
    const headers = Array.from(document.querySelectorAll('*')).filter(el => 
        el.textContent && el.textContent.trim() === 'Case Information' && el.offsetWidth > 0
    );
    
    let container = null;
    if (headers.length > 0) {
        let parent = headers[0].parentElement;
        while (parent && parent.tagName !== 'BODY') {
            if (parent.querySelector('input[type="button"][value="Close"]') || parent.innerText.includes('Ticket Number')) {
                container = parent;
                break;
            }
            parent = parent.parentElement;
        }
    }
    if (!container) container = document;
    
    const data = {};
    const cells = Array.from(container.querySelectorAll('td'));
    
    for (const target of targets) {
        const foundCell = cells.find(c => {
            const text = c.textContent ? c.textContent.trim().replace(/\s+/g, ' ') : '';
            return text === target || text === target + ':' || text === target + ' :';
        });
        
        if (foundCell) {
            const nextCell = foundCell.nextElementSibling;
            if (nextCell && nextCell.tagName === 'TD') {
                data[target] = nextCell.textContent.trim();
            } else {
                const row = foundCell.closest('tr');
                if (row) {
                    const rowCells = Array.from(row.querySelectorAll('td'));
                    const idx = rowCells.indexOf(foundCell);
                    if (idx !== -1 && idx + 1 < rowCells.length) {
                        data[target] = rowCells[idx + 1].textContent.trim();
                    }
                }
            }
        } else {
            const divs = Array.from(container.querySelectorAll('div, span, label'));
            const foundDiv = divs.find(d => {
                const text = d.textContent ? d.textContent.trim().replace(/\s+/g, ' ') : '';
                return (text === target || text === target + ':') && d.children.length === 0 && d.offsetWidth > 0;
            });
            if (foundDiv) {
                let next = foundDiv.nextElementSibling;
                if (next) {
                    data[target] = next.textContent.trim();
                }
            }
        }
    }

    // Extract Solution Given / Work Log / Employee Remarks from modal text
    let solutionGivenText = "";
    let fullText = (container.innerText || container.textContent || "").trim();
    
    const solutionMatch = fullText.match(/Solution Given[\s\S]*/i);
    if (solutionMatch) {
        solutionGivenText = solutionMatch[0].replace(/^Solution Given\s*/i, '').replace(/Close\s*$/i, '').trim();
    }
    
    data["Solution Given"] = solutionGivenText || "";
    data["Employee Remark / Solution Note"] = solutionGivenText || fullText;
    return data;
}
"""

close_modal_js = r"""
() => {
    const buttons = Array.from(document.querySelectorAll('input[type="button"], button, a')).filter(el => 
        el.offsetWidth > 0
    );
    const closeBtn = buttons.find(el => {
        const val = (el.value || el.textContent || '').trim().toLowerCase();
        return val === 'close';
    });
    if (closeBtn) {
        closeBtn.click();
        return true;
    }
    return false;
}
"""

get_pagination_info_js = r"""
() => {
    const gridTable = document.getElementById('ContentPlaceHolder1_gdhistory') || 
                      document.querySelector('table[id*="gdhistory" i]') || 
                      document.querySelector('table.table2');
    if (!gridTable) return null;
    
    let links = Array.from(gridTable.querySelectorAll('a[id*="lnbPg"], span[id*="lnbPg"]'));
    if (links.length === 0) {
        const rows = Array.from(gridTable.querySelectorAll('tr'));
        if (rows.length > 0) {
            const lastRow = rows[rows.length - 1];
            const candidateLinks = Array.from(lastRow.querySelectorAll('a, span'));
            links = candidateLinks.filter(el => {
                const txt = el.textContent.trim();
                return /^\d+$/.test(txt) || txt === '...';
            });
        }
    }
    
    if (links.length <= 1) return null;

    return links.map(el => {
        const text = el.textContent.trim();
        const hasHref = el.hasAttribute('href') && el.getAttribute('href').length > 0;
        const isDisabled = el.classList.contains('aspNetDisabled') || !hasHref;
        return {
            text: text,
            active: isDisabled,
            clickable: !isDisabled && hasHref
        };
    }).filter(item => /^\d+$/.test(item.text) || item.text === '...');
}
"""

click_page_js = r"""
(pageNumStr) => {
    const gridTable = document.getElementById('ContentPlaceHolder1_gdhistory') || 
                      document.querySelector('table[id*="gdhistory" i]') || 
                      document.querySelector('table.table2');
    if (!gridTable) return false;
    
    const links = Array.from(gridTable.querySelectorAll('a'));
    
    // 1. Check for exact page number
    let targetLink = links.find(el => {
        const txt = el.textContent.trim();
        return txt === String(pageNumStr) && el.hasAttribute('href') && !el.classList.contains('aspNetDisabled');
    });
    
    // 2. If page number not directly visible, click the trailing ellipsis (Next set of 10 pages)
    if (!targetLink) {
        const ellipsisLinks = links.filter(el => el.textContent.trim() === '...' && el.hasAttribute('href') && !el.classList.contains('aspNetDisabled'));
        if (ellipsisLinks.length > 0) {
            targetLink = ellipsisLinks[ellipsisLinks.length - 1];
        }
    }
    
    if (targetLink) {
        const href = targetLink.getAttribute('href') || '';
        if (href.toLowerCase().startsWith('javascript:')) {
            try {
                window.eval(href.substring(11));
            } catch(e) {
                targetLink.click();
            }
        } else {
            targetLink.click();
        }
        return true;
    }
    return false;
}
"""


def scrape_single_employee(target_emp, from_date, to_date, launch_kwargs, known_keys=None):
    """
    Worker function executed in parallel threads or sequentially.
    Each worker has its own dedicated Playwright browser instance and session.
    If known_keys is provided, it stops scraping early the moment it encounters
    a record that was already processed in a previous run.
    """
    emp_records = []
    with sync_playwright() as p:
        log(f"[{target_emp}] 🚀 Launching dedicated browser context...")
        browser = p.chromium.launch(**launch_kwargs)
        context = browser.new_context(viewport={"width": 1280, "height": 720}, ignore_https_errors=True)
        page = context.new_page()
        
        # Abort heavy assets (images, fonts, media) to dramatically accelerate page loads
        page.route("**/*.{png,jpg,jpeg,svg,gif,webp,woff,woff2,ttf,eot,ico}", lambda route: route.abort())
        
        login_url = "https://billing.cgnet.com.np/h8ssrms/Login.aspx"
        target_url = "https://billing.cgnet.com.np/h8ssrms/Auditpage.aspx"
        
        log(f"[{target_emp}] Connecting to CGNET login portal...")
        try:
            page.goto(login_url, timeout=60000)
            safe_wait_for_networkidle(page, 10000)
        except Exception as e:
            log(f"[{target_emp}] Navigation issue: {e}", "WARNING")
            
        handle_login(page)
        
        log(f"[{target_emp}] Loading Audit page...")
        try:
            page.goto(target_url, timeout=30000)
            safe_wait_for_networkidle(page, 10000)
        except Exception as e:
            log(f"[{target_emp}] Could not load audit page: {e}", "ERROR")
            browser.close()
            return []

        # Auto-configure grid dropdowns
        log(f"[{target_emp}] Auto-applying dropdown filters (AuditFor='Employee', Module='Case', Operation='Update')...")
        for attempt in range(10):
            step = safe_eval(page, automate_selects_js, ["Employee", "Case", "Update", target_emp])
            if step == "done":
                break
            wait_for_postback(page, timeout_ms=3000)
            page.wait_for_timeout(150)
            safe_wait_for_networkidle(page, 2000)

        active_selected_text = safe_eval(page, """() => {
            const s = document.getElementById('ContentPlaceHolder1_combouserddl');
            return s ? (s.options[s.selectedIndex] ? s.options[s.selectedIndex].text : '') : '';
        }""")
        log(f"[{target_emp}] Active dropdown: '{active_selected_text}'")

        # HARD SAFETY GUARD: verify active dropdown matches target employee
        active_clean = str(active_selected_text or '').strip().lower()
        target_words = [w for w in target_emp.strip().lower().split() if len(w) > 1]
        is_match = all(w in active_clean for w in target_words) if target_words else (target_emp.lower() in active_clean)
        if not is_match:
            log(f"[{target_emp}] ⚠️ ERROR: Dropdown shows '{active_selected_text}', which does NOT match target '{target_emp}'. Aborting scrape for this employee to prevent corrupting data!", "ERROR")
            browser.close()
            return []

        # Set Dates
        log(f"[{target_emp}] Setting date inputs (From: {from_date}, To: {to_date})...")
        dates_set = safe_eval(page, set_date_inputs_js, [from_date, to_date])
        page.wait_for_timeout(100)

        # Click search
        log(f"[{target_emp}] Submitting query...")
        searched = safe_eval(page, click_search_js)
        if searched:
            wait_for_postback(page, timeout_ms=5000)
            page.wait_for_timeout(400)
            safe_wait_for_networkidle(page, 4000)
        else:
            page.wait_for_timeout(800)

        try:
            page.locator('table').first.wait_for(state="visible", timeout=4000)
        except Exception:
            pass

        page_num = 1
        stop_employee_early = False
        while True:
            log(f"[{target_emp}] Scanning page {page_num}...")
            has_grid = safe_eval(page, find_grid_table_js)
            if not has_grid:
                log(f"[{target_emp}] No audit records grid table found on page {page_num}.")
                break
                
            rows_count = safe_eval(page, get_rows_count_js) or 0
            log(f"[{target_emp}] Page {page_num}: Found {rows_count} records.")
            
            if rows_count == 0:
                break
                
            for i in range(rows_count):
                row_info = safe_eval(page, get_row_data_js, i)
                if not row_info:
                    continue
                
                # ⚡ Incremental Early-Stopping Check:
                # If this record was already scraped in a previous run, stop immediately!
                row_key = f"{target_emp.strip().lower()}|{str(row_info.get('date', '')).strip()}|{str(row_info.get('remark', '')).strip()}"
                if known_keys and row_key in known_keys:
                    log(f"[{target_emp}] ⚡ Reached previously scraped record ('{row_info.get('date')}'). Stopping early for {target_emp}!")
                    stop_employee_early = True
                    break
                
                log(f"[{target_emp}] Record {i+1}/{rows_count} (Page {page_num}): Date='{row_info['date']}' | User='{row_info['userName']}' | Remark='{row_info['remark']}'")
                
                # Click row link to open modal
                clicked = safe_eval(page, click_row_link_js, i)
                if not clicked:
                    continue
                
                wait_for_postback(page, timeout_ms=4000)
                page.wait_for_timeout(50)
                
                # Robust check for modal frame
                modal_frame = None
                for _ in range(40):
                    direct_f = page.frame(name="ContentPlaceHolder1_ifrm")
                    if direct_f and direct_f.url != "about:blank":
                        try:
                            body_len = direct_f.evaluate("() => document.body ? document.body.innerText.length : 0")
                            if body_len > 20:
                                modal_frame = direct_f
                                break
                        except Exception:
                            pass
                    candidate_frames = [direct_f] if direct_f else page.frames
                    for frame in candidate_frames:
                        if not frame or frame.url == "about:blank":
                            continue
                        try:
                            has_header = frame.evaluate(r"""
                            () => {
                                const headers = Array.from(document.querySelectorAll('*')).filter(el => 
                                    el.textContent && el.textContent.trim().toLowerCase().includes('case information') && el.offsetWidth > 0
                                );
                                return headers.length > 0;
                            }
                            """)
                            if has_header:
                                modal_frame = frame
                                break
                        except Exception:
                            pass
                    if modal_frame:
                        break
                    page.wait_for_timeout(50)

                modal_data = {}
                if modal_frame:
                    try:
                        modal_frame.evaluate(scroll_modal_js)
                    except Exception:
                        pass
                    page.wait_for_timeout(30)
                    try:
                        modal_data = modal_frame.evaluate(extract_modal_data_js) or {}
                    except Exception:
                        modal_data = {}
                    
                    # Close modal swiftly
                    closed = False
                    try:
                        closed = modal_frame.evaluate(close_modal_js)
                    except Exception:
                        pass
                    if not closed:
                        try:
                            modal_frame.locator('text=Close').first.click(timeout=800)
                        except Exception:
                            pass
                else:
                    # Fallback if modal iframe didn't open: close via iframe click
                    try:
                        frame = page.frame(name="ContentPlaceHolder1_ifrm")
                        if frame:
                            frame.evaluate(close_modal_js)
                    except Exception:
                        pass
                
                wait_for_postback(page, timeout_ms=3000)
                for _ in range(20):
                    try:
                        if not page.locator('iframe[name="ContentPlaceHolder1_ifrm"]').is_visible():
                            break
                    except Exception:
                        break
                    page.wait_for_timeout(30)

                # Fallback: Extract Ticket Number from remark if missing in modal_data
                if not modal_data.get("Ticket Number"):
                    tkt_match = re.search(r'TKT\d+', row_info["remark"], re.I)
                    if tkt_match:
                        modal_data["Ticket Number"] = tkt_match.group(0).upper()

                combined_record = {
                    "Target Employee": target_emp,
                    "Grid Date": row_info["date"],
                    "Grid User Name": row_info["userName"],
                    "Grid Employee Name": row_info["employeeName"],
                    "Grid Module": row_info["moduleName"],
                    "Grid Operation": row_info["operation"],
                    "Grid IP Address": row_info["ipAddress"],
                    "Grid Remark": row_info["remark"],
                    **modal_data
                }
                emp_records.append(combined_record)
                
            if stop_employee_early:
                break
                
            pagination_items = page.evaluate(get_pagination_info_js)
            if not pagination_items:
                break
                
            active_item = next((item for item in pagination_items if item["active"]), None)
            if not active_item:
                break
                
            current_page_val = int(active_item["text"]) if active_item["text"].isdigit() else page_num
            next_page_val = current_page_val + 1
            has_next_number = any(item["text"] == str(next_page_val) and item["clickable"] for item in pagination_items)
            has_next_ellipsis = any(item["text"] == "..." and item["clickable"] for item in pagination_items)
            
            if has_next_number or has_next_ellipsis:
                first_row_before = page.evaluate(get_row_data_js, 0)
                date_before = first_row_before["date"] if first_row_before else ""
                remark_before = first_row_before["remark"] if first_row_before else ""
                
                page.evaluate(click_page_js, str(next_page_val))
                wait_for_postback(page)
                page.wait_for_timeout(400)
                safe_wait_for_networkidle(page, 4000)
                
                # Wait up to 5 seconds for page content to genuinely switch
                page_switched = False
                for _ in range(25):
                    post_pagination = page.evaluate(get_pagination_info_js)
                    new_active = next((item for item in post_pagination if item["active"]), None) if post_pagination else None
                    first_row_after = page.evaluate(get_row_data_js, 0)
                    date_after = first_row_after["date"] if first_row_after else ""
                    remark_after = first_row_after["remark"] if first_row_after else ""
                    
                    if (new_active and new_active["text"].isdigit() and int(new_active["text"]) >= next_page_val) or (date_before != date_after or remark_before != remark_after):
                        page_num = int(new_active["text"]) if (new_active and new_active["text"].isdigit()) else next_page_val
                        page_switched = True
                        break
                    page.wait_for_timeout(200)
                    
                if not page_switched:
                    break
            else:
                break

        log(f"[{target_emp}] ✨ Finished scraping! Total records: {len(emp_records)}")
        browser.close()
    return emp_records


def main():
    print("=" * 80)
    print("            CGNET AUTOMATED AUDIT SCRAPER & REPORT GENERATOR")
    print("=" * 80)
    
    # Parse CLI Arguments
    today_str = datetime.now().strftime("%d %b %Y")  # e.g., "05 Aug 2026"
    parser = argparse.ArgumentParser(description="CGNET Employee Audit Scraper")
    parser.add_argument("--employee", default="Om Neupane", help="Employee name (default: Om Neupane)")
    parser.add_argument("--from-date", default=today_str, help=f"From Date, e.g. '05 Aug 2026' (default: {today_str})")
    parser.add_argument("--to-date", default=today_str, help=f"To Date, e.g. '05 Aug 2026' (default: {today_str})")
    parser.add_argument("--non-interactive", action="store_true", help="Run in full automated mode without prompt")
    parser.add_argument("--force-full-scrape", action="store_true", help="Force full scrape without using incremental cache")
    args = parser.parse_args()

    employee_name = args.employee
    from_date = args.from_date
    to_date = args.to_date

    # Interactive Override
    if not args.non_interactive:
        print("Please configure the run details (Press Enter to use defaults):")
        emp_input = input(f"Employee Name [{employee_name}]: ").strip()
        if emp_input:
            employee_name = emp_input
            
        from_input = input(f"From Date [{from_date}]: ").strip()
        if from_input:
            from_date = from_input
            
        to_input = input(f"To Date [{to_date}]: ").strip()
        if to_input:
            to_date = to_input
            
    log(f"Configuration set: Employee={employee_name}, From={from_date}, To={to_date}")
    
    # Setup folders
    output_dir = os.path.dirname(os.path.abspath(__file__))
    # Format a safe filename with employee and date
    if ',' in employee_name or 'ALL TEAM' in employee_name.upper():
        safe_emp = "ALL_TEAM"
    else:
        safe_emp = re.sub(r'[^a-zA-Z0-9]', '_', employee_name)
    safe_from = re.sub(r'[^a-zA-Z0-9]', '_', from_date)
    output_file = os.path.join(output_dir, f"audit_report_{safe_emp}_{safe_from}.xlsx")

    # ⚡ Load Existing Records for Incremental / Delta Scraping
    existing_records = []
    known_keys = set()
    if os.path.exists(output_file) and not args.force_full_scrape:
        try:
            existing_df = pd.read_excel(output_file, sheet_name="Audit Details")
            if not existing_df.empty:
                existing_records = existing_df.to_dict(orient="records")
                for r in existing_records:
                    emp_k = str(r.get("Target Employee", r.get("Employee Name", ""))).strip().lower()
                    date_k = str(r.get("Grid Date", "")).strip()
                    rem_k = str(r.get("Grid Remark", "")).strip()
                    if emp_k and (date_k or rem_k):
                        known_keys.add(f"{emp_k}|{date_k}|{rem_k}")
                log(f"⚡ Incremental Mode Active: Loaded {len(existing_records)} existing records ({len(known_keys)} unique keys) from {os.path.basename(output_file)}.")
        except Exception as read_ex:
            log(f"Could not load existing file for incremental check: {read_ex}. Running full scrape.", "WARNING")
            existing_records = []
            known_keys = set()

    # Prepare list of target employees
    if ',' in employee_name:
        emp_list = [e.strip() for e in employee_name.split(',') if e.strip()]
    elif employee_name.strip().upper() in ["ALL", "ALL TEAM", "TEAM"]:
        emp_list = [
            "Ajit Shrestha", "Chandramani Tharu", "Om Neupane", "Rajesh Maharjan",
            "Sabin Giri", "Sanjeev Giri", "Shashikant Chaudhary", "Sunil Chaudhary"
        ]
    else:
        emp_list = [employee_name.strip()]

    is_headless = os.environ.get("HEADLESS", "true").lower() != "false"
    chrome_path = None
    for path in ["/usr/bin/chromium", "/usr/bin/chromium-browser", "/usr/bin/google-chrome"]:
        if os.path.exists(path):
            chrome_path = path
            break
            
    launch_kwargs = {
        "headless": is_headless,
        "args": ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
    }
    if chrome_path:
        log(f"Using system chromium at {chrome_path}")
        launch_kwargs["executable_path"] = chrome_path

    newly_scraped = []

    if len(emp_list) == 1:
        log(f"Starting audit scrape for: {emp_list[0]}")
        newly_scraped = scrape_single_employee(emp_list[0], from_date, to_date, launch_kwargs, known_keys=known_keys)
    else:
        max_workers = min(8, len(emp_list))
        log(f"⚡ Launching {max_workers} HIGH-SPEED PARALLEL WORKERS to scrape {len(emp_list)} employees simultaneously...")
        from concurrent.futures import ThreadPoolExecutor, as_completed
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            future_to_emp = {
                executor.submit(scrape_single_employee, emp, from_date, to_date, launch_kwargs, known_keys=known_keys): emp
                for emp in emp_list
            }
            for future in as_completed(future_to_emp):
                emp_done = future_to_emp[future]
                try:
                    emp_recs = future.result()
                    newly_scraped.extend(emp_recs)
                    log(f"✅ Finished [{emp_done}]: collected {len(emp_recs)} new records.")
                except Exception as exc:
                    log(f"❌ Error during scrape for [{emp_done}]: {exc}", "ERROR")

    # Combine existing records with newly scraped records safely
    if existing_records:
        records_dict = {}
        for r in existing_records:
            emp_k = str(r.get("Target Employee", r.get("Employee Name", ""))).strip().lower()
            date_k = str(r.get("Grid Date", "")).strip()
            rem_k = str(r.get("Grid Remark", "")).strip()
            key = f"{emp_k}|{date_k}|{rem_k}"
            records_dict[key] = r
            
        for r in newly_scraped:
            emp_k = str(r.get("Target Employee", r.get("Employee Name", ""))).strip().lower()
            date_k = str(r.get("Grid Date", "")).strip()
            rem_k = str(r.get("Grid Remark", "")).strip()
            key = f"{emp_k}|{date_k}|{rem_k}"
            records_dict[key] = r
            
        scraped_records = list(records_dict.values())
        log(f"✨ Total combined records after incremental update: {len(scraped_records)} ({len(newly_scraped)} newly added).")
    else:
        scraped_records = newly_scraped

    # --- Post-Processing & Report Generation ---
    if not scraped_records:
        log("No records were scraped. Unable to generate report.", "ERROR")
        return
        
    log(f"Scraped {len(scraped_records)} total audit records. Compiling report...")
        
    df = pd.DataFrame(scraped_records)
    
    required_cols = ['Ticket Number', 'Status', 'Category', 'Sub Category', 'Assigned Date', 'Created Date', 'Last Modified Date']
    for col in required_cols:
        if col not in df.columns:
            df[col] = None
            
    df['dt_assigned'] = df['Assigned Date'].apply(parse_date_robustly)
    df['dt_created'] = df['Created Date'].apply(parse_date_robustly)
    df['dt_completed'] = df['Last Modified Date'].apply(parse_date_robustly)
    
    df['duration_assigned_sec'] = (df['dt_completed'] - df['dt_assigned']).dt.total_seconds()
    df['duration_created_sec'] = (df['dt_completed'] - df['dt_created']).dt.total_seconds()
    
    df['duration_assigned'] = pd.to_timedelta(df['duration_assigned_sec'], unit='s')
    df['duration_created'] = pd.to_timedelta(df['duration_created_sec'], unit='s')
    df['Resolution Time (From Assigned)'] = df['duration_assigned'].apply(format_duration)
    df['Resolution Time (From Created)'] = df['duration_created'].apply(format_duration)
    
    df['Status_Cleaned'] = df['Status'].astype(str).str.strip().str.capitalize()
    
    def classify_ticket_handling(row):
        status_clean = str(row.get('Status', '')).strip().capitalize()
        raw_status = str(row.get('Status', '')).strip()
        grid_remark = str(row.get('Grid Remark', '')).strip()
        assign_team = str(row.get('Assign Team', '')).strip()
        assign_user = str(row.get('Assign User', '')).strip()
        
        remark_lower = grid_remark.lower()
        team_lower = assign_team.lower()
        user_lower = assign_user.lower()
        
        is_ms_or_transferred = (
            'ms' in team_lower or 'ms' in user_lower or
            'ms' in remark_lower or 'assign' in remark_lower or
            'transfer' in remark_lower or 'forward' in remark_lower or
            'handover' in remark_lower or 'handler' in remark_lower
        )
        
        if status_clean in ['Completed', 'Closed']:
            if is_ms_or_transferred and status_clean != 'Completed':
                res_type = "Assigned / Transferred to MS / Other Handler"
            else:
                res_type = f"Completed ({raw_status if raw_status else 'Completed'})"
            return True, res_type, grid_remark
        elif is_ms_or_transferred:
            res_type = "Assigned / Transferred to MS / Other Handler"
            return True, res_type, grid_remark
        else:
            res_type = f"Other / In Progress ({raw_status if raw_status else 'Unspecified'})"
            return False, res_type, grid_remark

    res_results = df.apply(classify_ticket_handling, axis=1)
    df['Is_Solved_Or_Assigned'] = [r[0] for r in res_results]
    df['Resolution Type'] = [r[1] for r in res_results]
    df['Action / Transfer Remark'] = [r[2] for r in res_results]
    
    def classify_work_type(row):
        remark = str(row.get('Grid Remark', '')).strip()
        solution = str(row.get('Solution Given', '')).strip()
        employee_note = str(row.get('Employee Remark / Solution Note', '')).strip()
        title = str(row.get('Title', '')).strip()
        category = str(row.get('Category', '')).strip()
        
        full_text = f"{title} {remark} {solution} {employee_note}".upper()
        
        # Robust WiFi 6 detection (Title, Category, Sub Category, Remark, Solution Note, Serial Number)
        wifi6_regex = r'wifi\s*6|wifi-6|wifi6|wi-fi\s*6|router\s*6|alcl|nokia.*wifi|upgrade.*wifi|wifi.*upgrade|dual\s*band'
        is_wifi6_upgrade = (
            bool(re.search(wifi6_regex, full_text, re.I)) or
            'WIFI 6' in full_text or 'WIFI-6' in full_text or 'WIFI6' in full_text or 'ALCL' in full_text
        )
        
        if is_wifi6_upgrade:
            return 'WiFi 6 Upgrade'
        elif 'IPTV' in full_text or 'IPTV' in category.upper():
            return 'IPTV Issue'
        else:
            return 'General / Other Issues'

    df['Task / Issue Type'] = df.apply(classify_work_type, axis=1)
    
    # Ensure Employee Remark / Solution Note column is clean
    if 'Employee Remark / Solution Note' not in df.columns:
        df['Employee Remark / Solution Note'] = ""
        
    df['Employee Remark / Solution Note'] = df.apply(
        lambda r: str(r.get('Solution Given', '')).strip() or str(r.get('Employee Remark / Solution Note', '')).strip() or str(r.get('Grid Remark', '')).strip(),
        axis=1
    )
    
    solved_mask = df['Is_Solved_Or_Assigned']
    
    total_records = len(df)
    solved_tickets = df[solved_mask]
    num_solved = len(solved_tickets)
    
    avg_assigned_duration = solved_tickets['duration_assigned'].mean()
    avg_created_duration = solved_tickets['duration_created'].mean()
    
    # 1. Category Breakdown for Solved Tickets
    cat_group = solved_tickets.groupby(['Category', 'Sub Category'], dropna=False).agg(
        Count=('Ticket Number', 'count'),
        Avg_Duration_Assigned=('duration_assigned_sec', 'mean'),
        Avg_Duration_Created=('duration_created_sec', 'mean')
    ).reset_index()
    cat_group['Avg Resolution Time (From Assigned)'] = pd.to_timedelta(cat_group['Avg_Duration_Assigned'], unit='s').apply(format_duration)
    cat_group['Avg Resolution Time (From Created)'] = pd.to_timedelta(cat_group['Avg_Duration_Created'], unit='s').apply(format_duration)
    cat_group.drop(columns=['Avg_Duration_Assigned', 'Avg_Duration_Created'], inplace=True)
    
    # 2. TOTAL Tickets Category Breakdown (All Tickets Scraped)
    cat_total_group = df.groupby(['Category', 'Sub Category'], dropna=False).agg(
        Total_Tickets=('Ticket Number', 'count'),
        Solved_Count=('Is_Solved_Or_Assigned', 'sum')
    ).reset_index()
    cat_total_group['Solution Rate %'] = (cat_total_group['Solved_Count'] / cat_total_group['Total_Tickets'] * 100).round(1).astype(str) + '%'
    cat_total_group.rename(columns={
        'Total_Tickets': 'Total Scraped Tickets',
        'Solved_Count': 'Solved / Handled Count'
    }, inplace=True)
    
    # 3. High-level Category Summary
    cat_summary_overall = df.groupby('Category', dropna=False).agg(
        Total_Tickets=('Ticket Number', 'count'),
        Solved_Count=('Is_Solved_Or_Assigned', 'sum')
    ).reset_index()
    cat_summary_overall['Category Solution Rate %'] = (cat_summary_overall['Solved_Count'] / cat_summary_overall['Total_Tickets'] * 100).round(1).astype(str) + '%'
    cat_summary_overall.rename(columns={
        'Total_Tickets': 'Total Scraped Tickets',
        'Solved_Count': 'Solved / Handled Count'
    }, inplace=True)
    
    solved_list_cols = [
        'Ticket Number', 'Account Name', 'Category', 'Sub Category', 
        'Status', 'Resolution Type', 'Employee Remark / Solution Note', 'Action / Transfer Remark',
        'Assigned Date', 'Created Date', 'Last Modified Date',
        'Resolution Time (From Assigned)', 'Resolution Time (From Created)'
    ]
    solved_list_cols = [c for c in solved_list_cols if c in solved_tickets.columns]
    solved_list = solved_tickets[solved_list_cols].copy()
    
    summary_metrics = pd.DataFrame({
        "Metric": [
            "Total Audit Records Scraped",
            "Total Tickets Solved / Handled (Completed or MS Assigned)",
            "Ticket Solution Rate",
            "Average Completion Time (From Assigned Date)",
            "Average Completion Time (From Created Date)"
        ],
        "Value": [
            total_records,
            num_solved,
            f"{(num_solved / total_records * 100):.1f}%" if total_records > 0 else "0.0%",
            format_duration(avg_assigned_duration),
            format_duration(avg_created_duration)
        ]
    })
    
    df_export = df.drop(columns=[
        'dt_assigned', 'dt_created', 'dt_completed', 
        'duration_assigned_sec', 'duration_created_sec',
        'duration_assigned', 'duration_created', 'Status_Cleaned',
        'Is_Solved_Or_Assigned'
    ], errors='ignore')
    
    log(f"Writing Excel report to {output_file}...")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_export.to_excel(writer, sheet_name="Audit Details", index=False)
        summary_metrics.to_excel(writer, sheet_name="Summary Report", index=False, startrow=0, startcol=0)
        
        start_row_overall = len(summary_metrics) + 3
        pd.DataFrame([["Overall Category Summary"]]).to_excel(writer, sheet_name="Summary Report", index=False, header=False, startrow=start_row_overall-1, startcol=0)
        cat_summary_overall.to_excel(writer, sheet_name="Summary Report", index=False, startrow=start_row_overall, startcol=0)
        
        start_row_total_cat = start_row_overall + len(cat_summary_overall) + 3
        pd.DataFrame([["Total Tickets Breakdown by Category & Sub Category"]]).to_excel(writer, sheet_name="Summary Report", index=False, header=False, startrow=start_row_total_cat-1, startcol=0)
        cat_total_group.to_excel(writer, sheet_name="Summary Report", index=False, startrow=start_row_total_cat, startcol=0)

        start_row_solved = start_row_total_cat + len(cat_total_group) + 3
        pd.DataFrame([["Detailed Solved / Handled Tickets List"]]).to_excel(writer, sheet_name="Summary Report", index=False, header=False, startrow=start_row_solved-1, startcol=0)
        solved_list.to_excel(writer, sheet_name="Summary Report", index=False, startrow=start_row_solved, startcol=0)

    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    
    wb = openpyxl.load_workbook(output_file)
    
    ws_details = wb["Audit Details"]
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    cell_font = Font(name="Segoe UI", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    for col in range(1, ws_details.max_column + 1):
        cell = ws_details.cell(row=1, column=col)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    for row in range(2, ws_details.max_row + 1):
        for col in range(1, ws_details.max_column + 1):
            cell = ws_details.cell(row=row, column=col)
            cell.font = cell_font
            cell.border = thin_border
            
    for col in ws_details.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws_details.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 40)
        
    ws_summary = wb["Summary Report"]
    
    def style_summary_range(start_r, end_r, start_c, end_c, is_header=False):
        fill = PatternFill(start_color="2F5597" if is_header else "F2F2F2", end_color="2F5597" if is_header else "F2F2F2", fill_type="solid")
        font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF" if is_header else "000000")
        for r in range(start_r, end_r + 1):
            for c in range(start_c, end_c + 1):
                cell = ws_summary.cell(row=r, column=c)
                if is_header:
                    cell.fill = fill
                    cell.font = font
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.font = Font(name="Segoe UI", size=10)
                    cell.border = thin_border
                    if r % 2 == 1:
                        cell.fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")

    style_summary_range(1, 1, 1, 2, is_header=True)
    style_summary_range(2, len(summary_metrics)+1, 1, 2, is_header=False)
    
    section_title_font = Font(name="Segoe UI", size=12, bold=True, color="1F4E78")
    ws_summary.cell(row=start_row_overall, column=1).font = section_title_font
    ws_summary.cell(row=start_row_total_cat, column=1).font = section_title_font
    ws_summary.cell(row=start_row_solved, column=1).font = section_title_font
    
    style_summary_range(start_row_overall+1, start_row_overall+1, 1, len(cat_summary_overall.columns), is_header=True)
    style_summary_range(start_row_overall+2, start_row_overall+1+len(cat_summary_overall), 1, len(cat_summary_overall.columns), is_header=False)
    
    style_summary_range(start_row_total_cat+1, start_row_total_cat+1, 1, len(cat_total_group.columns), is_header=True)
    style_summary_range(start_row_total_cat+2, start_row_total_cat+1+len(cat_total_group), 1, len(cat_total_group.columns), is_header=False)
    
    style_summary_range(start_row_solved+1, start_row_solved+1, 1, len(solved_list.columns), is_header=True)
    style_summary_range(start_row_solved+2, start_row_solved+1+len(solved_list), 1, len(solved_list.columns), is_header=False)
    
    for col in ws_summary.columns:
        max_len = 0
        for cell in col:
            val = str(cell.value or '')
            if val in ["Overall Category Summary", "Total Tickets Breakdown by Category & Sub Category", "Detailed Solved / Handled Tickets List"]:
                continue
            max_len = max(max_len, len(val))
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws_summary.column_dimensions[col_letter].width = min(max(max_len + 4, 12), 45)
        
    wb.save(output_file)
    
    # --- Print Console Summary Report ---
    print("\n" + "=" * 80)
    print("                        DAILY PERFORMANCE SUMMARY")
    print("=" * 80)
    for _, row in summary_metrics.iterrows():
        print(f"{row['Metric']:<50} : {row['Value']}")
    print("-" * 80)
    print("\nBreakdown by Ticket Category:")
    print(f"{'Category':<25} | {'Sub Category':<30} | {'Total Count':<12} | {'Solved Count'}")
    print("-" * 80)
    for _, row in cat_total_group.iterrows():
        cat = str(row['Category'])[:23]
        sub = str(row['Sub Category'])[:28]
        tot = row['Total Scraped Tickets']
        cnt = row['Solved / Handled Count']
        print(f"{cat:<25} | {sub:<30} | {tot:<12} | {cnt}")
    print("=" * 80)
    log(f"Scraped details and compiled report saved to: {output_file}")
    print("=" * 80)

    # Generate Executive Team Excel report if multiple employees or ALL TEAM
    if len(emp_list) > 1 or "ALL" in employee_name.upper():
        exec_output_file = os.path.join(output_dir, f"EXECUTIVE_TEAM_AUDIT_REPORT_{safe_from}.xlsx")
        log(f"Building Combined Executive Team Report: {exec_output_file}...")
        try:
            emp_col = 'Grid Employee Name' if 'Grid Employee Name' in df.columns else 'Target Employee'
            df['Employee Name'] = df[emp_col].fillna(df.get('Target Employee', 'Unknown'))
            
            def is_solved_exec(row):
                st_val = str(row.get("Status", "")).strip().lower()
                rm_val = str(row.get("Grid Remark", "")).strip().lower()
                if st_val in ["completed", "closed"]:
                    return True
                if "ms" in rm_val or "assign" in rm_val or "transfer" in rm_val or "forward" in rm_val:
                    return True
                return False

            df['Is_Solved_Val'] = df.apply(is_solved_exec, axis=1)

            DEFAULT_TEAM_MEMBERS = [
                "Ajit Shrestha", "Chandramani Tharu", "Om Neupane", "Rajesh Maharjan",
                "Sabin Giri", "Sanjeev Giri", "Shashikant Chaudhary", "Sunil Chaudhary"
            ]

            summary_rows = []
            all_exec_emps = sorted(list(set([e for e in df["Employee Name"].dropna().unique().tolist() if e])))
            for emp in all_exec_emps:
                grp = df[df["Employee Name"] == emp]
                tot = len(grp)
                if tot == 0:
                    continue  # Exclude day-off technicians (0 tickets)
                solved = grp["Is_Solved_Val"].sum()
                rate = f"{(solved / tot * 100):.1f}%" if tot > 0 else "0.0%"
                summary_rows.append({
                    "Employee Name": emp,
                    "Total Scraped Tickets": tot,
                    "Solved / Handled Count": solved,
                    "Solution Rate %": rate
                })
                
            team_summary_df = pd.DataFrame(summary_rows)
            tot_tickets_team = len(df)
            tot_solved_team = df["Is_Solved_Val"].sum()
            team_rate_val = f"{(tot_solved_team / tot_tickets_team * 100):.1f}%" if tot_tickets_team > 0 else "0.0%"
            
            total_team_row = pd.DataFrame([{
                "Employee Name": "GRAND TOTAL (ALL TEAM)",
                "Total Scraped Tickets": tot_tickets_team,
                "Solved / Handled Count": tot_solved_team,
                "Solution Rate %": team_rate_val
            }])
            team_summary_df = pd.concat([team_summary_df, total_team_row], ignore_index=True)

            work_summary_df = df.groupby("Task / Issue Type", dropna=False).agg(
                Total_Tickets=("Ticket Number", "count"),
                Solved_Count=("Is_Solved_Val", "sum")
            ).reset_index()
            work_summary_df["Solution Rate %"] = (work_summary_df["Solved_Count"] / work_summary_df["Total_Tickets"] * 100).round(1).astype(str) + '%'
            work_summary_df["% Share of Total"] = (work_summary_df["Total_Tickets"] / len(df) * 100).round(1).astype(str) + '%'

            export_master = df_export.copy()
            if 'Employee Name' not in export_master.columns:
                export_master['Employee Name'] = df['Employee Name']

            matrix_cat_col = 'Sub Category' if 'Sub Category' in df.columns else 'Task / Issue Type'
            df[matrix_cat_col] = df[matrix_cat_col].fillna('Other / Uncategorized')
            df_dedup_p = df.drop_duplicates(subset=['Ticket Number'], keep='first') if 'Ticket Number' in df.columns else df
            matrix_pivot = pd.crosstab(
                df_dedup_p[matrix_cat_col],
                df_dedup_p['Employee Name'],
                margins=True,
                margins_name="Grand Total"
            )
            emp_cols_p = sorted([c for c in matrix_pivot.columns if c != "Grand Total"])
            matrix_pivot = matrix_pivot[emp_cols_p + (["Grand Total"] if "Grand Total" in matrix_pivot.columns else [])]
            if "Grand Total" in matrix_pivot.index:
                d_rows_p = matrix_pivot.drop(index="Grand Total").sort_values(by="Grand Total", ascending=False)
                t_row_p = matrix_pivot.loc[["Grand Total"]]
                matrix_pivot = pd.concat([d_rows_p, t_row_p])

            with pd.ExcelWriter(exec_output_file, engine='openpyxl') as exec_writer:
                team_summary_df.to_excel(exec_writer, sheet_name="Team Executive Summary", index=False)
                work_summary_df.to_excel(exec_writer, sheet_name="Category Summary", index=False)
                matrix_pivot.to_excel(exec_writer, sheet_name="Technician Matrix Pivot")
                export_master.to_excel(exec_writer, sheet_name="Master Audit Details", index=False)

            wb_exec = openpyxl.load_workbook(exec_output_file)
            for sheetname in wb_exec.sheetnames:
                ws = wb_exec[sheetname]
                for col in range(1, ws.max_column + 1):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                for row in range(2, ws.max_row + 1):
                    for col in range(1, ws.max_column + 1):
                        cell = ws.cell(row=row, column=col)
                        cell.font = cell_font
                        cell.border = thin_border
                        if (sheetname == "Team Executive Summary" or sheetname == "Technician Matrix Pivot") and row == ws.max_row:
                            cell.font = Font(name="Segoe UI", size=11, bold=True, color="1F4E78")
                            cell.fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
                        elif sheetname == "Technician Matrix Pivot" and col == ws.max_column:
                            cell.font = Font(name="Segoe UI", size=10, bold=True, color="1F4E78")
                            cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
                for col in ws.columns:
                    max_len = max(len(str(cell.value or '')) for cell in col)
                    col_letter = openpyxl.utils.get_column_letter(col[0].column)
                    ws.column_dimensions[col_letter].width = min(max(max_len + 3, 12), 45)
            wb_exec.save(exec_output_file)
            log(f"Combined Executive Team report saved to: {exec_output_file}")
        except Exception as exec_err:
            log(f"Could not build executive report: {exec_err}", "WARNING")

    # --- Update Master Historical Database (master_audit_history.csv) ---
    try:
        master_csv_path = os.path.join(output_dir, "master_audit_history.csv")
        df_to_append = df_export.copy()
        df_to_append['Report Date'] = from_date
        if os.path.exists(master_csv_path):
            existing_master = pd.read_csv(master_csv_path)
            combined_master = pd.concat([existing_master, df_to_append], ignore_index=True)
        else:
            combined_master = df_to_append
        
        subset_cols = [c for c in ['Ticket Number', 'Grid Date', 'Grid Remark', 'Grid Operation'] if c in combined_master.columns]
        if subset_cols:
            combined_master = combined_master.drop_duplicates(subset=subset_cols, keep='last')
        combined_master.to_csv(master_csv_path, index=False, encoding='utf-8')
        log(f"Updated master historical records database ({len(combined_master)} total records): {master_csv_path}")
    except Exception as hist_err:
        log(f"Could not update master history CSV: {hist_err}", "WARNING")

if __name__ == "__main__":
    main()
