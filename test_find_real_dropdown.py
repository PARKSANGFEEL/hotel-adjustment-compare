#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Find the actual visible dropdown/select UI for the page size selector.
The native select is hidden in a visually-hidden LABEL, so the real UI must be elsewhere.
"""

import json
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Initialize Chrome driver
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

try:
    # Load cookies
    with open('expedia_cookies.json', 'r') as f:
        cookies = json.load(f)
    
    # Navigate to base URL first
    driver.get('https://partnercentral.expedia.com/')
    time.sleep(2)
    
    # Add cookies
    for cookie in cookies:
        try:
            driver.add_cookie(cookie)
        except:
            pass
    
    # Navigate to statements page
    driver.get('https://partnercentral.expedia.com/htid_17300293/financial/statements_and_invoices?tab=invoices')
    
    # Wait for JavaScript payload
    WebDriverWait(driver, 10).until(
        EC.presence_of_all_elements_located((By.CLASS_NAME, 'fds-table-row'))
    )
    time.sleep(2)
    
    print("[INFO] Statements page loaded")
    
    # === STRATEGY 1: Check if there's a visible select-like element in the DOM ===
    result = driver.execute_script("""
        // Find all select elements
        const selects = document.querySelectorAll('select');
        const visibleSelects = Array.from(selects).filter(s => {
            return s.offsetParent !== null; // offsetParent === null means hidden
        });
        
        console.log(`Total selects: ${selects.length}, Visible: ${visibleSelects.length}`);
        
        // Find all elements with "aria-label" containing "view", "show", "display", "per" (common select labels)
        const labeledElements = document.querySelectorAll('[aria-label*="view"], [aria-label*="show"], [aria-label*="display"], [aria-label*="per"]');
        console.log(`Found ${labeledElements.length} elements with view/show/display/per labels`);
        
        // Find all role="combobox" or role="listbox"
        const comboboxes = document.querySelectorAll('[role="combobox"], [role="listbox"]');
        console.log(`Found ${comboboxes.length} combobox/listbox elements`);
        
        return {
            visibleSelectCount: visibleSelects.length,
            comboboxCount: comboboxes.length,
            labeledElementCount: labeledElements.length
        };
    """)
    
    print(f"[INFO] DOM Search Results: {result}")
    
    # === STRATEGY 2: Look for buttons near the hidden select ===
    result2 = driver.execute_script("""
        // Find the hidden select
        const select = document.querySelector('select.fds-field-select');
        if (!select) return { error: 'select not found' };
        
        // Go up to container (likely the form or nearby container)
        let container = select.closest('div') || select.parentElement;
        while (container && !container.className.includes('container') && container.tagName !== 'BODY') {
            // Check if this container has visible children
            const visibleChildren = Array.from(container.children).filter(el => el.offsetParent !== null);
            if (visibleChildren.length > 0) {
                return {
                    containerTag: container.tagName,
                    containerClass: container.className,
                    visibleChildCount: visibleChildren.length,
                    visibleChildTags: visibleChildren.map(c => `${c.tagName}.${c.className}`).slice(0, 5)
                };
            }
            container = container.parentElement;
        }
        
        return { error: 'no visible container found' };
    """)
    
    print(f"[INFO] Container Search Results: {result2}")
    
    # === STRATEGY 3: Search for buttons with specific text or icons ===
    result3 = driver.execute_script("""
        // Look for any visible button or div that might be the dropdown trigger
        const allButtons = document.querySelectorAll('button[type]:not([type="hidden"]), div[role="button"], span[role="button"]');
        const visibleButtons = Array.from(allButtons).filter(b => b.offsetParent !== null);
        
        // Filter for those near the select area or with specific attributes
        const interestingButtons = visibleButtons.filter(b => {
            const text = b.textContent.toLowerCase();
            const className = (b.className || '').toLowerCase();
            // Looking for pagination, page size, rows per page indicators
            return text.includes('10') || text.includes('25') || text.includes('50') || text.includes('100') ||
                   className.includes('pagination') || className.includes('page') || className.includes('select');
        });
        
        return {
            totalVisibleButtons: visibleButtons.length,
            relevantButtons: interestingButtons.map(b => ({
                tag: b.tagName,
                class: b.className.substring(0, 100),
                text: b.textContent.substring(0, 50),
                offsetHeight: b.offsetHeight,
                offsetWidth: b.offsetWidth
            })).slice(0, 10)
        };
    """)
    
    print(f"[INFO] Button Search Results:")
    print(json.dumps(result3, indent=2, ensure_ascii=False))
    
    # === STRATEGY 4: Check the entire page structure near the table ===
    result4 = driver.execute_script("""
        // Find the table
        const table = document.querySelector('[role="table"], .fds-table, table');
        if (!table) return { error: 'table not found' };
        
        // Get parent container of the table
        const parent = table.parentElement;
        const siblings = Array.from(parent.parentElement.children);
        
        // Find visible siblings of the table's parent
        const visibleSiblings = siblings.filter(s => s.offsetParent !== null)
            .map(s => ({
                tag: s.tagName,
                class: s.className.substring(0, 80),
                hasSelect: !!s.querySelector('select'),
                hasButton: !!s.querySelector('button'),
                childCount: s.children.length,
                text: s.textContent.substring(0, 50)
            }));
        
        return {
            tableParentClass: parent.className,
            visibleSiblingsCount: visibleSiblings.length,
            siblings: visibleSiblings
        };
    """)
    
    print(f"[INFO] Table Context Results:")
    print(json.dumps(result4, indent=2, ensure_ascii=False))
    
    print("\n[INFO] Press Enter to close browser...")
    input()
    
finally:
    driver.quit()
