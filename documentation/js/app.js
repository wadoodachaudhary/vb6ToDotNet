document.addEventListener('DOMContentLoaded', () => {
  const sidebarNav = document.getElementById('sidebar-nav');
  const contentContainer = document.getElementById('content-container');
  const searchInput = document.getElementById('search-input');
  const searchResults = document.getElementById('search-results');
  const themeToggleBtn = document.getElementById('theme-toggle-btn');
  const themeIcon = document.getElementById('theme-icon');
  const menuToggleBtn = document.getElementById('menu-toggle');
  const sidebar = document.querySelector('aside');
  const breadcrumbCategory = document.getElementById('breadcrumb-category');
  const breadcrumbPage = document.getElementById('breadcrumb-page');
  
  // Edit & Export DOM Elements
  const editPageBtn = document.getElementById('edit-page-btn');
  const exportDataBtn = document.getElementById('export-data-btn');
  const exportModal = document.getElementById('export-modal');
  const modalCloseBtn = document.getElementById('modal-close-btn');
  const exportTextarea = document.getElementById('export-textarea');
  const copyExportBtn = document.getElementById('copy-export-btn');
  const downloadExportBtn = document.getElementById('download-export-btn');

  let currentDocId = 'overview';
  let isEditing = false;

  // Deep copy of original data for resetting pages
  const ORIGINAL_DOCS = JSON.parse(JSON.stringify(DOCS_DATA));

  // Navigation Group Config
  const NAV_STRUCTURE = [
    {
      group: "Get Started",
      items: ["overview"]
    },
    {
      group: "Desktop Clients (VB6)",
      items: ["hfsystem", "hfest", "hfpayables"]
    },
    {
      group: "Modern Web Platform",
      items: ["blazor", "database"]
    },
    {
      group: "Workflows & Status",
      items: ["workflow", "migration"]
    }
  ];

  // Load local edits from localStorage
  const loadLocalEdits = () => {
    DOCS_DATA.forEach(doc => {
      const saved = localStorage.getItem('docs_edits_' + doc.id);
      if (saved) {
        try {
          const parsed = JSON.parse(saved);
          doc.title = parsed.title;
          doc.content = parsed.content;
        } catch (e) {
          console.error("Failed to parse local edit for " + doc.id, e);
        }
      }
    });
  };

  // Theme Management
  const initTheme = () => {
    const savedTheme = localStorage.getItem('theme') || 'light';
    if (savedTheme === 'dark') {
      document.body.classList.add('dark-theme');
      themeIcon.textContent = '☀️';
    } else {
      document.body.classList.remove('dark-theme');
      themeIcon.textContent = '🌙';
    }
  };

  themeToggleBtn.addEventListener('click', () => {
    document.body.classList.toggle('dark-theme');
    const isDark = document.body.classList.contains('dark-theme');
    themeIcon.textContent = isDark ? '☀️' : '🌙';
    localStorage.setItem('theme', isDark ? 'dark' : 'light');
  });

  // Mobile Menu Toggle
  menuToggleBtn.addEventListener('click', () => {
    sidebar.classList.toggle('active');
  });

  // Close sidebar when clicking outside on mobile
  document.addEventListener('click', (e) => {
    if (window.innerWidth <= 768) {
      if (!sidebar.contains(e.target) && !menuToggleBtn.contains(e.target) && sidebar.classList.contains('active')) {
        sidebar.classList.remove('active');
      }
    }
  });

  // Render Sidebar Tree
  const renderSidebar = () => {
    sidebarNav.innerHTML = '';
    
    NAV_STRUCTURE.forEach(group => {
      // Group header
      const groupEl = document.createElement('div');
      groupEl.className = 'tree-node';
      
      const groupLink = document.createElement('div');
      groupLink.className = 'tree-link';
      groupLink.style.fontWeight = '700';
      groupLink.style.fontSize = '12px';
      groupLink.style.textTransform = 'uppercase';
      groupLink.style.color = 'var(--text-muted)';
      groupLink.style.letterSpacing = '0.5px';
      groupLink.style.padding = '12px 12px 6px 12px';
      groupLink.innerHTML = `<span class="tree-arrow expanded">▸</span>${group.group}`;
      
      const childrenContainer = document.createElement('div');
      childrenContainer.className = 'tree-children expanded';
      
      // Toggle folder visibility
      groupLink.addEventListener('click', () => {
        const arrow = groupLink.querySelector('.tree-arrow');
        arrow.classList.toggle('expanded');
        childrenContainer.classList.toggle('expanded');
      });
      
      groupEl.appendChild(groupLink);
      
      // Group child items
      group.items.forEach(docId => {
        const doc = DOCS_DATA.find(d => d.id === docId);
        if (doc) {
          const itemEl = document.createElement('a');
          itemEl.className = `tree-link ${doc.id === currentDocId ? 'active' : ''}`;
          itemEl.innerHTML = `<span class="tree-icon">${doc.icon}</span>${doc.title}`;
          itemEl.href = `#${doc.id}`;
          
          itemEl.addEventListener('click', (e) => {
            e.preventDefault();
            loadDocument(doc.id);
            if (window.innerWidth <= 768) {
              sidebar.classList.remove('active');
            }
          });
          
          childrenContainer.appendChild(itemEl);
        }
      });
      
      groupEl.appendChild(childrenContainer);
      sidebarNav.appendChild(groupEl);
    });
  };

  // Load Document Content
  const loadDocument = (docId) => {
    const doc = DOCS_DATA.find(d => d.id === docId);
    if (!doc) return;
    
    currentDocId = docId;
    window.location.hash = docId;
    isEditing = false;
    editPageBtn.innerHTML = '<span>✏️</span> Edit';
    editPageBtn.classList.remove('btn-secondary');
    
    // Update breadcrumbs
    breadcrumbCategory.textContent = doc.path[0];
    breadcrumbPage.textContent = doc.title;
    
    // Render content HTML
    contentContainer.innerHTML = `
      <h1 class="doc-title">${doc.title}</h1>
      <div class="doc-metadata">
        <span>📅 Last Updated: ${doc.lastUpdated}</span>
        <span>📁 Path: ${doc.path.join(' ➔ ')}</span>
      </div>
      <div class="doc-content">
        ${doc.content}
      </div>
    `;

    // Re-highlight active links in sidebar
    document.querySelectorAll('.tree-children a').forEach(el => {
      const hrefId = el.getAttribute('href').replace('#', '');
      if (hrefId === docId) {
        el.classList.add('active');
      } else {
        el.classList.remove('active');
      }
    });

    // Scroll back to top
    document.querySelector('.content-wrapper').scrollTop = 0;
  };

  // Toggle Edit/View Modes
  const toggleEditMode = () => {
    const doc = DOCS_DATA.find(d => d.id === currentDocId);
    if (!doc) return;

    isEditing = !isEditing;

    if (isEditing) {
      editPageBtn.innerHTML = '<span>❌</span> Cancel';
      editPageBtn.classList.add('btn-secondary');
      
      const hasLocalEdit = !!localStorage.getItem('docs_edits_' + currentDocId);
      
      contentContainer.innerHTML = `
        <h1 class="doc-title">Edit: ${escapeHTML(doc.title)}</h1>
        <div class="doc-metadata">
          <span>📁 Editing Path: ${doc.path.join(' ➔ ')}</span>
        </div>
        
        <div class="edit-form">
          <div class="edit-field">
            <label for="edit-title">Page Title</label>
            <input type="text" id="edit-title" class="edit-input" value="${escapeHTML(doc.title)}">
          </div>
          <div class="edit-field">
            <label for="edit-content">Content (HTML / Markdown)</label>
            <textarea id="edit-content" class="edit-textarea">${escapeHTML(doc.content)}</textarea>
          </div>
          <div class="edit-actions">
            <button id="edit-save-btn" class="btn btn-primary">Save Changes</button>
            <button id="edit-cancel-btn" class="btn btn-secondary">Cancel</button>
            ${hasLocalEdit ? `<button id="edit-reset-btn" class="btn btn-danger" style="margin-left: auto;">Reset to Default</button>` : ''}
          </div>
        </div>
      `;

      // Bind form buttons
      document.getElementById('edit-save-btn').addEventListener('click', saveDocumentEdits);
      document.getElementById('edit-cancel-btn').addEventListener('click', cancelDocumentEdits);
      
      const resetBtn = document.getElementById('edit-reset-btn');
      if (resetBtn) {
        resetBtn.addEventListener('click', resetDocumentToDefault);
      }
    } else {
      editPageBtn.innerHTML = '<span>✏️</span> Edit';
      editPageBtn.classList.remove('btn-secondary');
      loadDocument(currentDocId);
    }
  };

  const saveDocumentEdits = () => {
    const doc = DOCS_DATA.find(d => d.id === currentDocId);
    if (!doc) return;

    const newTitle = document.getElementById('edit-title').value.trim();
    const newContent = document.getElementById('edit-content').value;

    if (!newTitle) {
      alert("Page title cannot be empty!");
      return;
    }

    doc.title = newTitle;
    doc.content = newContent;

    // Save in localStorage
    localStorage.setItem('docs_edits_' + currentDocId, JSON.stringify({
      title: newTitle,
      content: newContent
    }));

    isEditing = false;
    editPageBtn.innerHTML = '<span>✏️</span> Edit';
    editPageBtn.classList.remove('btn-secondary');
    
    renderSidebar();
    loadDocument(currentDocId);
  };

  const cancelDocumentEdits = () => {
    isEditing = false;
    editPageBtn.innerHTML = '<span>✏️</span> Edit';
    editPageBtn.classList.remove('btn-secondary');
    loadDocument(currentDocId);
  };

  const resetDocumentToDefault = () => {
    if (!confirm("Revert all edits to the original documentation content?")) return;

    const original = ORIGINAL_DOCS.find(d => d.id === currentDocId);
    const doc = DOCS_DATA.find(d => d.id === currentDocId);
    
    if (original && doc) {
      doc.title = original.title;
      doc.content = original.content;
      localStorage.removeItem('docs_edits_' + currentDocId);
    }

    isEditing = false;
    editPageBtn.innerHTML = '<span>✏️</span> Edit';
    editPageBtn.classList.remove('btn-secondary');
    
    renderSidebar();
    loadDocument(currentDocId);
  };

  // Bind Edit button click
  editPageBtn.addEventListener('click', toggleEditMode);

  // Open Export Modal
  exportDataBtn.addEventListener('click', () => {
    let fileContent = `// Unified data storage for HomeFront Confluence documentation system\nconst DOCS_DATA = [\n`;
    
    DOCS_DATA.forEach((doc, idx) => {
      fileContent += `  {\n`;
      fileContent += `    id: ${JSON.stringify(doc.id)},\n`;
      fileContent += `    title: ${JSON.stringify(doc.title)},\n`;
      fileContent += `    icon: ${JSON.stringify(doc.icon)},\n`;
      fileContent += `    path: ${JSON.stringify(doc.path)},\n`;
      fileContent += `    lastUpdated: ${JSON.stringify(doc.lastUpdated)},\n`;
      // Escape backticks in content template strings
      const sanitizedContent = doc.content.replace(/`/g, '\\`').replace(/\${/g, '\\${');
      fileContent += `    content: \`${sanitizedContent}\`\n`;
      fileContent += `  }${idx < DOCS_DATA.length - 1 ? ',' : ''}\n`;
    });
    
    fileContent += `];\n`;

    exportTextarea.value = fileContent;
    exportModal.classList.add('active');
  });

  // Close Export Modal
  const closeModal = () => {
    exportModal.classList.remove('active');
  };

  modalCloseBtn.addEventListener('click', closeModal);
  exportModal.addEventListener('click', (e) => {
    if (e.target === exportModal) {
      closeModal();
    }
  });

  // Copy to Clipboard
  copyExportBtn.addEventListener('click', () => {
    exportTextarea.select();
    document.execCommand('copy');
    
    const originalText = copyExportBtn.textContent;
    copyExportBtn.textContent = 'Copied!';
    copyExportBtn.style.backgroundColor = 'var(--tip-border)';
    
    setTimeout(() => {
      copyExportBtn.textContent = originalText;
      copyExportBtn.style.backgroundColor = '';
    }, 2000);
  });

  // Download file data.js
  downloadExportBtn.addEventListener('click', () => {
    const text = exportTextarea.value;
    const blob = new Blob([text], { type: 'text/javascript' });
    const url = URL.createObjectURL(blob);
    
    const a = document.createElement('a');
    a.href = url;
    a.download = 'data.js';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

  // Live Search Engine
  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      searchResults.classList.remove('active');
      return;
    }
    
    searchResults.innerHTML = '';
    const matches = [];
    
    DOCS_DATA.forEach(doc => {
      const titleMatch = doc.title.toLowerCase().includes(query);
      
      const textContent = doc.content.replace(/<[^>]*>/g, ' ');
      const contentIndex = textContent.toLowerCase().indexOf(query);
      const contentMatch = contentIndex !== -1;
      
      if (titleMatch || contentMatch) {
        let snippet = '';
        if (contentMatch) {
          const start = Math.max(0, contentIndex - 40);
          const end = Math.min(textContent.length, contentIndex + query.length + 80);
          snippet = '...' + textContent.slice(start, end).trim() + '...';
        } else {
          snippet = textContent.slice(0, 120).trim() + '...';
        }
        
        matches.push({
          id: doc.id,
          title: doc.title,
          path: doc.path.join(' ➔ '),
          snippet: snippet
        });
      }
    });
    
    if (matches.length > 0) {
      matches.forEach(match => {
        const item = document.createElement('div');
        item.className = 'search-item';
        item.innerHTML = `
          <div class="search-item-title">${match.title}</div>
          <div class="search-item-path">${match.path}</div>
          <div class="search-item-snippet">${escapeHTML(match.snippet)}</div>
        `;
        
        item.addEventListener('click', () => {
          loadDocument(match.id);
          searchInput.value = '';
          searchResults.classList.remove('active');
        });
        
        searchResults.appendChild(item);
      });
      searchResults.classList.add('active');
    } else {
      searchResults.innerHTML = '<div style="padding: 16px; text-align: center; color: var(--text-muted); font-size: 14px;">No results found</div>';
      searchResults.classList.add('active');
    }
  });

  // Hide search results on click outside
  document.addEventListener('click', (e) => {
    if (!searchInput.contains(e.target) && !searchResults.contains(e.target)) {
      searchResults.classList.remove('active');
    }
  });

  // Tab switcher in content
  window.switchTab = (event, tabId) => {
    const tabContainer = event.currentTarget.parentElement;
    tabContainer.querySelectorAll('.tab-btn').forEach(btn => {
      btn.classList.remove('active');
    });
    
    event.currentTarget.classList.add('active');
    
    const tabPanesContainer = tabContainer.parentElement;
    tabPanesContainer.querySelectorAll('.tab-pane').forEach(pane => {
      pane.classList.remove('active');
    });
    
    const targetPane = document.getElementById(tabId);
    if (targetPane) {
      targetPane.classList.add('active');
    }
  };

  // Helper function to escape HTML tags in snippets
  function escapeHTML(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  }

  // Hash-based Routing
  const handleRouting = () => {
    const hash = window.location.hash.replace('#', '');
    if (hash && DOCS_DATA.some(d => d.id === hash)) {
      loadDocument(hash);
    } else {
      loadDocument('overview');
    }
  };

  // Initialize App
  initTheme();
  loadLocalEdits(); // Load local storage modifications on boot
  renderSidebar();
  handleRouting();
  
  // Listen for hash changes
  window.addEventListener('hashchange', handleRouting);
});
