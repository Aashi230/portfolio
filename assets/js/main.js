/**
 * AASHI SRIVASTAVA — PERSONAL PORTFOLIO & ARCHIVE
 * JavaScript Controller (Vintage Editorial Edition)
 */

document.addEventListener('DOMContentLoaded', () => {
  'use strict';

  // ========================================================================
  // 1. SUBTLE VINTAGE EDITORIAL CURSOR (Desktop only)
  // ========================================================================
  const cursor = document.getElementById('customCursor');
  const cursorDot = document.getElementById('customCursorDot');

  if (cursor && cursorDot && window.matchMedia('(pointer: fine)').matches) {
    let mouseX = -100, mouseY = -100;
    let cursorX = -100, cursorY = -100;

    window.addEventListener('mousemove', (e) => {
      mouseX = e.clientX;
      mouseY = e.clientY;
      cursorDot.style.left = `${mouseX}px`;
      cursorDot.style.top = `${mouseY}px`;
    });

    const renderCursor = () => {
      cursorX += (mouseX - cursorX) * 0.18;
      cursorY += (mouseY - cursorY) * 0.18;
      cursor.style.left = `${cursorX}px`;
      cursor.style.top = `${cursorY}px`;
      requestAnimationFrame(renderCursor);
    };
    requestAnimationFrame(renderCursor);

    // Hover scale effects on interactive elements
    const interactiveElements = document.querySelectorAll('a, button, input, textarea, .cert-item-card, .project-feature-card');
    interactiveElements.forEach(el => {
      el.addEventListener('mouseenter', () => cursor.classList.add('cursor-hover'));
      el.addEventListener('mouseleave', () => cursor.classList.remove('cursor-hover'));
    });
  }

  // ========================================================================
  // 2. NAVIGATION SCROLL & ACTIVE LINK SPY
  // ========================================================================
  const header = document.getElementById('mainHeader');
  const navLinks = document.querySelectorAll('.nav-link');
  const sections = document.querySelectorAll('section[id], footer[id]');
  const mobileToggle = document.getElementById('mobileToggle');
  const navMenu = document.getElementById('navMenu');

  window.addEventListener('scroll', () => {
    if (window.scrollY > 40) {
      header.classList.add('scrolled');
    } else {
      header.classList.remove('scrolled');
    }
  }, { passive: true });

  const sectionObserver = new IntersectionObserver((entries) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        const id = entry.target.getAttribute('id');
        navLinks.forEach(link => {
          if (link.getAttribute('href') === `#${id}`) {
            link.classList.add('active');
          } else {
            link.classList.remove('active');
          }
        });
      }
    });
  }, { root: null, rootMargin: '-20% 0px -70% 0px', threshold: 0 });

  sections.forEach(section => sectionObserver.observe(section));

  if (mobileToggle && navMenu) {
    mobileToggle.addEventListener('click', () => {
      const isOpen = navMenu.classList.toggle('active');
      mobileToggle.setAttribute('aria-expanded', isOpen);
    });

    navLinks.forEach(link => {
      link.addEventListener('click', () => {
        navMenu.classList.remove('active');
        mobileToggle.setAttribute('aria-expanded', 'false');
      });
    });

    document.addEventListener('click', (e) => {
      if (!navMenu.contains(e.target) && !mobileToggle.contains(e.target) && navMenu.classList.contains('active')) {
        navMenu.classList.remove('active');
        mobileToggle.setAttribute('aria-expanded', 'false');
      }
    });
  }

  // Smooth scroll offset adjustment for sticky header
  document.querySelectorAll('a[href^="#"]').forEach(anchor => {
    anchor.addEventListener('click', function(e) {
      const targetId = this.getAttribute('href');
      if (targetId === '#' || targetId === '') return;
      const targetElement = document.querySelector(targetId);
      if (targetElement) {
        e.preventDefault();
        const headerOffset = 76;
        const elementPosition = targetElement.getBoundingClientRect().top;
        const offsetPosition = elementPosition + window.pageYOffset - headerOffset;

        window.scrollTo({
          top: offsetPosition,
          behavior: 'smooth'
        });
      }
    });
  });

  // ========================================================================
  // 3. BACK TO TOP BUTTON
  // ========================================================================
  const backToTop = document.getElementById('backToTop');
  if (backToTop) {
    window.addEventListener('scroll', () => {
      if (window.scrollY > 400) {
        backToTop.classList.add('visible');
      } else {
        backToTop.classList.remove('visible');
      }
    }, { passive: true });

    backToTop.addEventListener('click', () => {
      window.scrollTo({
        top: 0,
        behavior: 'smooth'
      });
    });
  }

  // ========================================================================
  // 4. CERTIFICATE FILTERING
  // ========================================================================
  const certFilterBtns = document.querySelectorAll('.cert-filter-btn');
  const certCards = document.querySelectorAll('.cert-item-card');

  certFilterBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      certFilterBtns.forEach(b => b.classList.remove('active'));
      btn.classList.add('active');

      const filter = btn.getAttribute('data-category');

      certCards.forEach(card => {
        const category = card.getAttribute('data-category');
        if (filter === 'all' || category === filter) {
          card.style.display = 'flex';
          setTimeout(() => {
            card.style.opacity = '1';
            card.style.transform = 'translateY(0)';
          }, 10);
        } else {
          card.style.opacity = '0';
          card.style.transform = 'translateY(10px)';
          setTimeout(() => {
            card.style.display = 'none';
          }, 200);
        }
      });
    });
  });

  // ========================================================================
  // 5. CERTIFICATE LIGHTBOX MODAL
  // ========================================================================
  const certModal = document.getElementById('certModalBackdrop');
  const modalCertImg = document.getElementById('modalCertImg');
  const modalCertTitle = document.getElementById('modalCertTitle');
  const modalCertOrg = document.getElementById('modalCertOrg');
  const modalCertDate = document.getElementById('modalCertDate');
  const modalCertCategory = document.getElementById('modalCertCategory');
  const modalCertOpenDoc = document.getElementById('modalCertOpenDoc');
  const btnCloseCertModal = document.getElementById('btnCloseCertModal');
  const modalCertCloseBtn = document.getElementById('modalCertCloseBtn');

  const openCertModal = (card) => {
    const title = card.getAttribute('data-title');
    const org = card.getAttribute('data-org');
    const date = card.getAttribute('data-date');
    const category = card.getAttribute('data-category');
    const imgUrl = card.getAttribute('data-img');
    const pdfUrl = card.getAttribute('data-pdf');

    modalCertTitle.textContent = title;
    modalCertOrg.textContent = org;
    modalCertDate.textContent = `Issued: ${date}`;
    modalCertCategory.textContent = category.toUpperCase();
    modalCertImg.src = imgUrl;

    if (pdfUrl) {
      modalCertOpenDoc.href = pdfUrl;
      modalCertOpenDoc.style.display = 'inline-flex';
    } else {
      modalCertOpenDoc.style.display = 'none';
    }

    certModal.classList.add('active');
    document.body.style.overflow = 'hidden';
  };

  const closeCertModal = () => {
    certModal.classList.remove('active');
    document.body.style.overflow = '';
  };

  certCards.forEach(card => {
    card.addEventListener('click', () => openCertModal(card));
  });

  // Photo Cell Lightbox Interaction (16 Event Photos)
  const photoCells = document.querySelectorAll('.photo-cell');
  photoCells.forEach(cell => {
    cell.addEventListener('click', () => {
      const img = cell.querySelector('img');
      const title = cell.getAttribute('data-title') || (img ? img.alt : 'Archive Photograph');
      const event = cell.getAttribute('data-event') || 'Event Chronicle &bull; Aashi Srivastava';
      const date = cell.getAttribute('data-date') || 'Curated Photographic Record';
      const category = cell.getAttribute('data-category') || 'FIELD ARCHIVE';
      const imgSrc = cell.getAttribute('data-img') || (img ? img.src : '');

      if (modalCertTitle) modalCertTitle.textContent = title;
      if (modalCertOrg) modalCertOrg.innerHTML = event;
      if (modalCertDate) modalCertDate.textContent = date;
      if (modalCertCategory) modalCertCategory.textContent = category.toUpperCase();
      if (modalCertImg) modalCertImg.src = imgSrc;

      if (modalCertOpenDoc) {
        modalCertOpenDoc.href = imgSrc;
        modalCertOpenDoc.textContent = 'Inspect Full Image \u2197';
        modalCertOpenDoc.style.display = 'inline-flex';
      }

      if (certModal) {
        certModal.classList.add('active');
        document.body.style.overflow = 'hidden';
      }
    });
  });

  if (btnCloseCertModal) btnCloseCertModal.addEventListener('click', closeCertModal);
  if (modalCertCloseBtn) modalCertCloseBtn.addEventListener('click', closeCertModal);
  if (certModal) {
    certModal.addEventListener('click', (e) => {
      if (e.target === certModal) closeCertModal();
    });
  }

  // ========================================================================
  // 6. PROJECT BRIEF MODALS (5 PROJECTS)
  // ========================================================================
  const projectModal = document.getElementById('projectModalBackdrop');
  const modalProjectTag = document.getElementById('modalProjectTag');
  const modalProjectTitle = document.getElementById('modalProjectTitle');
  const modalProjectSubtitle = document.getElementById('modalProjectSubtitle');
  const modalProjectContent = document.getElementById('modalProjectContent');
  const modalProjectGithubLink = document.getElementById('modalProjectGithubLink');
  const btnCloseProjectModal = document.getElementById('btnCloseProjectModal');
  const modalProjectCloseBtn = document.getElementById('modalProjectCloseBtn');

  const projectDossiers = {
    hunter: {
      tag: 'OPPORTUNITY HUNTER // DISCOVERY PLATFORM',
      title: 'Opportunity Hunter',
      subtitle: 'Internship & Scholarship Discovery Platform for Students',
      github: 'https://github.com/Aashi230',
      content: `
        <div style="font-size: 0.96rem; line-height: 1.7; color: var(--text-secondary);">
          <p style="margin-bottom: 14px;">
            <strong>Problem Statement:</strong> Students frequently miss out on critical undergraduate scholarships, research fellowships, and industry internships simply due to fragmented notice boards, inaccurate deadlines, and predatory application portals.
          </p>
          <p style="margin-bottom: 14px;">
            <strong>Implementation &amp; Tech Stack:</strong> Built using <strong>HTML5, CSS3, and JavaScript</strong>. Opportunity Hunter delivers a structured discovery interface that organizes and categorizes verified scholarship and internship listings by academic discipline, eligibility criteria, and deadlines.
          </p>
          <div style="background: var(--bg-mat); padding: 14px; border: 1px solid var(--color-soft-beige); border-radius: 3px; margin-bottom: 14px;">
            <strong style="color: var(--color-burgundy);">Key System Capabilities:</strong>
            <ul style="padding-left: 18px; margin-top: 6px;">
              <li>Curated directory of legitimate academic scholarships, grants, and internships.</li>
              <li>Verification of official application links to avoid fraudulent third parties.</li>
              <li>Filterable discovery by eligibility constraints, semester, and academic focus.</li>
              <li>Clean, accessible client-side architecture with zero intermediary tracking.</li>
            </ul>
          </div>
          <p>
            <strong>Role &amp; Context:</strong> Developed and showcased during Buildethon (organized by Nerds Room). Designed responsive layouts, structured opportunity catalogs, and verified user flows.
          </p>
        </div>
      `
    },
    ars: {
      tag: 'SMART INDIA HACKATHON 2026 // INTERNAL SELECTION CLEARED',
      title: 'ARS Tourist Guide',
      subtitle: 'Cleared University-Level Internal Round • Problem Statement SIH26207 • Team TOTAL CODERS',
      github: 'https://github.com/Aashi230',
      content: `
        <div style="font-size: 0.96rem; line-height: 1.7; color: var(--text-secondary);">
          <p style="margin-bottom: 14px;">
            <strong>Smart India Hackathon Achievement:</strong> Participated in the <strong>Smart India Hackathon 2026</strong> internal selection with <strong>Team TOTAL CODERS</strong> for the <strong>Travel &amp; Tourism problem statement SIH26207</strong> and cleared the university-level internal round at Shri Ramswaroop Memorial University (SRMU).
          </p>
          <div style="background: var(--bg-mat); padding: 14px; border: 1px solid var(--color-soft-beige); border-radius: 3px; margin-bottom: 14px;">
            <strong style="color: var(--color-burgundy);">Full Technology Stack:</strong>
            <p style="margin-top: 4px;">
              <strong>Backend Architecture:</strong> Node.js &amp; Express.js REST API<br>
              <strong>Database:</strong> SQLite lightweight relational engine<br>
              <strong>Payment Gateway:</strong> Razorpay API integration for tour pass reservations<br>
              <strong>Spatial &amp; Weather Services:</strong> OpenStreetMap API &amp; Open-Meteo REST endpoints
            </p>
          </div>
          <p style="margin-bottom: 14px;">
            <strong>Engineering Highlights:</strong> Integrated asynchronous third-party RESTful APIs, designed clean spatial navigation overlays, handled webhooks for Razorpay transaction states, and defended live working prototypes before university evaluation juries.
          </p>
          <p>
            <strong>Team Collaboration:</strong> Cleared internal rounds as part of Team TOTAL CODERS, demonstrating rapid feature engineering, system architecture, and technical pitch defense.
          </p>
        </div>
      `
    },
    hostel: {
      tag: 'CAMPUS OPERATIONS // WEB APPLICATION',
      title: 'Hostel Complaint Management System',
      subtitle: 'Streamlined Grievance Resolution & Tracking Portal',
      github: 'https://github.com/Aashi230',
      content: `
        <div style="font-size: 0.96rem; line-height: 1.7; color: var(--text-secondary);">
          <p style="margin-bottom: 14px;">
            <strong>Objective:</strong> Modernize campus hostel operations by replacing slow, manual register logs with a real-time web portal that tracks issues from submission to sign-off.
          </p>
          <div style="background: var(--bg-mat); padding: 14px; border: 1px solid var(--color-soft-beige); border-radius: 3px; margin-bottom: 14px;">
            <strong style="color: var(--color-burgundy);">Architecture &amp; Database:</strong>
            <p style="margin-top: 4px;">
              Built on <strong>Node.js</strong> and <strong>Express.js</strong> with <strong>MongoDB</strong> and <strong>Mongoose</strong>. Employs a robust document schema that retains complaint category, room details, priority level, warden assignments, and audit timestamp history.
            </p>
          </div>
          <p>
            <strong>Key Features:</strong> Distinct student filing views, administrative management dashboards, real-time status transitions (Pending ➔ In Progress ➔ Resolved), and turnaround analytics.
          </p>
        </div>
      `
    },
    voting: {
      tag: 'DIGITAL VOTING // HARDWARE LOGIC',
      title: 'Interactive Digital Voting Machine',
      subtitle: 'Verilog HDL Synthesized Electronic Voting Machine',
      github: 'https://github.com/Aashi230',
      content: `
        <div style="font-size: 0.96rem; line-height: 1.7; color: var(--text-secondary);">
          <p style="margin-bottom: 14px;">
            <strong>System Description:</strong> An electronic voting machine logic circuit modeled and synthesized in Verilog HDL. Explores digital design fundamentals including state machines, counters, and registers.
          </p>
          <div style="background: var(--bg-mat); padding: 14px; border: 1px solid var(--color-soft-beige); border-radius: 3px; margin-bottom: 14px;">
            <strong style="color: var(--color-dark-brown);">Hardware Logic Highlights:</strong>
            <ul style="padding-left: 18px; margin-top: 6px;">
              <li>Synchronous clock-driven vote accumulation for individual candidates.</li>
              <li>Input button debouncing to eliminate mechanical switch glitches.</li>
              <li>Combinational magnitude comparators for real-time winner detection.</li>
              <li>Comprehensive testbench verification across edge tally tie conditions.</li>
            </ul>
          </div>
          <p>
            <strong>Relevance:</strong> Demonstrates a rigorous understanding of hardware description languages (Verilog), digital gates, and sequential logic alongside software engineering skills.
          </p>
        </div>
      `
    }
  };

  const openProjectModal = (projectKey) => {
    const data = projectDossiers[projectKey];
    if (!data) return;

    modalProjectTag.textContent = data.tag;
    modalProjectTitle.textContent = data.title;
    modalProjectSubtitle.textContent = data.subtitle;
    modalProjectContent.innerHTML = data.content;
    modalProjectGithubLink.href = data.github;

    projectModal.classList.add('active');
    document.body.style.overflow = 'hidden';
  };

  const closeProjectModal = () => {
    projectModal.classList.remove('active');
    document.body.style.overflow = '';
  };

  document.querySelectorAll('.open-project-modal').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.preventDefault();
      const projectKey = btn.getAttribute('data-project');
      openProjectModal(projectKey);
    });
  });

  if (btnCloseProjectModal) btnCloseProjectModal.addEventListener('click', closeProjectModal);
  if (modalProjectCloseBtn) modalProjectCloseBtn.addEventListener('click', closeProjectModal);
  if (projectModal) {
    projectModal.addEventListener('click', (e) => {
      if (e.target === projectModal) closeProjectModal();
    });
  }

  // Close modals on Escape key
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      closeCertModal();
      closeProjectModal();
    }
  });

  // ========================================================================
  // 7. TOAST NOTIFICATION UTILITY
  // ========================================================================
  const toastContainer = document.getElementById('toastContainer');

  const showToast = (message, icon = '✦') => {
    if (!toastContainer) return;
    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerHTML = `<span style="color: var(--color-olive);">${icon}</span> <span>${message}</span>`;
    toastContainer.appendChild(toast);

    setTimeout(() => {
      toast.classList.add('show');
    }, 10);

    setTimeout(() => {
      toast.classList.remove('show');
      setTimeout(() => {
        toast.remove();
      }, 300);
    }, 3500);
  };

  // ========================================================================
  // 8. COPY EMAIL TO CLIPBOARD
  // ========================================================================
  const btnCopyEmail = document.getElementById('btnCopyEmail');
  if (btnCopyEmail) {
    btnCopyEmail.addEventListener('click', () => {
      const email = btnCopyEmail.getAttribute('data-email');
      navigator.clipboard.writeText(email).then(() => {
        showToast('Email address copied to clipboard!');
      }).catch(() => {
        showToast(`Email: ${email}`);
      });
    });
  }

  // ========================================================================
  // 9. CONTACT FORM VALIDATION & SUBMISSION (TRANSMIT TO GMAIL)
  // ========================================================================
  const contactForm = document.getElementById('contactForm');
  const btnSubmitForm = document.getElementById('btnSubmitForm');
  const contactFormStatus = document.getElementById('contactFormStatus');

  if (contactForm) {
    contactForm.addEventListener('submit', async (e) => {
      e.preventDefault();

      const name = document.getElementById('contactName').value.trim();
      const email = document.getElementById('contactEmail').value.trim();
      const subject = document.getElementById('contactSubject') ? document.getElementById('contactSubject').value.trim() : 'Portfolio Inquiry';
      const message = document.getElementById('contactMessage').value.trim();

      if (!name || !email || !message) {
        showToast('Please fill in all required fields (Name, Email, Message).', '⚠');
        return;
      }

      const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
      if (!emailRegex.test(email)) {
        showToast('Please provide a valid email address.', '⚠');
        return;
      }

      // Indicate transmission in progress
      const originalBtnHTML = btnSubmitForm ? btnSubmitForm.innerHTML : 'Transmit Message';
      if (btnSubmitForm) {
        btnSubmitForm.disabled = true;
        btnSubmitForm.innerHTML = `<span>Transmitting Dispatch...</span> <span class="status-pulse-dot" style="margin-left: 8px;"></span>`;
      }

      try {
        const response = await fetch("https://formsubmit.co/ajax/aashimaahi23@gmail.com", {
          method: "POST",
          headers: {
            'Content-Type': 'application/json',
            'Accept': 'application/json'
          },
          body: JSON.stringify({
            name: name,
            email: email,
            subject: subject || 'New Portfolio Dispatch',
            message: message,
            _captcha: "false",
            _template: 'table',
            _subject: `[Portfolio Dispatch] ${subject || 'New Message'} from ${name}`
          })
        });

        const data = await response.json().catch(() => ({}));

        if (response.ok && (data.success === true || data.success === "true")) {
          showToast(`Dispatch transmitted! Delivered to Aashi.`, '✓');
          if (contactFormStatus) {
            contactFormStatus.style.display = 'block';
            contactFormStatus.style.background = 'rgba(104, 106, 76, 0.15)';
            contactFormStatus.style.borderColor = 'var(--color-olive)';
            contactFormStatus.style.color = 'var(--color-dark-brown)';
            contactFormStatus.innerHTML = `&check; <strong>Dispatch Delivered:</strong> Thank you, ${name}! Your message has been transmitted directly to Aashi's inbox.`;
          }
          contactForm.reset();
        } else if (data.message && data.message.toLowerCase().includes('activation')) {
          showToast(`Activation required: Check aashimaahi23@gmail.com`, '⚠');
          if (contactFormStatus) {
            contactFormStatus.style.display = 'block';
            contactFormStatus.style.background = 'rgba(99, 45, 56, 0.12)';
            contactFormStatus.style.borderColor = 'var(--color-burgundy)';
            contactFormStatus.style.color = 'var(--color-dark-brown)';
            contactFormStatus.innerHTML = `&#9888; <strong>Activation Email Sent:</strong> FormSubmit sent a one-time activation link to <strong>aashimaahi23@gmail.com</strong>. Please check your Inbox (or Spam folder) and click <em>"Activate Form"</em> to enable incoming dispatches!`;
          }
        } else {
          throw new Error(data.message || 'FormSubmit endpoint response not OK');
        }
      } catch (err) {
        // Fallback to mailto client opener
        const encodedSubject = encodeURIComponent(subject || `Portfolio Dispatch from ${name}`);
        const encodedBody = encodeURIComponent(`From: ${name} (${email})\n\nMessage:\n${message}`);
        window.location.href = `mailto:aashimaahi23@gmail.com?subject=${encodedSubject}&body=${encodedBody}`;

        showToast(`Preparing email to aashimaahi23@gmail.com...`, '✓');
        if (contactFormStatus) {
          contactFormStatus.style.display = 'block';
          contactFormStatus.style.background = 'rgba(99, 45, 56, 0.1)';
          contactFormStatus.style.borderColor = 'var(--color-burgundy)';
          contactFormStatus.style.color = 'var(--color-dark-brown)';
          contactFormStatus.innerHTML = `&nearr; <strong>Email Client Prepared:</strong> Message drafted to <code>aashimaahi23@gmail.com</code>.`;
        }
      } finally {
        if (btnSubmitForm) {
          btnSubmitForm.disabled = false;
          btnSubmitForm.innerHTML = originalBtnHTML;
        }
      }
    });
  }

  // ========================================================================
  // 10. REVEAL ANIMATIONS ON SCROLL
  // ========================================================================
  const animatedElements = document.querySelectorAll('.skills-category-card, .project-feature-card, .cert-item-card, .timeline-inner-card');
  
  const revealObserver = new IntersectionObserver((entries, observer) => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.style.opacity = '1';
        entry.target.style.transform = 'translateY(0)';
        observer.unobserve(entry.target);
      }
    });
  }, { threshold: 0.08 });

  animatedElements.forEach(el => {
    el.style.opacity = '0';
    el.style.transform = 'translateY(16px)';
    el.style.transition = 'opacity 0.45s ease, transform 0.45s ease';
    revealObserver.observe(el);
  });

  console.log('✦ Aashi Srivastava Portfolio Engine initialized in Luxury Vintage Editorial Edition.');
});
