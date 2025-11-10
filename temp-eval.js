const vm = require('vm'); const code = @'

  function availabilityApp() {
    return {
      email: '',
      services: [],
      selectedServiceIds: [],
      loadingServices: true,
      loadingAvailability: false,
      saving: false,
      error: '',
      statusMessage: '',
      prefilledFor: '',
      memberVerified: false,
      verifiedEmail: '',
      shareLink: '',
      emailSubject: '',
      emailBody: '',
      emailEditorOpen: false,
      isAdmin: false,
      viewerProfile: null,
      recipientLoading: false,
      recipientError: '',
      recipientTeams: [{ value: 'all', label: 'All members' }],
      recipientMembers: [],
      filteredRecipients: [],
      selectedRecipientTeam: 'all',
      selectedRecipientEmails: [],

      init() {
        this.prepareEmailTemplate();
        this.loadViewerProfile();
        this.fetchServices();
        try {
          const params = new URLSearchParams(location.search || '');
          const emailParam = params.get('email');
          if (emailParam) {
            this.email = emailParam;
            this.prefillAvailability(true);
          }
        } catch (_) { /* ignore */ }
      },

      async loadViewerProfile() {
        try {
          const profile = await callRpc('getViewerProfile', null);
          this.viewerProfile = profile || null;
          this.isAdmin = Boolean(profile?.isAdmin);
          if (this.isAdmin) {
            this.loadEmailRecipients();
          }
        } catch (e) {
          console.error(e);
          this.viewerProfile = null;
          this.isAdmin = false;
        }
      },

      prepareEmailTemplate() {
        this.shareLink = this.availabilityShareUrl();
        if (!this.emailSubject) {
          this.emailSubject = 'Please update your availability';
        }
        if (!this.emailBody) {
          this.emailBody = `Hi team,\n\nPlease record your upcoming availability here:\n${this.shareLink}\n\nThank you!`;
        }
      },

      availabilityShareUrl() {
        const fallbackBase = `${location.origin}${location.pathname}`;
        const base = window.__BASE__ || fallbackBase;
        try {
          const url = new URL(base);
          url.searchParams.set('mode', 'guest');
          url.searchParams.set('tab', 'availability');
          if (!url.hash || url.hash === '#') {
            url.hash = '#/availability';
          }
          return url.toString();
        } catch (_) {
          const join = base.includes('?') ? '&' : '?';
          return `${base}${join}mode=guest&tab=availability#/availability`;
        }
      },

      async copyShareLink() {
        await this.copyToClipboard(this.shareLink || '', 'Share link copied.');
      },

      async copyEmailSubject() {
        await this.copyToClipboard(this.emailSubject || '', 'Subject copied.');
      },

      async copyEmailBody() {
        await this.copyToClipboard(this.emailBody || '', 'Email text copied.');
      },

      async copyRecipientEmails() {
        await this.copyToClipboard(this.selectedRecipientEmails.join(', '), 'Recipient emails copied.');
      },

      async copyToClipboard(text, successMsg = 'Copied.') {
        if (!text) {
          notify('Nothing to copy.', 'error');
          return;
        }
        try {
          if (navigator?.clipboard?.writeText) {
            await navigator.clipboard.writeText(text);
          } else {
            const textarea = document.createElement('textarea');
            textarea.value = text;
            textarea.setAttribute('readonly', 'true');
            textarea.style.position = 'absolute';
            textarea.style.left = '-9999px';
            document.body.appendChild(textarea);
            textarea.select();
            document.execCommand('copy');
            textarea.remove();
          }
          notify(successMsg, 'success');
        } catch (err) {
          console.error(err);
          notify('Unable to copy to clipboard.', 'error');
        }
      },

      async loadEmailRecipients() {
        this.recipientLoading = true;
        this.recipientError = '';
        try {
          const res = await callRpc('listRoles', null);
          const items = Array.isArray(res?.items) ? res.items : [];
          const members = items
            .map(row => {
              const email = String(row?.email || '').trim().toLowerCase();
              if (!email) return null;
              const first = String(row?.first || '').trim();
              const last = String(row?.last || '').trim();
              const teams = Array.isArray(row?.teams)
                ? row.teams.filter(Boolean)
                : String(row?.teamRaw || '')
                    .split(/[,;|]/)
                    .map(v => v.trim())
                    .filter(Boolean);
              return {
                email,
                name: [first, last].filter(Boolean).join(' '),
                teams
              };
            })
            .filter(Boolean)
            .sort((a, b) => {
              const nameA = a.name || a.email;
              const nameB = b.name || b.email;
              return nameA.localeCompare(nameB);
            });
          const teamSet = new Set<string>();
          members.forEach(m => {
            if (!Array.isArray(m.teams) || !m.teams.length) return;
            m.teams.forEach(team => {
              if (team) teamSet.add(team);
            });
          });
          const teams = [{ value: 'all', label: 'All members' }];
          Array.from(teamSet)
            .sort((a, b) => a.localeCompare(b))
            .forEach(team => teams.push({ value: team, label: team }));
          this.recipientMembers = members;
          this.recipientTeams = teams;
          this.applyRecipientFilter(this.selectedRecipientTeam || 'all');
        } catch (e) {
          console.error(e);
          this.recipientError = (e && e.message) ? e.message : 'Unable to load team members.';
          this.filteredRecipients = [];
          this.selectedRecipientEmails = [];
        } finally {
          this.recipientLoading = false;
        }
      },

      applyRecipientFilter(teamValue) {
        const team = teamValue || 'all';
        this.selectedRecipientTeam = team;
        if (team === 'all') {
          this.filteredRecipients = this.recipientMembers.slice();
        } else {
          this.filteredRecipients = this.recipientMembers.filter(m => Array.isArray(m.teams) && m.teams.includes(team));
        }
        this.selectedRecipientEmails = this.filteredRecipients.map(m => m.email);
      },

      onRecipientTeamChange(value) {
        this.applyRecipientFilter(value);
      },

      todayISO() {
        const now = new Date();
        const y = now.getFullYear();
        const m = String(now.getMonth() + 1).padStart(2, '0');
        const d = String(now.getDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      },

      emailValid() {
        return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(this.email || '').trim());
      },

      normalizedEmail() {
        return String(this.email || '').trim().toLowerCase();
      },

      onEmailInput() {
        if (this.normalizedEmail() !== this.prefilledFor) {
          this.memberVerified = false;
          this.verifiedEmail = '';
          this.selectedServiceIds = [];
          this.statusMessage = '';
        }
        if (!this.emailValid()) {
          this.error = '';
        }
      },

      async verifyMemberEmail(force = false) {
        const email = this.normalizedEmail();
        if (!email) {
          this.memberVerified = false;
          this.verifiedEmail = '';
          return false;
        }
        if (!force && this.memberVerified && email === this.verifiedEmail) {
          return true;
        }
        try {
          const res = await callRpc('memberExistsInRoles', { email });
          this.memberVerified = Boolean(res?.exists);
          this.verifiedEmail = this.memberVerified ? email : '';
          if (!this.memberVerified) {
            this.error = 'Please contact Belinda to add your email to the database.';
          } else {
            this.error = '';
          }
          return this.memberVerified;
        } catch (e) {
          console.error(e);
          this.memberVerified = false;
          this.verifiedEmail = '';
          this.error = (e && e.message) ? e.message : 'Unable to verify your email.';
          return false;
        }
      },

      async fetchServices() {
        this.loadingServices = true;
        this.error = '';
        try {
          const res = await callRpc('listServices', {
            includePast: false,
            sort: 'asc',
            startDate: this.todayISO(),
            limit: 20
          });
          const items = Array.isArray(res?.items) ? res.items : [];
          this.services = items.map(it => ({
            ...it,
            label: this.formatServiceLabel(it),
            subtitle: [it.type || '', it.time || '10:00 AM'].filter(Boolean).join(' • ') || 'Service'
          }));
        } catch (e) {
          console.error(e);
          this.error = (e && e.message) ? e.message : 'Unable to load services.';
        } finally {
          this.loadingServices = false;
        }
      },

      formatServiceLabel(item) {
        const iso = String(item?.date || '').trim();
        if (!iso) return item?.id || 'Service';
        try {
          const [y, m, d] = iso.split('-').map(Number);
          const dt = new Date(y, m - 1, d);
          return dt.toLocaleDateString(undefined, { weekday: 'long', month: 'short', day: 'numeric' });
        } catch (_) {
          return iso;
        }
      },

      async prefillAvailability(force = false) {
        const email = this.normalizedEmail();
        if (!email) {
          this.error = 'Enter your email to load saved dates.';
          return;
        }
        if (!force && email === this.prefilledFor && this.memberVerified) return;
        this.loadingAvailability = true;
        this.error = '';
        try {
          const verified = await this.verifyMemberEmail(force);
          if (!verified) {
            this.selectedServiceIds = [];
            this.statusMessage = '';
            this.prefilledFor = '';
            return;
          }
          const res = await callRpc('getMemberAvailability', { email });
          this.selectedServiceIds = Array.isArray(res?.unavailable) ? res.unavailable.slice() : [];
          this.prefilledFor = email;
          this.statusMessage = this.selectedServiceIds.length
            ? 'Loaded your saved unavailable dates.'
            : 'No saved dates yet. Check every service you cannot serve.';
        } catch (e) {
          console.error(e);
          this.error = (e && e.message) ? e.message : 'Unable to load availability.';
        } finally {
          this.loadingAvailability = false;
        }
      },

      clearSelections() {
        this.selectedServiceIds = [];
      },

      async saveAvailability() {
        const email = this.normalizedEmail();
        if (!email) {
          this.error = 'Email is required.';
          return;
        }
        if (!this.memberVerified || email !== this.verifiedEmail) {
          this.error = 'Load dates for your email before saving.';
          return;
        }
        this.saving = true;
        this.error = '';
        this.statusMessage = '';
        try {
          await callRpc('saveMemberAvailability', {
            email,
            unavailableServiceIds: this.selectedServiceIds
          });
          this.statusMessage = 'Availability saved. Thank you!';
          this.prefilledFor = email;
        } catch (e) {
          console.error(e);
          this.error = (e && e.message) ? e.message : 'Unable to save availability.';
        } finally {
          this.saving = false;
        }
      }
    };
  }


'@; vm.runInNewContext(code, { console }); console.log('ok');
