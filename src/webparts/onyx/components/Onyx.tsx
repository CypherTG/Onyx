import * as React from 'react';
import styles from './Onyx.module.scss';
import type { IOnyxProps } from './IOnyxProps';
import OnyxSharePointService, {
  IOnyxBriefItem,
  IOnyxFeatureItem,
  IOnyxRoleItem,
  IMyBriefSummary
} from '../services/OnyxSharePointService';

interface IOnyxFormState {
  ProductName: string;
  OneLineSummary: string;
  BriefVersion: string;
  Department: string;
  TargetDate: string;
  BriefPriority: string;
  Status: string;
  ProductDescription: string;
  CurrentProcess: string;
  CurrentSituation: string;
  RootCause: string;
  Opportunity: string;
  MainGoal: string;
  Objectives: string;
  SuccessDefinition: string;
  PrimaryUsers: string;
  SecondaryUsers: string;
  UserPersonas: string;
  InScopeItems: string;
  OutOfScopeItems: string;
  SecurityRequirements: string;
  PerformanceRequirements: string;
  ComplianceRequirements: string;
  Risks: string;
  Assumptions: string;
  KpisAndSuccessMetrics: string;
  FutureIdeas: string;
  OpenQuestions: string;
}

interface IMessageState {
  type: 'success' | 'error' | 'info';
  text: string;
}

const initialFormState: IOnyxFormState = {
  ProductName: '',
  OneLineSummary: '',
  BriefVersion: 'v1.0',
  Department: '',
  TargetDate: '',
  BriefPriority: 'Medium',
  Status: 'Draft',
  ProductDescription: '',
  CurrentProcess: '',
  CurrentSituation: '',
  RootCause: '',
  Opportunity: '',
  MainGoal: '',
  Objectives: '',
  SuccessDefinition: '',
  PrimaryUsers: '',
  SecondaryUsers: '',
  UserPersonas: '',
  InScopeItems: '',
  OutOfScopeItems: '',
  SecurityRequirements: '',
  PerformanceRequirements: '',
  ComplianceRequirements: '',
  Risks: '',
  Assumptions: '',
  KpisAndSuccessMetrics: '',
  FutureIdeas: '',
  OpenQuestions: ''
};

const initialFeatures: IOnyxFeatureItem[] = [
  { FeatureName: '', Priority: 'Must Have', Phase: 'Phase 1', Notes: '', SortOrder: 1 },
  { FeatureName: '', Priority: 'Nice To Have', Phase: 'Phase 2+', Notes: '', SortOrder: 2 }
];

const initialRoles: IOnyxRoleItem[] = [
  { RoleName: '', Permissions: '', SortOrder: 1 },
  { RoleName: '', Permissions: '', SortOrder: 2 }
];

const fieldSections: Array<{
  title: string;
  eyebrow: string;
  fields: Array<{
    key: keyof IOnyxFormState;
    label: string;
    hint: string;
    required?: boolean;
    rows?: number;
  }>;
}> = [
  {
    title: 'Product Overview',
    eyebrow: '01',
    fields: [
      { key: 'ProductDescription', label: 'What is this product?', hint: 'Describe what you are building and what it will allow people to do.', required: true },
      { key: 'CurrentProcess', label: 'What does it replace or improve on?', hint: 'Explain the current manual, spreadsheet, email, or tool-based process.' }
    ]
  },
  {
    title: 'The Problem',
    eyebrow: '02',
    fields: [
      { key: 'CurrentSituation', label: 'What is happening right now?', hint: 'Describe the slow, broken, risky, or frustrating current situation.', required: true },
      { key: 'RootCause', label: 'Why is this happening?', hint: 'Capture the root cause if it is already known.' },
      { key: 'Opportunity', label: 'What becomes possible if we fix this?', hint: 'Describe the business or user benefit unlocked by the product.' }
    ]
  },
  {
    title: 'Goals And Success',
    eyebrow: '03',
    fields: [
      { key: 'MainGoal', label: 'Main goal', hint: 'State the primary goal in plain language.', required: true },
      { key: 'Objectives', label: 'Objectives', hint: 'List the concrete outcomes this product must achieve.', required: true },
      { key: 'SuccessDefinition', label: 'Success definition', hint: 'Define what approval, adoption, or operational success looks like.', required: true }
    ]
  },
  {
    title: 'Users',
    eyebrow: '04',
    fields: [
      { key: 'PrimaryUsers', label: 'Primary users', hint: 'Who will use this most often and what do they need?', required: true },
      { key: 'SecondaryUsers', label: 'Secondary users', hint: 'Who else will interact with this occasionally?' },
      { key: 'UserPersonas', label: 'User personas', hint: 'Describe each user type, their needs, and friction points.', rows: 6 }
    ]
  },
  {
    title: 'Scope',
    eyebrow: '05',
    fields: [
      { key: 'InScopeItems', label: 'In scope', hint: 'List what this version will do.', required: true },
      { key: 'OutOfScopeItems', label: 'Out of scope', hint: 'List related items intentionally excluded from this version.', required: true }
    ]
  },
  {
    title: 'Other Requirements',
    eyebrow: '08',
    fields: [
      { key: 'SecurityRequirements', label: 'Security and access', hint: 'Access rules, data handling, audit, and visibility requirements.', required: true },
      { key: 'PerformanceRequirements', label: 'Speed and reliability', hint: 'Expected load times, availability, notification timing, and reliability.', required: true },
      { key: 'ComplianceRequirements', label: 'Rules and standards', hint: 'Legal, policy, branding, accessibility, and audit requirements.', required: true }
    ]
  },
  {
    title: 'Risks And Assumptions',
    eyebrow: '09',
    fields: [
      { key: 'Risks', label: 'Risks', hint: 'What could cause delay, pushback, data issues, or delivery problems?' },
      { key: 'Assumptions', label: 'Assumptions', hint: 'What are we assuming to be true before build starts?' }
    ]
  },
  {
    title: 'Measurement And Next Phases',
    eyebrow: '10',
    fields: [
      { key: 'KpisAndSuccessMetrics', label: 'KPIs and success metrics', hint: 'List measurable success targets.', required: true },
      { key: 'FutureIdeas', label: 'Future ideas', hint: 'Capture later-phase improvements that are not needed now.' },
      { key: 'OpenQuestions', label: 'Open questions', hint: 'Record decisions still needed before or during the build.' }
    ]
  }
];

const hasValue = (value: string): boolean => value.trim().length > 0;

const getTrackingStatus = (brief: IMyBriefSummary): 'Pending' | 'Complete' | 'Overdue' => {
  const status = (brief.Status || '').toLowerCase();

  if (status === 'approved' || status === 'rejected') {
    return 'Complete';
  }

  if (brief.TargetDate) {
    const targetDate = new Date(brief.TargetDate);
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    if (!isNaN(targetDate.getTime()) && targetDate < today) {
      return 'Overdue';
    }
  }

  return 'Pending';
};

const escapeCsvValue = (value?: string | number): string => {
  const text = value === undefined || value === null ? '' : String(value);
  return '"' + text.replace(/"/g, '""') + '"';
};

const getDecisionLabel = (status?: string): string => {
  if (status === 'Approved') {
    return 'Approved';
  }

  if (status === 'Rejected') {
    return 'Rejected';
  }

  if (status === 'Under Review') {
    return 'On hold';
  }

  return 'Awaiting decision';
};



const Onyx: React.FC<IOnyxProps> = (props: IOnyxProps) => {
  const opensFromLink = new URLSearchParams(window.location.search).has('onyxOpen') || new URLSearchParams(window.location.search).has('productBriefOpen');
  const queryBriefId = new URLSearchParams(window.location.search).get('briefId');
  const highlightBriefId = queryBriefId ? parseInt(queryBriefId, 10) : undefined;
  const [isAppOpen, setIsAppOpen] = React.useState<boolean>(opensFromLink);
  const [form, setForm] = React.useState<IOnyxFormState>(initialFormState);
  const [features, setFeatures] = React.useState<IOnyxFeatureItem[]>(initialFeatures);
  const [roles, setRoles] = React.useState<IOnyxRoleItem[]>(initialRoles);
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [message, setMessage] = React.useState<IMessageState | undefined>();
  const [myBriefs, setMyBriefs] = React.useState<IMyBriefSummary[]>([]);
  const [isLoadingBriefs, setIsLoadingBriefs] = React.useState<boolean>(false);
  const [decisionBriefId, setDecisionBriefId] = React.useState<number | undefined>();

  const stopPropagationRef = React.useCallback((el: HTMLTextAreaElement | null) => {
    if (el && !el.dataset.shielded) {
      el.dataset.shielded = 'true';
      const stopNative = (e: Event): void => {
        e.stopPropagation();
      };
      el.addEventListener('keydown', stopNative);
      el.addEventListener('keypress', stopNative);
      el.addEventListener('keyup', stopNative);
      el.addEventListener('input', stopNative);
      el.addEventListener('click', stopNative);
      el.addEventListener('mousedown', stopNative);
      el.addEventListener('focus', stopNative);
    }
  }, []);
  const service = React.useMemo(
    () => new OnyxSharePointService(props.context, props.dataSiteUrl, props.staffDirectorySiteUrl),
    [props.context, props.dataSiteUrl, props.staffDirectorySiteUrl]
  );
  const isAdmin = React.useMemo((): boolean => {
    const urlParams = new URLSearchParams(window.location.search);
    const hasAdminParam = urlParams.has('admin') || urlParams.get('admin') === 'true' || urlParams.has('onyxAdmin') || urlParams.has('productBriefAdmin');
    if (hasAdminParam) {
      return true;
    }

    const currentEmail = (props.currentUserEmail || '').trim().toLowerCase();

    if (!currentEmail) {
      return false;
    }

    return props.adminEmails
      .split(',')
      .map((email: string) => email.trim().toLowerCase())
      .filter((email: string) => email.length > 0)
      .indexOf(currentEmail) > -1;
  }, [props.adminEmails, props.currentUserEmail]);

  const loadMyBriefs = React.useCallback(async (): Promise<void> => {
    if (!isAdmin && !props.currentUserId) {
      return;
    }

    setIsLoadingBriefs(true);

    try {
      setMyBriefs(isAdmin ? await service.getAllBriefs() : await service.getMyBriefs(props.currentUserId));
    } catch (error) {
      setMessage({
        type: 'info',
        text: 'Product Brief could not load your saved records yet. You can still submit a new request.'
      });
    } finally {
      setIsLoadingBriefs(false);
    }
  }, [isAdmin, props.currentUserId, service]);

  React.useEffect(() => {
    let isMounted = true;

    const loadStaffProfile = async (): Promise<void> => {
      if (!props.currentUserEmail || !props.staffDirectoryListName) {
        return;
      }

      try {
        const profile = await service.getStaffProfile(props.currentUserEmail, props.staffDirectoryListName);

        if (!isMounted || !profile) {
          return;
        }

        setForm((current: IOnyxFormState) => ({
          ...current,
          Department: current.Department || profile.Department || ''
        }));
      } catch (error) {
        if (isMounted) {
          const lookupError = error instanceof Error ? error.message : 'Unknown lookup error.';
          setMessage({
            type: 'info',
            text: 'Staff directory lookup failed: ' + lookupError + '. Department can still be entered manually.'
          });
        }
      }
    };

    loadStaffProfile().catch(() => undefined);

    return () => {
      isMounted = false;
    };
  }, [props.currentUserEmail, props.staffDirectoryListName, service]);

  React.useEffect(() => {
    loadMyBriefs().catch(() => undefined);
  }, [loadMyBriefs]);

  React.useEffect(() => {
    if (highlightBriefId && myBriefs.length > 0) {
      const timer = setTimeout(() => {
        const element = document.getElementById(`record-item-${highlightBriefId}`);
        if (element) {
          element.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }, 500);
      return () => clearTimeout(timer);
    }
  }, [myBriefs, highlightBriefId]);



  const completedRequiredFields = [
    form.ProductName,
    form.OneLineSummary,
    form.ProductDescription,
    form.CurrentSituation,
    form.MainGoal,
    form.Objectives,
    form.SuccessDefinition,
    form.PrimaryUsers,
    form.InScopeItems,
    form.OutOfScopeItems,
    form.SecurityRequirements,
    form.PerformanceRequirements,
    form.ComplianceRequirements,
    form.KpisAndSuccessMetrics
  ].filter(hasValue).length + (props.currentUserId || props.currentUserEmail ? 1 : 0);

  const progress = Math.round((completedRequiredFields / 15) * 100);
  const exportRecords = (): void => {
    const headers = ['Id', 'ProductName', 'Department', 'Priority', 'WorkflowStatus', 'TrackingStatus', 'TargetDate', 'Modified'];
    const rows = myBriefs.map((brief: IMyBriefSummary) => [
      brief.Id,
      brief.ProductName || brief.Title || '',
      brief.Department || '',
      brief.BriefPriority || '',
      brief.Status || '',
      getTrackingStatus(brief),
      brief.TargetDate || '',
      brief.Modified || ''
    ]);
    const csv = [headers, ...rows]
      .map((row) => row.map((value) => escapeCsvValue(value)).join(','))
      .join('\r\n');
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');

    link.href = url;
    link.download = 'product-brief-submissions.csv';
    link.click();
    URL.revokeObjectURL(url);
  };

  const getTrackingClassName = (brief: IMyBriefSummary): string => {
    const trackingStatus = getTrackingStatus(brief);

    if (trackingStatus === 'Complete') {
      return styles.statusComplete;
    }

    if (trackingStatus === 'Overdue') {
      return styles.statusOverdue;
    }

    return styles.statusPending;
  };

  const decideBrief = async (briefId: number, status: 'Approved' | 'Under Review' | 'Rejected'): Promise<void> => {
    if (!isAdmin) {
      setMessage({ type: 'error', text: 'Only Product Brief admins can make a final PRD decision.' });
      return;
    }

    const commentInput = document.getElementById(`decision-comment-${briefId}`) as HTMLTextAreaElement;
    const decisionComment = commentInput ? commentInput.value.trim() : '';

    if ((status === 'Under Review' || status === 'Rejected') && !decisionComment) {
      setMessage({ type: 'error', text: 'Add an admin comment before placing a PRD on hold or rejecting it.' });
      return;
    }

    setDecisionBriefId(briefId);
    setMessage({ type: 'info', text: 'Updating PRD decision...' });

    try {
      const decisionDate = new Date().toISOString();
      await service.updateBriefStatus(briefId, status, decisionComment, props.currentUserId);
      setMyBriefs((current: IMyBriefSummary[]) => current.map((brief: IMyBriefSummary) => {
        return brief.Id === briefId ? { ...brief, DecisionComment: decisionComment, DecisionDate: decisionDate, Status: status } : brief;
      }));
      setMessage({ type: 'success', text: 'PRD decision saved as ' + getDecisionLabel(status) + '. The SharePoint item status has been updated.' });
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'The PRD decision could not be saved.';
      setMessage({ type: 'error', text: errorMessage });
    } finally {
      setDecisionBriefId(undefined);
    }
  };

  const updateForm = (key: keyof IOnyxFormState, value: string): void => {
    setForm((current: IOnyxFormState) => ({ ...current, [key]: value }));
  };

  const updateFeature = (index: number, key: keyof IOnyxFeatureItem, value: string): void => {
    setFeatures((current: IOnyxFeatureItem[]) => current.map((feature: IOnyxFeatureItem, featureIndex: number) => {
      return featureIndex === index ? { ...feature, [key]: value } : feature;
    }));
  };

  const updateRole = (index: number, key: keyof IOnyxRoleItem, value: string): void => {
    setRoles((current: IOnyxRoleItem[]) => current.map((role: IOnyxRoleItem, roleIndex: number) => {
      return roleIndex === index ? { ...role, [key]: value } : role;
    }));
  };

  const addFeature = (): void => {
    setFeatures((current: IOnyxFeatureItem[]) => current.concat([{
      FeatureName: '',
      Priority: 'Must Have',
      Phase: 'Phase 1',
      Notes: '',
      SortOrder: current.length + 1
    }]));
  };

  const addRole = (): void => {
    setRoles((current: IOnyxRoleItem[]) => current.concat([{
      RoleName: '',
      Permissions: '',
      SortOrder: current.length + 1
    }]));
  };

  const removeFeature = (index: number): void => {
    setFeatures((current: IOnyxFeatureItem[]) => current.filter((_: IOnyxFeatureItem, featureIndex: number) => featureIndex !== index));
  };

  const removeRole = (index: number): void => {
    setRoles((current: IOnyxRoleItem[]) => current.filter((_: IOnyxRoleItem, roleIndex: number) => roleIndex !== index));
  };

  const validate = (): string | undefined => {
    if (!hasValue(form.ProductName)) {
      return 'Product name is required.';
    }

    if (!hasValue(form.OneLineSummary)) {
      return 'One-line summary is required.';
    }

    if (!props.currentUserId && !props.currentUserEmail) {
      return 'Product Brief could not identify the current SharePoint user. Refresh the page and try again.';
    }

    return undefined;
  };

  const notifyAdminFlow = async (briefId: number, status: string): Promise<void> => {
    const pageUrl = window.location.href.split('?')[0];
    const itemUrl = pageUrl + '?productBriefOpen=true&admin=true&briefId=' + briefId;
    const response = await fetch(props.adminFlowUrl, {
      body: JSON.stringify({
        briefId: briefId,
        department: form.Department,
        itemUrl: itemUrl,
        priority: form.BriefPriority,
        productName: form.ProductName,
        status: status,
        submittedAt: new Date().toISOString(),
        submitterEmail: props.currentUserEmail,
        submitterName: props.userDisplayName,
        summary: form.OneLineSummary
      }),
      headers: {
        'Content-Type': 'application/json'
      },
      method: 'POST'
    });

    if (!response.ok) {
      throw new Error('The brief was saved, but the admin notification flow returned HTTP ' + response.status + '.');
    }
  };

  const saveBrief = async (status: string): Promise<void> => {
    const validationMessage = validate();

    if (validationMessage) {
      setMessage({ type: 'error', text: validationMessage });
      return;
    }

    setIsSaving(true);
    setMessage({ type: 'info', text: 'Saving product brief...' });

    const submitterId = props.currentUserId || await service.ensureUserId(props.currentUserEmail);

    const brief: IOnyxBriefItem = {
      ...form,
      Status: status,
      SubmitterId: submitterId,
      SubmissionDate: new Date().toISOString()
    };

    try {
      const result = await service.saveBrief(brief, features, roles);
      setForm((current: IOnyxFormState) => ({ ...current, Status: status }));
      if (status === 'Submitted' && props.adminFlowUrl) {
        await notifyAdminFlow(result.briefId, status);
      }
      setMessage({ type: 'success', text: 'Product brief saved. SharePoint item ID: ' + result.briefId + '.' });
      await loadMyBriefs();
    } catch (error) {
      const errorMessage = error instanceof Error ? error.message : 'The brief could not be saved.';
      setMessage({
        type: 'error',
        text: errorMessage.indexOf('does not exist at site') > -1
          ? 'Product Brief could not find the SharePoint lists at ' + props.dataSiteUrl + '. Open the web part property pane and set the SharePoint data site URL to the site where the three Onyx lists were created.'
          : errorMessage
      });
    } finally {
      setIsSaving(false);
    }
  };

  const renderTextArea = (
    key: keyof IOnyxFormState,
    label: string,
    hint: string,
    required?: boolean,
    rows?: number
  ): React.ReactElement => (
    <label className={styles.field}>
      <span>{label}{required ? ' *' : ''}</span>
      <small>{hint}</small>
      <textarea
        rows={rows || 4}
        value={form[key]}
        onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) => updateForm(key, event.currentTarget.value)}
      />
    </label>
  );

  const closePreview = (): void => {
    if (window.history.length > 1) {
      window.history.back();
      return;
    }

    window.location.href = props.context.pageContext.web.absoluteUrl;
  };


  const shieldFromSharePoint = (event: React.SyntheticEvent<HTMLElement>): void => {
    event.stopPropagation();
    event.nativeEvent.stopPropagation();
  };


  const blockFormSubmit = (event: React.FormEvent<HTMLElement>): void => {
    event.preventDefault();
    event.stopPropagation();
  };

  if (!isAppOpen) {
    return (
      <section className={styles.onyx}>
        <div className={styles.startShell}>
          <div className={styles.startCard}>
            <img src={require('../assets/kcc.png')} alt="Konstructum" className={styles.startLogo} />
            <h1 style={{ marginTop: '8px' }}>Product Brief</h1>
            <p>Submit a structured product brief to the ORBIT build team.</p>
            <button type="button" className={styles.primaryButton} onClick={() => setIsAppOpen(true)}>
              Open Product Brief
            </button>
          </div>
        </div>
      </section>
    );
  }

  return (
    <section className={`${styles.onyx} ${styles.onyxOpen}`}>
      <div className={styles.previewBar}>
        <span>Preview</span>
        <button type="button" onClick={closePreview}>Close preview</button>
      </div>

      <div className={styles.header}>
        <div className={styles.brandPanel}>
          <div className={styles.brandRow}>
            <img src={require('../assets/kcc.png')} alt="Konstructum" className={styles.logo} />
            <div>
              <h1 style={{ margin: 0, fontSize: 'clamp(32px, 4vw, 52px)' }}>Product Brief</h1>
            </div>
          </div>
          <p className={styles.subtitle}>Submit a structured product brief to the ORBIT build team.</p>
        </div>
        <div className={styles.progressCard}>
          <small>Brief readiness</small>
          <span>{progress}%</span>
          <div className={styles.progressTrack}>
            <div className={styles.progressFill} style={{ width: progress + '%' }} />
          </div>
          <small>Required fields completed</small>
        </div>
      </div>

      {message && (message.text.indexOf('saved as') > -1 || message.text.indexOf('saved') > -1 || message.text.indexOf('decision') > -1 || message.text.indexOf('PRD') > -1) && (
        <div className={`${styles.message} ${styles[message.type]}`}>{message.text}</div>
      )}

      {message && message.text.indexOf('saved as') === -1 && message.text.indexOf('saved') === -1 && message.text.indexOf('decision') === -1 && message.text.indexOf('PRD') === -1 && message.text.indexOf('SharePoint item ID') === -1 && (
        <div className={`${styles.message} ${styles[message.type]}`}>{message.text}</div>
      )}

      <section className={styles.recordsPanel}>
        <div>
          <h2>{isAdmin ? 'Submission Dashboard' : 'My Records'}</h2>
          <p>{isAdmin ? 'Admin view across recent Product Brief submissions.' : 'Drafts and submissions created by you.'}</p>
        </div>
        <div className={styles.recordsActions}>
          <button type="button" className={styles.secondaryButtonCompact} onClick={() => loadMyBriefs()} disabled={isLoadingBriefs}>
            {isLoadingBriefs ? 'Loading...' : 'Refresh'}
          </button>
          <button type="button" className={styles.secondaryButtonCompact} onClick={exportRecords} disabled={myBriefs.length === 0}>
            Export CSV
          </button>
        </div>
        <div className={styles.recordsList}>
          {myBriefs.length === 0 && <span className={styles.emptyState}>No saved records yet.</span>}
          {myBriefs.map((brief: IMyBriefSummary) => (
            <div
              id={`record-item-${brief.Id}`}
              style={brief.Id === highlightBriefId ? { border: '2px solid #0078d4', backgroundColor: 'rgba(0, 120, 212, 0.05)' } : undefined}
              className={styles.recordItem}
              key={brief.Id}
              onClick={isAdmin ? shieldFromSharePoint : undefined}
              onFocus={isAdmin ? shieldFromSharePoint : undefined}
              onMouseDown={isAdmin ? shieldFromSharePoint : undefined}
              onKeyDown={isAdmin ? shieldFromSharePoint : undefined}
              onSubmit={isAdmin ? blockFormSubmit : undefined}
            >
              <strong>{brief.ProductName || brief.Title || 'Untitled request'}</strong>
              <span>{brief.Department || 'No department'} · {brief.BriefPriority || 'No priority'} · #{brief.Id}</span>
              <span className={`${styles.statusPill} ${getTrackingClassName(brief)}`}>{getTrackingStatus(brief)}</span>
              <span>{brief.Status || 'Draft'} workflow status</span>
              {isAdmin && (
                <div
                  className={styles.decisionActions}
                  onClick={shieldFromSharePoint}
                  onFocus={shieldFromSharePoint}
                  onKeyDown={shieldFromSharePoint}
                  onMouseDown={shieldFromSharePoint}
                  onChange={shieldFromSharePoint}
                  onInput={shieldFromSharePoint}
                  onPaste={shieldFromSharePoint}
                  onSubmit={blockFormSubmit}
                >
                  <span className={styles.decisionState}>{getDecisionLabel(brief.Status)}</span>
                  <label className={styles.decisionComment}>
                    <span>Admin comment</span>
                    <textarea
                      id={`decision-comment-${brief.Id}`}
                      ref={stopPropagationRef}
                      disabled={decisionBriefId === brief.Id}
                      defaultValue={brief.DecisionComment || ''}
                      placeholder="Add decision context for the submitter"
                      rows={3}
                    />
                  </label>
                  <button type="button" className={`${styles.decisionButton} ${brief.Status === 'Approved' ? styles.selectedDecision : styles.approveButton}`} disabled={decisionBriefId === brief.Id || brief.Status === 'Approved' || brief.Status === 'Rejected'} onClick={(event: React.MouseEvent<HTMLButtonElement>) => {
                    event.preventDefault();
                    event.stopPropagation();
                    event.nativeEvent.stopPropagation();
                    decideBrief(brief.Id, 'Approved').catch(() => undefined);
                  }}>
                    Approve
                  </button>
                  <button type="button" className={`${styles.decisionButton} ${brief.Status === 'Under Review' ? styles.selectedDecision : styles.holdButton}`} disabled={decisionBriefId === brief.Id || brief.Status === 'Under Review' || brief.Status === 'Approved' || brief.Status === 'Rejected'} onClick={(event: React.MouseEvent<HTMLButtonElement>) => {
                    event.preventDefault();
                    event.stopPropagation();
                    event.nativeEvent.stopPropagation();
                    decideBrief(brief.Id, 'Under Review').catch(() => undefined);
                  }}>
                    Hold
                  </button>
                  <button type="button" className={`${styles.decisionButton} ${brief.Status === 'Rejected' ? styles.selectedDecision : styles.rejectButton}`} disabled={decisionBriefId === brief.Id || brief.Status === 'Rejected' || brief.Status === 'Approved'} onClick={(event: React.MouseEvent<HTMLButtonElement>) => {
                    event.preventDefault();
                    event.stopPropagation();
                    event.nativeEvent.stopPropagation();
                    decideBrief(brief.Id, 'Rejected').catch(() => undefined);
                  }}>
                    Reject
                  </button>
                </div>
              )}
            </div>
          ))}
        </div>
      </section>

      <div className={styles.metaGrid}>
        <label>
          <span>Product name *</span>
          <input value={form.ProductName} onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateForm('ProductName', event.currentTarget.value)} />
        </label>
        <label>
          <span>One-line summary *</span>
          <input value={form.OneLineSummary} onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateForm('OneLineSummary', event.currentTarget.value)} />
        </label>
        <label>
          <span>Version</span>
          <input value={form.BriefVersion} onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateForm('BriefVersion', event.currentTarget.value)} />
        </label>
        <label>
          <span>Department</span>
          <input value={form.Department} onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateForm('Department', event.currentTarget.value)} />
        </label>
        <label>
          <span>Target date</span>
          <input value={form.TargetDate} placeholder="Q3 2026" onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateForm('TargetDate', event.currentTarget.value)} />
        </label>
        <label>
          <span>Priority</span>
          <select value={form.BriefPriority} onChange={(event: React.ChangeEvent<HTMLSelectElement>) => updateForm('BriefPriority', event.currentTarget.value)}>
            <option>High</option>
            <option>Medium</option>
            <option>Low</option>
          </select>
        </label>
      </div>

      <div className={styles.peopleGrid}>
        <div className={styles.peopleCard}>
          <div className={styles.autoPerson}>
            <strong>{props.currentUserEmail || 'Current SharePoint user'}</strong>
          </div>
        </div>
      </div>

      {fieldSections.slice(0, 5).map((section) => (
        <section className={styles.section} key={section.title}>
          <div className={styles.sectionHead}>
            <span>{section.eyebrow}</span>
            <h2>{section.title}</h2>
          </div>
          <div className={styles.fieldGrid}>
            {section.fields.map((field) => renderTextArea(field.key, field.label, field.hint, field.required, field.rows))}
          </div>
        </section>
      ))}

      <section className={styles.section}>
        <div className={styles.sectionHead}>
          <span>06</span>
          <h2>Features</h2>
        </div>
        <div className={styles.featureTable}>
          <div className={styles.featureHeader}>
            <span>Feature</span>
            <span>Priority</span>
            <span>Phase</span>
            <span>Notes</span>
            <span />
          </div>
          {features.map((feature: IOnyxFeatureItem, index: number) => (
            <div className={styles.featureRow} key={index}>
              <input value={feature.FeatureName} onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateFeature(index, 'FeatureName', event.currentTarget.value)} />
              <select value={feature.Priority} onChange={(event: React.ChangeEvent<HTMLSelectElement>) => updateFeature(index, 'Priority', event.currentTarget.value)}>
                <option>Must Have</option>
                <option>Nice To Have</option>
              </select>
              <select value={feature.Phase} onChange={(event: React.ChangeEvent<HTMLSelectElement>) => updateFeature(index, 'Phase', event.currentTarget.value)}>
                <option>Phase 1</option>
                <option>Phase 2+</option>
                <option>Future</option>
              </select>
              <input value={feature.Notes} onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateFeature(index, 'Notes', event.currentTarget.value)} />
              <button type="button" className={styles.iconButton} onClick={() => removeFeature(index)} aria-label="Remove feature">x</button>
            </div>
          ))}
        </div>
        <button type="button" className={styles.secondaryButton} onClick={addFeature}>Add feature</button>
      </section>

      <section className={styles.section}>
        <div className={styles.sectionHead}>
          <span>07</span>
          <h2>User Roles</h2>
        </div>
        <div className={styles.rolesGrid}>
          {roles.map((role: IOnyxRoleItem, index: number) => (
            <div className={styles.roleCard} key={index}>
              <div className={styles.roleHead}>
                <input value={role.RoleName} placeholder="Role name" onChange={(event: React.ChangeEvent<HTMLInputElement>) => updateRole(index, 'RoleName', event.currentTarget.value)} />
                <button type="button" className={styles.iconButton} onClick={() => removeRole(index)} aria-label="Remove role">x</button>
              </div>
              <textarea value={role.Permissions} placeholder="What can this role see or do?" onChange={(event: React.ChangeEvent<HTMLTextAreaElement>) => updateRole(index, 'Permissions', event.currentTarget.value)} />
            </div>
          ))}
        </div>
        <button type="button" className={styles.secondaryButton} onClick={addRole}>Add role</button>
      </section>

      {fieldSections.slice(5).map((section) => (
        <section className={styles.section} key={section.title}>
          <div className={styles.sectionHead}>
            <span>{section.eyebrow}</span>
            <h2>{section.title}</h2>
          </div>
          <div className={styles.fieldGrid}>
            {section.fields.map((field) => renderTextArea(field.key, field.label, field.hint, field.required, field.rows))}
          </div>
        </section>
      ))}

      <div className={styles.actionBar}>
        <div style={{ width: '100%' }}>
          <div style={{ display: 'flex', gap: '10px', justifyContent: 'flex-end' }}>
            <button type="button" disabled={isSaving} onClick={() => saveBrief('Draft')}>Save draft</button>
            <button type="button" disabled={isSaving} className={styles.primaryButton} onClick={() => saveBrief('Submitted')}>Submit brief</button>
          </div>
          {message && (message.text.indexOf('saved as') === -1 && message.text.indexOf('SharePoint item ID') > -1 || message.text.indexOf('Product brief saved') > -1) && (
            <div className={`${styles.message} ${styles[message.type]}`} style={{ marginTop: '12px', width: '100%', textAlign: 'right' }}>
              {message.text}
            </div>
          )}
        </div>
      </div>
    </section>
  );
};

export default Onyx;
