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



const loadHtml2Pdf = (): Promise<any> => {
  return new Promise((resolve, reject) => {
    if ((window as any).html2pdf) {
      resolve((window as any).html2pdf);
      return;
    }
    const script = document.createElement('script');
    script.src = 'https://cdnjs.cloudflare.com/ajax/libs/html2pdf.js/0.10.1/html2pdf.bundle.min.js';
    script.onload = () => resolve((window as any).html2pdf);
    script.onerror = () => reject(new Error('Failed to load html2pdf script from CDN.'));
    document.head.appendChild(script);
  });
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
  const [selectedDetail, setSelectedDetail] = React.useState<{ brief: IOnyxBriefItem; features: IOnyxFeatureItem[]; roles: IOnyxRoleItem[] } | undefined>();
  const [isLoadingDetails, setIsLoadingDetails] = React.useState<boolean>(false);
  const [resolvedUserId, setResolvedUserId] = React.useState<number | undefined>(props.currentUserId);
  const [activeTab, setActiveTab] = React.useState<'Submissions' | 'Drafts'>('Submissions');

  const exportToPdf = React.useCallback(async () => {
    if (!selectedDetail) return;
    const { brief, features, roles } = selectedDetail;

    const escapeHtml = (text?: string): string => {
      if (!text) return 'N/A';
      return text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#039;')
        .replace(/\n/g, '<br />');
    };

    let featuresHtml = '';
    if (features.length === 0) {
      featuresHtml = '<p>No features specified.</p>';
    } else {
      featuresHtml = '<table style="width:100%; border-collapse:collapse; margin-top:8px;"><thead><tr><th style="width:30%; border:1px solid #e5e5e5; padding:6px 10px; text-align:left; font-size:12px; background:#fafafa; font-weight:600; text-transform:uppercase; letter-spacing:0.3px;">Feature</th><th style="width:15%; border:1px solid #e5e5e5; padding:6px 10px; text-align:left; font-size:12px; background:#fafafa; font-weight:600; text-transform:uppercase; letter-spacing:0.3px;">Priority</th><th style="width:15%; border:1px solid #e5e5e5; padding:6px 10px; text-align:left; font-size:12px; background:#fafafa; font-weight:600; text-transform:uppercase; letter-spacing:0.3px;">Phase</th><th style="width:40%; border:1px solid #e5e5e5; padding:6px 10px; text-align:left; font-size:12px; background:#fafafa; font-weight:600; text-transform:uppercase; letter-spacing:0.3px;">Notes</th></tr></thead><tbody>';
      for (const f of features) {
        featuresHtml += '<tr><td style="border:1px solid #e5e5e5; padding:6px 10px; font-size:12px; vertical-align:top;"><strong>' + escapeHtml(f.FeatureName) + '</strong></td><td style="border:1px solid #e5e5e5; padding:6px 10px; font-size:12px; vertical-align:top;">' + escapeHtml(f.Priority) + '</td><td style="border:1px solid #e5e5e5; padding:6px 10px; font-size:12px; vertical-align:top;">' + escapeHtml(f.Phase) + '</td><td style="border:1px solid #e5e5e5; padding:6px 10px; font-size:12px; vertical-align:top;">' + escapeHtml(f.Notes) + '</td></tr>';
      }
      featuresHtml += '</tbody></table>';
    }

    let rolesHtml = '';
    if (roles.length === 0) {
      rolesHtml = '<p>No user roles specified.</p>';
    } else {
      rolesHtml = '<table style="width:100%; border-collapse:collapse; margin-top:8px;"><thead><tr><th style="width:30%; border:1px solid #e5e5e5; padding:6px 10px; text-align:left; font-size:12px; background:#fafafa; font-weight:600; text-transform:uppercase; letter-spacing:0.3px;">Role</th><th style="width:70%; border:1px solid #e5e5e5; padding:6px 10px; text-align:left; font-size:12px; background:#fafafa; font-weight:600; text-transform:uppercase; letter-spacing:0.3px;">Permissions / Capabilities</th></tr></thead><tbody>';
      for (const r of roles) {
        rolesHtml += '<tr><td style="border:1px solid #e5e5e5; padding:6px 10px; font-size:12px; vertical-align:top;"><strong>' + escapeHtml(r.RoleName) + '</strong></td><td style="border:1px solid #e5e5e5; padding:6px 10px; font-size:12px; vertical-align:top;">' + escapeHtml(r.Permissions) + '</td></tr>';
      }
      rolesHtml += '</tbody></table>';
    }

    const htmlContent = `
      <div style="font-family: system-ui, -apple-system, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1a1a1a; line-height: 1.5; padding: 20px; background: #ffffff;">
        <div style="border-bottom: 2px solid #a8201a; padding-bottom: 12px; margin-bottom: 20px; display: flex; justify-content: space-between; align-items: flex-end;">
          <div>
            <div style="font-size: 10px; text-transform: uppercase; letter-spacing: 1px; color: #a8201a; font-weight: 600; margin-bottom: 4px;">Onyx Request Archive</div>
            <h1 style="margin: 0; font-size: 24px; font-weight: 700; color: #1a1a1a; letter-spacing: -0.5px;">${escapeHtml(brief.ProductName)}</h1>
          </div>
          <div style="text-align: right; font-size: 11px; color: #666666;">
            ID: #${(brief as any).Id}<br />
            Date: ${new Date().toLocaleDateString()}
          </div>
        </div>

        <div style="display: grid; grid-template-columns: repeat(3, 1fr); gap: 12px; margin-bottom: 24px; background: #fafafa; border: 1px solid #e5e5e5; padding: 12px; font-size: 12px;">
          <div style="margin-bottom: 8px;">
            <strong style="color: #444444; display: block; font-weight: 600; margin-bottom: 2px;">Department</strong>
            ${escapeHtml(brief.Department)}
          </div>
          <div style="margin-bottom: 8px;">
            <strong style="color: #444444; display: block; font-weight: 600; margin-bottom: 2px;">Priority</strong>
            ${escapeHtml(brief.BriefPriority)}
          </div>
          <div style="margin-bottom: 8px;">
            <strong style="color: #444444; display: block; font-weight: 600; margin-bottom: 2px;">Target Date</strong>
            ${escapeHtml(brief.TargetDate)}
          </div>
          <div>
            <strong style="color: #444444; display: block; font-weight: 600; margin-bottom: 2px;">Version</strong>
            ${escapeHtml(brief.BriefVersion)}
          </div>
          <div>
            <strong style="color: #444444; display: block; font-weight: 600; margin-bottom: 2px;">Workflow Status</strong>
            ${escapeHtml(brief.Status)}
          </div>
          <div>
            <strong style="color: #444444; display: block; font-weight: 600; margin-bottom: 2px;">One-Line Summary</strong>
            ${escapeHtml(brief.OneLineSummary || (brief as any).Title)}
          </div>
        </div>

        <div style="margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">01 Product Overview</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Product Description</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.ProductDescription)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Current Process / What it Replaces</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.CurrentProcess)}</div>
          </div>
        </div>

        <div style="margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">02 The Problem</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Current Situation</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.CurrentSituation)}</div>
          </div>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Root Cause</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.RootCause)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Opportunity</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.Opportunity)}</div>
          </div>
        </div>

        <div style="margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">03 Goals & Success</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Main Goal</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.MainGoal)}</div>
          </div>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Objectives</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.Objectives)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Success Definition</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.SuccessDefinition)}</div>
          </div>
        </div>

        <div style="page-break-before: always; padding-top: 10px; margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">04 Target Users</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Primary Users</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.PrimaryUsers)}</div>
          </div>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Secondary Users</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.SecondaryUsers)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">User Personas</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.UserPersonas)}</div>
          </div>
        </div>

        <div style="margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">05 Scope</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">In Scope Items</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.InScopeItems)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Out of Scope Items</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.OutOfScopeItems)}</div>
          </div>
        </div>

        <div style="margin-bottom: 20px; page-break-inside: avoid;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">06 Proposed Features</h2>
          ${featuresHtml}
        </div>

        <div style="margin-bottom: 20px; page-break-inside: avoid;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">07 User Roles & Permissions</h2>
          ${rolesHtml}
        </div>

        <div style="page-break-before: always; padding-top: 10px; margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">08 Other Requirements</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Security & Access</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.SecurityRequirements)}</div>
          </div>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Performance Requirements</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.PerformanceRequirements)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Compliance Requirements</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.ComplianceRequirements)}</div>
          </div>
        </div>

        <div style="margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">09 Risks & Assumptions</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Risks</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.Risks)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Assumptions</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.Assumptions)}</div>
          </div>
        </div>

        <div style="margin-bottom: 20px;">
          <h2 style="font-size: 13px; text-transform: uppercase; letter-spacing: 0.5px; color: #a8201a; border-bottom: 1px solid #e5e5e5; padding-bottom: 4px; margin-bottom: 10px; font-weight: 600;">10 Measurement & Future</h2>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">KPIs and Success Metrics</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.KpisAndSuccessMetrics)}</div>
          </div>
          <div style="margin-bottom: 10px;">
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Future Ideas</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.FutureIdeas)}</div>
          </div>
          <div>
            <strong style="display: block; font-size: 11px; color: #555555; text-transform: uppercase; letter-spacing: 0.3px; margin-bottom: 2px;">Open Questions</strong>
            <div style="font-size: 13px; white-space: pre-wrap;">${escapeHtml(brief.OpenQuestions)}</div>
          </div>
        </div>

        <div style="margin-top: 30px; border-top: 1px solid #e5e5e5; padding-top: 10px; font-size: 10px; color: #666666; text-align: center;">
          Confidential - Onyx Product Development Brief System
        </div>
      </div>
    `;

    setMessage({ type: 'info', text: 'Generating PDF...' });

    try {
      const html2pdf = await loadHtml2Pdf();
      const element = document.createElement('div');
      element.style.position = 'absolute';
      element.style.left = '-9999px';
      element.style.top = '-9999px';
      element.style.width = '800px';
      element.innerHTML = htmlContent;
      document.body.appendChild(element);

      const opt = {
        margin:       15,
        filename:     `${brief.ProductName || 'Onyx_Brief'}_#${(brief as any).Id || 'Draft'}.pdf`,
        image:        { type: 'jpeg', quality: 0.98 },
        html2canvas:  { scale: 2, useCORS: true, logging: false },
        jsPDF:        { unit: 'mm', format: 'a4', orientation: 'portrait' }
      };

      await html2pdf().from(element).set(opt).save();
      document.body.removeChild(element);
      setMessage({ type: 'success', text: 'PDF downloaded successfully!' });
    } catch (err) {
      console.error('html2pdf error, falling back to native print window:', err);
      // Fallback: popup window print
      const printWindow = window.open('', '_blank');
      if (printWindow) {
        printWindow.document.open();
        printWindow.document.write(`
          <!DOCTYPE html>
          <html>
          <head>
            <meta charset="utf-8" />
            <title>Onyx Product Brief - ${escapeHtml(brief.ProductName)}</title>
            <style>
              body { font-family: system-ui, sans-serif; margin: 40px; }
            </style>
          </head>
          <body>
            ${htmlContent}
            <script>
              window.onload = function() {
                setTimeout(function() {
                  window.print();
                }, 300);
              };
            </script>
          </body>
          </html>
        `);
        printWindow.document.close();
        setMessage(undefined);
      } else {
        setMessage({ type: 'error', text: 'Failed to open print window. Please allow popups.' });
      }
    }
  }, [selectedDetail]);

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

  React.useEffect(() => {
    let isMounted = true;
    const resolveUserId = async (): Promise<void> => {
      if (props.currentUserId) {
        setResolvedUserId(props.currentUserId);
        return;
      }
      if (!props.currentUserEmail) {
        return;
      }
      try {
        const id = await service.ensureUserId(props.currentUserEmail);
        if (isMounted && id) {
          setResolvedUserId(id);
        }
      } catch (error) {
        // ignore
      }
    };
    resolveUserId().catch(() => undefined);
    return () => {
      isMounted = false;
    };
  }, [props.currentUserId, props.currentUserEmail, service]);

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
    const userId = resolvedUserId || props.currentUserId;
    if (!isAdmin && !userId) {
      return;
    }

    setIsLoadingBriefs(true);

    try {
      setMyBriefs(isAdmin ? await service.getAllBriefs() : await service.getMyBriefs(userId));
    } catch (error) {
      setMessage({
        type: 'info',
        text: 'Product Brief could not load your saved records yet. You can still submit a new request.'
      });
    } finally {
      setIsLoadingBriefs(false);
    }
  }, [isAdmin, resolvedUserId, props.currentUserId, service]);

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
      await service.updateBriefStatus(briefId, status, decisionComment, resolvedUserId || props.currentUserId);
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

  const loadBriefDetails = async (briefId: number): Promise<void> => {
    setIsLoadingDetails(true);
    try {
      const details = await service.getBriefDetails(briefId);
      setSelectedDetail(details);
    } catch (error) {
      setMessage({ type: 'error', text: 'Failed to load details for brief #' + briefId });
    } finally {
      setIsLoadingDetails(false);
    }
  };

  React.useEffect(() => {
    if (highlightBriefId) {
      const autoLoad = async (): Promise<void> => {
        setIsLoadingDetails(true);
        try {
          const details = await service.getBriefDetails(highlightBriefId);
          if ((details.brief.Status || '').toLowerCase() === 'draft') {
            setForm({
              ProductName: details.brief.ProductName || '',
              OneLineSummary: details.brief.OneLineSummary || '',
              BriefVersion: details.brief.BriefVersion || 'v1.0',
              Department: details.brief.Department || '',
              TargetDate: details.brief.TargetDate || '',
              BriefPriority: details.brief.BriefPriority || 'Medium',
              Status: 'Draft',
              ProductDescription: details.brief.ProductDescription || '',
              CurrentProcess: details.brief.CurrentProcess || '',
              CurrentSituation: details.brief.CurrentSituation || '',
              RootCause: details.brief.RootCause || '',
              Opportunity: details.brief.Opportunity || '',
              MainGoal: details.brief.MainGoal || '',
              Objectives: details.brief.Objectives || '',
              SuccessDefinition: details.brief.SuccessDefinition || '',
              PrimaryUsers: details.brief.PrimaryUsers || '',
              SecondaryUsers: details.brief.SecondaryUsers || '',
              UserPersonas: details.brief.UserPersonas || '',
              InScopeItems: details.brief.InScopeItems || '',
              OutOfScopeItems: details.brief.OutOfScopeItems || '',
              SecurityRequirements: details.brief.SecurityRequirements || '',
              PerformanceRequirements: details.brief.PerformanceRequirements || '',
              ComplianceRequirements: details.brief.ComplianceRequirements || '',
              Risks: details.brief.Risks || '',
              Assumptions: details.brief.Assumptions || '',
              KpisAndSuccessMetrics: details.brief.KpisAndSuccessMetrics || '',
              FutureIdeas: details.brief.FutureIdeas || '',
              OpenQuestions: details.brief.OpenQuestions || ''
            });
            setFeatures(details.features.length > 0 ? details.features : initialFeatures);
            setRoles(details.roles.length > 0 ? details.roles : initialRoles);
          } else {
            setSelectedDetail(details);
          }
          setIsAppOpen(true);
        } catch (error) {
          // ignore
        } finally {
          setIsLoadingDetails(false);
        }
      };
      autoLoad().catch(() => undefined);
    }
  }, [highlightBriefId, service]);

  const resumeDraft = async (briefId: number): Promise<void> => {
    setIsLoadingDetails(true);
    try {
      const details = await service.getBriefDetails(briefId);
      setForm({
        ProductName: details.brief.ProductName || '',
        OneLineSummary: details.brief.OneLineSummary || '',
        BriefVersion: details.brief.BriefVersion || 'v1.0',
        Department: details.brief.Department || '',
        TargetDate: details.brief.TargetDate || '',
        BriefPriority: details.brief.BriefPriority || 'Medium',
        Status: 'Draft',
        ProductDescription: details.brief.ProductDescription || '',
        CurrentProcess: details.brief.CurrentProcess || '',
        CurrentSituation: details.brief.CurrentSituation || '',
        RootCause: details.brief.RootCause || '',
        Opportunity: details.brief.Opportunity || '',
        MainGoal: details.brief.MainGoal || '',
        Objectives: details.brief.Objectives || '',
        SuccessDefinition: details.brief.SuccessDefinition || '',
        PrimaryUsers: details.brief.PrimaryUsers || '',
        SecondaryUsers: details.brief.SecondaryUsers || '',
        UserPersonas: details.brief.UserPersonas || '',
        InScopeItems: details.brief.InScopeItems || '',
        OutOfScopeItems: details.brief.OutOfScopeItems || '',
        SecurityRequirements: details.brief.SecurityRequirements || '',
        PerformanceRequirements: details.brief.PerformanceRequirements || '',
        ComplianceRequirements: details.brief.ComplianceRequirements || '',
        Risks: details.brief.Risks || '',
        Assumptions: details.brief.Assumptions || '',
        KpisAndSuccessMetrics: details.brief.KpisAndSuccessMetrics || '',
        FutureIdeas: details.brief.FutureIdeas || '',
        OpenQuestions: details.brief.OpenQuestions || ''
      });
      setFeatures(details.features.length > 0 ? details.features : initialFeatures);
      setRoles(details.roles.length > 0 ? details.roles : initialRoles);
      
      setMessage({ type: 'success', text: 'Draft ' + (details.brief.ProductName || '#' + briefId) + ' loaded! Scroll down to edit and submit.' });
      
      const formElement = document.querySelector(`.${styles.metaGrid}`);
      if (formElement) {
        formElement.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    } catch (error) {
      setMessage({ type: 'error', text: 'Failed to load draft details.' });
    } finally {
      setIsLoadingDetails(false);
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
    const itemUrl = pageUrl + '?productBriefOpen=true&admin=true&briefId=' + briefId + '#onyx-app-root';
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
      <section id="onyx-app-root" className={styles.onyx}>
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
    <section id="onyx-app-root" className={`${styles.onyx} ${styles.onyxOpen}`}>
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
        <div className={styles.dashboardHeaderColumn}>
          <h2>{isAdmin ? 'Submission Dashboard' : 'My Records'}</h2>
          <p>{isAdmin ? 'Admin view across recent Product Brief submissions.' : 'Drafts and submissions created by you.'}</p>
          {!isAdmin && (
            <div className={styles.tabBar} style={{ marginTop: '16px' }}>
              <button
                type="button"
                className={`${styles.tabItem} ${activeTab === 'Submissions' ? styles.tabItemActive : ''}`}
                onClick={() => setActiveTab('Submissions')}
              >
                Submissions ({myBriefs.filter((b) => (b.Status || '').toLowerCase() !== 'draft').length})
              </button>
              <button
                type="button"
                className={`${styles.tabItem} ${activeTab === 'Drafts' ? styles.tabItemActive : ''}`}
                onClick={() => setActiveTab('Drafts')}
              >
                Drafts ({myBriefs.filter((b) => (b.Status || '').toLowerCase() === 'draft').length})
              </button>
            </div>
          )}
          <div className={styles.recordsActions}>
            <button type="button" className={styles.secondaryButtonCompact} onClick={() => loadMyBriefs()} disabled={isLoadingBriefs}>
              {isLoadingBriefs ? 'Loading...' : 'Refresh'}
            </button>
            <button type="button" className={styles.secondaryButtonCompact} onClick={exportRecords} disabled={myBriefs.length === 0}>
              Export CSV
            </button>
          </div>
        </div>
        <div className={styles.recordsList}>
          {(!isAdmin ? myBriefs.filter((b) => activeTab === 'Submissions' ? (b.Status || '').toLowerCase() !== 'draft' : (b.Status || '').toLowerCase() === 'draft') : myBriefs).length === 0 && (
            <span className={styles.emptyState}>
              {activeTab === 'Drafts' ? 'No saved drafts found.' : 'No submitted records yet.'}
            </span>
          )}
          {(!isAdmin ? myBriefs.filter((b) => activeTab === 'Submissions' ? (b.Status || '').toLowerCase() !== 'draft' : (b.Status || '').toLowerCase() === 'draft') : myBriefs).map((brief: IMyBriefSummary) => (
            <div
              id={`record-item-${brief.Id}`}
              style={brief.Id === highlightBriefId ? { borderColor: '#1a1a1a', backgroundColor: '#fafafa', borderWidth: '2px' } : undefined}
              className={`${styles.recordItem} ${selectedDetail && selectedDetail.brief.ProductName === brief.ProductName ? styles.recordItemActive : ''}`}
              key={brief.Id}
              onClick={(event: React.MouseEvent<HTMLDivElement>) => {
                if (isAdmin) {
                  shieldFromSharePoint(event);
                }
                if ((brief.Status || '').toLowerCase() === 'draft') {
                  resumeDraft(brief.Id).catch(() => undefined);
                } else {
                  loadBriefDetails(brief.Id).catch(() => undefined);
                }
              }}
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

      {(isLoadingDetails || selectedDetail) && (
        <div className={styles.detailsTakeover} onClick={() => setSelectedDetail(undefined)}>
          <div className={styles.detailsDocument} onClick={shieldFromSharePoint}>
            {isLoadingDetails ? (
              <div className={styles.detailsLoading}>
                <span className={styles.emptyState}>Loading specification details...</span>
              </div>
            ) : (
              selectedDetail && (
                <>
                  <div className={styles.detailsHeader}>
                    <div>
                      <div className={styles.documentBadge}>Read-Only Request Archive</div>
                      <h3>{selectedDetail.brief.ProductName || 'Untitled brief'}</h3>
                    </div>
                    <div style={{ display: 'flex', gap: '8px', alignItems: 'center' }}>
                      <button type="button" className={styles.secondaryButtonCompact} onClick={exportToPdf}>
                        Export PDF
                      </button>
                      <button type="button" className={styles.secondaryButtonCompact} onClick={() => setSelectedDetail(undefined)}>
                        Close Brief
                      </button>
                    </div>
                  </div>

                  <div className={styles.documentArchiveInfo}>
                    <span>Locked Record</span>
                    <p>This is a read-only document representing the brief of the request submitted by staff. It cannot be modified.</p>
                  </div>
                  
                  <div className={styles.detailsContent}>
                    <div className={styles.detailsMeta}>
                      <span><strong>Department:</strong> {selectedDetail.brief.Department || 'N/A'}</span>
                      <span><strong>Priority:</strong> {selectedDetail.brief.BriefPriority || 'N/A'}</span>
                      <span><strong>Target Date:</strong> {selectedDetail.brief.TargetDate || 'N/A'}</span>
                      <span><strong>Version:</strong> {selectedDetail.brief.BriefVersion || 'N/A'}</span>
                      <span><strong>Status:</strong> {selectedDetail.brief.Status || 'N/A'}</span>
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>Overview</h4>
                      <p><strong>Description:</strong> {selectedDetail.brief.ProductDescription || 'N/A'}</p>
                      <p><strong>Replaces/Improves:</strong> {selectedDetail.brief.CurrentProcess || 'N/A'}</p>
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>The Problem</h4>
                      <p><strong>Current Situation:</strong> {selectedDetail.brief.CurrentSituation || 'N/A'}</p>
                      <p><strong>Root Cause:</strong> {selectedDetail.brief.RootCause || 'N/A'}</p>
                      <p><strong>Opportunity:</strong> {selectedDetail.brief.Opportunity || 'N/A'}</p>
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>Goals & Success Definition</h4>
                      <p><strong>Main Goal:</strong> {selectedDetail.brief.MainGoal || 'N/A'}</p>
                      <p><strong>Objectives:</strong> {selectedDetail.brief.Objectives || 'N/A'}</p>
                      <p><strong>Success Definition:</strong> {selectedDetail.brief.SuccessDefinition || 'N/A'}</p>
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>Target Users & Scope</h4>
                      <p><strong>Primary Users:</strong> {selectedDetail.brief.PrimaryUsers || 'N/A'}</p>
                      <p><strong>Secondary Users:</strong> {selectedDetail.brief.SecondaryUsers || 'N/A'}</p>
                      <p><strong>User Personas:</strong> {selectedDetail.brief.UserPersonas || 'N/A'}</p>
                      <p><strong>In Scope:</strong> {selectedDetail.brief.InScopeItems || 'N/A'}</p>
                      <p><strong>Out of Scope:</strong> {selectedDetail.brief.OutOfScopeItems || 'N/A'}</p>
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>Proposed Features</h4>
                      {selectedDetail.features.length === 0 ? <p>No features specified.</p> : (
                        <div className={styles.detailsList}>
                          {selectedDetail.features.map((f, i) => (
                            <div key={i} className={styles.detailsListItem}>
                              <strong>{f.FeatureName}</strong> ({f.Priority} · {f.Phase})
                              {f.Notes && <p className={styles.detailsListItemNotes}>{f.Notes}</p>}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>User Roles & Permissions</h4>
                      {selectedDetail.roles.length === 0 ? <p>No user roles specified.</p> : (
                        <div className={styles.detailsList}>
                          {selectedDetail.roles.map((r, i) => (
                            <div key={i} className={styles.detailsListItem}>
                              <strong>{r.RoleName}</strong>
                              {r.Permissions && <p className={styles.detailsListItemNotes}>{r.Permissions}</p>}
                            </div>
                          ))}
                        </div>
                      )}
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>Non-Functional Requirements</h4>
                      <p><strong>Security:</strong> {selectedDetail.brief.SecurityRequirements || 'N/A'}</p>
                      <p><strong>Performance:</strong> {selectedDetail.brief.PerformanceRequirements || 'N/A'}</p>
                      <p><strong>Compliance:</strong> {selectedDetail.brief.ComplianceRequirements || 'N/A'}</p>
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>Risks & Assumptions</h4>
                      <p><strong>Risks:</strong> {selectedDetail.brief.Risks || 'N/A'}</p>
                      <p><strong>Assumptions:</strong> {selectedDetail.brief.Assumptions || 'N/A'}</p>
                    </div>

                    <div className={styles.detailsSection}>
                      <h4>Measurement & Future Phases</h4>
                      <p><strong>KPIs:</strong> {selectedDetail.brief.KpisAndSuccessMetrics || 'N/A'}</p>
                      <p><strong>Future Ideas:</strong> {selectedDetail.brief.FutureIdeas || 'N/A'}</p>
                      <p><strong>Open Questions:</strong> {selectedDetail.brief.OpenQuestions || 'N/A'}</p>
                    </div>
                  </div>
                </>
              )
            )}
          </div>
        </div>
      )}

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
