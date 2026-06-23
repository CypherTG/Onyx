import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, SPFx, spfi } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/fields';
import '@pnp/sp/items';
import '@pnp/sp/site-users/web';

export interface IOnyxFeatureItem {
  FeatureName: string;
  Priority: string;
  Phase: string;
  Notes: string;
  SortOrder: number;
}

export interface IOnyxRoleItem {
  RoleName: string;
  Permissions: string;
  SortOrder: number;
}

export interface IOnyxBriefItem {
  ProductName: string;
  OneLineSummary: string;
  BriefVersion: string;
  Department: string;
  TargetDate: string;
  BriefPriority: string;
  SubmitterId?: number;
  ReviewerId?: number;
  SubmissionDate: string;
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

export interface ISaveBriefResult {
  briefId: number;
}

export interface IMyBriefSummary {
  BriefPriority?: string;
  Created?: string;
  DecisionComment?: string;
  DecisionDate?: string;
  Department?: string;
  Id: number;
  Modified?: string;
  ProductName?: string;
  Status?: string;
  TargetDate?: string;
  Title?: string;
}

export interface IStaffDirectoryProfile {
  Department?: string;
  EmailAddress?: string;
  FullName?: string;
  LineManager?: string;
  Title?: string;
}

interface ISharePointFieldInfo {
  InternalName: string;
  Title: string;
}

const productBriefsList = 'OnyxProductBriefs';
const featuresList = 'OnyxBriefFeatures';
const rolesList = 'OnyxBriefRoles';

export default class OnyxSharePointService {
  private readonly _sp: SPFI;
  private readonly _staffSp: SPFI;

  public constructor(context: WebPartContext, dataSiteUrl: string, staffDirectorySiteUrl: string) {
    const normalizedSiteUrl = dataSiteUrl.trim().replace(/\/$/, '');
    const normalizedStaffSiteUrl = staffDirectorySiteUrl.trim().replace(/\/$/, '') || normalizedSiteUrl;
    this._sp = normalizedSiteUrl ? spfi(normalizedSiteUrl).using(SPFx(context)) : spfi().using(SPFx(context));
    this._staffSp = normalizedStaffSiteUrl ? spfi(normalizedStaffSiteUrl).using(SPFx(context)) : this._sp;
  }

  public async saveBrief(
    brief: IOnyxBriefItem,
    features: IOnyxFeatureItem[],
    roles: IOnyxRoleItem[]
  ): Promise<ISaveBriefResult> {
    const briefAddResult = await this._sp.web.lists.getByTitle(productBriefsList).items.add({
      ...brief,
      Title: brief.ProductName
    });

    const briefId: number = briefAddResult.data.Id;
    const featureCreates: Promise<unknown>[] = [];
    const roleCreates: Promise<unknown>[] = [];

    features
      .filter((feature: IOnyxFeatureItem) => feature.FeatureName.trim().length > 0)
      .forEach((feature: IOnyxFeatureItem, index: number) => {
        featureCreates.push(this._sp.web.lists.getByTitle(featuresList).items.add({
          Title: feature.FeatureName,
          BriefId: briefId,
          FeatureName: feature.FeatureName,
          Priority: feature.Priority,
          Phase: feature.Phase,
          Notes: feature.Notes,
          SortOrder: index + 1
        }));
      });

    roles
      .filter((role: IOnyxRoleItem) => role.RoleName.trim().length > 0)
      .forEach((role: IOnyxRoleItem, index: number) => {
        roleCreates.push(this._sp.web.lists.getByTitle(rolesList).items.add({
          Title: role.RoleName,
          BriefId: briefId,
          RoleName: role.RoleName,
          Permissions: role.Permissions,
          SortOrder: index + 1
        }));
      });

    await Promise.all(featureCreates.concat(roleCreates));

    return { briefId };
  }

  public async getMyBriefs(submitterId?: number): Promise<IMyBriefSummary[]> {
    if (!submitterId) {
      return [];
    }

    const items = await this._sp.web.lists
      .getByTitle(productBriefsList)
      .items
      .select('Id', 'Title', 'ProductName', 'Department', 'BriefPriority', 'Status', 'TargetDate', 'DecisionComment', 'DecisionDate', 'Created', 'Modified')
      .filter('SubmitterId eq ' + submitterId)
      .orderBy('Modified', false)
      .top(25)();

    return items as IMyBriefSummary[];
  }

  public async getAllBriefs(): Promise<IMyBriefSummary[]> {
    const items = await this._sp.web.lists
      .getByTitle(productBriefsList)
      .items
      .select('Id', 'Title', 'ProductName', 'Department', 'BriefPriority', 'Status', 'TargetDate', 'DecisionComment', 'DecisionDate', 'Created', 'Modified')
      .orderBy('Modified', false)
      .top(200)();

    return items as IMyBriefSummary[];
  }

  public async getBriefDetails(briefId: number): Promise<{
    brief: IOnyxBriefItem;
    features: IOnyxFeatureItem[];
    roles: IOnyxRoleItem[];
  }> {
    const brief = await this._sp.web.lists
      .getByTitle(productBriefsList)
      .items
      .getById(briefId)();

    const features = await this._sp.web.lists
      .getByTitle(featuresList)
      .items
      .filter('BriefId eq ' + briefId)
      .orderBy('SortOrder', true)();

    const roles = await this._sp.web.lists
      .getByTitle(rolesList)
      .items
      .filter('BriefId eq ' + briefId)
      .orderBy('SortOrder', true)();

    return {
      brief: brief as unknown as IOnyxBriefItem,
      features: features as unknown as IOnyxFeatureItem[],
      roles: roles as unknown as IOnyxRoleItem[]
    };
  }

  public async updateBriefStatus(
    briefId: number,
    status: string,
    decisionComment: string,
    decisionById?: number
  ): Promise<void> {
    const updatePayload: Record<string, unknown> = {
      DecisionComment: decisionComment,
      DecisionDate: new Date().toISOString(),
      DecisionEmailSent: false,
      Status: status
    };

    if (decisionById) {
      updatePayload.DecisionById = decisionById;
    }

    await this._sp.web.lists
      .getByTitle(productBriefsList)
      .items
      .getById(briefId)
      .update(updatePayload);
  }

  public async getStaffProfile(emailAddress: string, staffDirectoryListName: string): Promise<IStaffDirectoryProfile | undefined> {
    if (!emailAddress || !staffDirectoryListName) {
      return undefined;
    }

    const staffList = this._staffSp.web.lists.getByTitle(staffDirectoryListName);
    const fields = await staffList.fields.select('Title', 'InternalName')();
    const emailFieldNames = this._resolveFieldNames(fields, ['EmailAddress', 'Email Address', 'Email', 'Work Email']);
    const departmentFieldNames = this._resolveFieldNames(fields, ['Department']);
    const fullNameFieldNames = this._resolveFieldNames(fields, ['Full Name', 'FullName', 'Title']);
    const lineManagerFieldNames = this._resolveFieldNames(fields, ['LineManager', 'Line Manager', 'Manager']);

    const items = await staffList.items.top(5000)();

    const normalizedEmail = emailAddress.trim().toLowerCase();
    const match = items.filter((item: Record<string, unknown>) => {
      const itemEmail = this._getTextField(item, emailFieldNames.concat(['EmailAddress', 'Email', 'Email_x0020_Address', 'WorkEmail']));
      return itemEmail.trim().toLowerCase() === normalizedEmail;
    })[0] as IStaffDirectoryProfile | undefined;

    if (!match) {
      return undefined;
    }

    return {
      Department: this._getTextField(match as unknown as Record<string, unknown>, departmentFieldNames.concat(['Department'])),
      EmailAddress: this._getTextField(match as unknown as Record<string, unknown>, emailFieldNames.concat(['EmailAddress', 'Email', 'Email_x0020_Address', 'WorkEmail'])),
      FullName: this._getTextField(match as unknown as Record<string, unknown>, fullNameFieldNames.concat(['FullName', 'Full_x0020_Name', 'Title'])),
      LineManager: this._getTextField(match as unknown as Record<string, unknown>, lineManagerFieldNames.concat(['LineManager', 'Line_x0020_Manager', 'LineManager0', 'Line_x0020_Manager0', 'Manager']))
        || this._findFieldByName(match as unknown as Record<string, unknown>, ['linemanager', 'line manager', 'manager'])
    };
  }

  public async ensureUserId(emailAddress: string): Promise<number | undefined> {
    if (!emailAddress) {
      return undefined;
    }

    const ensuredUser = await this._sp.web.ensureUser(emailAddress);
    return ensuredUser.data.Id;
  }

  private _resolveFieldNames(fields: ISharePointFieldInfo[], displayNames: string[]): string[] {
    const normalizedDisplayNames = displayNames.map((displayName: string) => this._normalizeFieldName(displayName));
    const fieldNames: string[] = [];

    fields.forEach((field: ISharePointFieldInfo) => {
      const normalizedTitle = this._normalizeFieldName(field.Title);
      const normalizedInternalName = this._normalizeFieldName(field.InternalName);
      const isMatch = normalizedDisplayNames.some((displayName: string) => {
        return normalizedTitle === displayName || normalizedInternalName === displayName;
      });

      if (isMatch) {
        fieldNames.push(field.InternalName);
      }
    });

    return fieldNames;
  }

  private _normalizeFieldName(fieldName: string): string {
    return fieldName.replace(/_x0020_/g, ' ').replace(/\s/g, '').toLowerCase();
  }

  private _getTextField(item: Record<string, unknown>, fieldNames: string[]): string {
    for (let index = 0; index < fieldNames.length; index++) {
      const value = item[fieldNames[index]];

      if (typeof value === 'string') {
        return value;
      }

      if (value && typeof value === 'object') {
        const typedValue = value as { EMail?: string; Email?: string; Title?: string };
        return typedValue.EMail || typedValue.Email || typedValue.Title || '';
      }
    }

    return '';
  }

  private _findFieldByName(item: Record<string, unknown>, nameParts: string[]): string {
    const keys = Object.keys(item);

    for (let index = 0; index < keys.length; index++) {
      const normalizedKey = keys[index].replace(/_x0020_/g, ' ').toLowerCase();
      const isMatch = nameParts.some((namePart: string) => normalizedKey.indexOf(namePart) > -1);

      if (isMatch) {
        const value = this._getTextField(item, [keys[index]]);

        if (value) {
          return value;
        }
      }
    }

    return '';
  }
}
