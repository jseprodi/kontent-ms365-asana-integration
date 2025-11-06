/**
 * Management API utilities for fetching additional context data
 * that's not available in the Custom App SDK context
 */

export interface WorkflowStep {
  id: string;
  name: string;
  codename: string;
}

export interface Contributor {
  id: string;
  email: string;
  firstName?: string;
  lastName?: string;
}

/**
 * Fetch workflow step information for a content item
 * Note: This requires Management API access token
 */
export async function getWorkflowStep(
  contentItemId: string,
  languageId: string,
  managementApiToken: string,
  projectId: string
): Promise<WorkflowStep | null> {
  try {
    const response = await fetch(
      `https://manage.kontent.ai/v2/projects/${projectId}/items/${contentItemId}/variants/${languageId}`,
      {
        headers: {
          Authorization: `Bearer ${managementApiToken}`,
          'Content-Type': 'application/json',
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch variant: ${response.statusText}`);
    }

    const variant = await response.json();
    return variant.workflow_step ? {
      id: variant.workflow_step.id,
      name: variant.workflow_step.name,
      codename: variant.workflow_step.codename,
    } : null;
  } catch (error) {
    console.error('Failed to fetch workflow step:', error);
    return null;
  }
}

/**
 * Fetch contributors for a content item
 * Note: This requires Management API access token
 */
export async function getContributors(
  contentItemId: string,
  languageId: string,
  managementApiToken: string,
  projectId: string
): Promise<Contributor[]> {
  try {
    const response = await fetch(
      `https://manage.kontent.ai/v2/projects/${projectId}/items/${contentItemId}/variants/${languageId}`,
      {
        headers: {
          Authorization: `Bearer ${managementApiToken}`,
          'Content-Type': 'application/json',
        },
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch variant: ${response.statusText}`);
    }

    const variant = await response.json();
    return variant.contributors?.map((c: any) => ({
      id: c.id,
      email: c.email,
      firstName: c.first_name,
      lastName: c.last_name,
    })) || [];
  } catch (error) {
    console.error('Failed to fetch contributors:', error);
    return [];
  }
}

