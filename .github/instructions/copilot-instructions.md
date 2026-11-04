This project targets SharePoint Server Subscription Edition as the baseline but also supports SharePoint Online.

Mandatory rules:

Do not assume SharePoint Online.
Do not use Microsoft Graph unless already present.
Verify all APIs against SharePoint SE support.
Verify SPFx version compatibility.
Do not introduce unsupported Fluent UI packages.
Do not change package-solution.json IDs.
Preserve all existing component IDs.
Validate skipsFeatureDeployment behavior.
Validate feature.xml generation.
Validate deployment to on-prem App Catalog.