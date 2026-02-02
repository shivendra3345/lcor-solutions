Property Manager WebPart component

Overview
- Lightweight React component and PnPJS service to manage the `Property` list in the root site (`https://lcor1.sharepoint.com`).
- Provides basic CRUD and configurable "views" (a view is a set of fields that the form shows).

Files added
- `src/webparts/propertyManager/services/PropertyService.ts` — PnPJS wrapper (init, getItems, getItem, create, update, delete).
- `src/webparts/propertyManager/components/PropertyManager.tsx` — React component using Fluent UI to list and edit items by view.
- `src/webparts/propertyManager/components/PropertyManager.module.scss` — minimal styles.
- `src/webparts/propertyManager/index.ts` — simple export.

Integration
1. Install PnPJS dependencies and rebuild:

```bash
npm install @pnp/sp @pnp/common @pnp/logging @pnp/odata --save
npm run build
```

2. Use the `PropertyManager` component from any existing SPFx React web part. Example in your web part render:

```tsx
import * as React from 'react';
import PropertyManager from '../propertyManager/components/PropertyManager';

export default class Reports extends React.Component<any, any> {
  public render(): React.ReactElement<any> {
    return (
      <PropertyManager context={this.props.context} />
    );
  }
}
```

Notes & next steps
- This is a starting point. Optional improvements:
  - Fetch list field schema and render dynamic input types (Choice, Number, Date, People) instead of simple text fields.
  - Allow admins to define and persist custom views (stored in a configuration list or web part properties).
  - Add paging, filtering, and column formatting for the list.
  - Add role/permission checks before CRUD operations.

Security
- The web part acts using the current user's permissions. Ensure users who will manage list entries have appropriate list permissions.
