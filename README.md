# Desktop Flow Dependency Fixer for XrmToolBox

![XrmToolBox](https://img.shields.io/badge/XrmToolBox-Plugin-blue)
![NuGet](https://img.shields.io/nuget/v/DesktopFlowDependencyFixer)

**A specialized utility to resolve "Missing Dependency" errors for Desktop Flows (Power Automate Desktop) during solution deployments.**

## The Problem

When deploying solutions containing Desktop Flows (RPA) from Development to Production, you may encounter a **"Missing Dependency"** error regarding a `desktopflowbinary` component.

This often happens in two scenarios:

1.  **The "Base Layer" Conflict:** You are trying to delete a managed solution in Prod, but a `desktopflowbinary` component is missing from the base layer in Dev, blocking the uninstall.
2.  **Accidental Deletion:** A UI element screenshot or binary component was deleted in Dev but still exists in Prod, creating a layering conflict.

This tool fixes these issues by programmatically injecting the missing component into your target unmanaged solution in Development, allowing you to export a valid update.

## Features

- **Targeted Injection:** Adds specific components (by GUID) to an unmanaged solution using the Dataverse `AddSolutionComponent` API.
- **Component Type Selector:** Supports only `desktopflowbinary` (UI Screenshots/Images).
- **Dynamic Metadata:** Automatically detects the correct Object Type Codes for your environment to prevent integer ID errors.
- **Safety Checks:** Validates GUIDs and Solution existence before execution.

## 📦 Installation

1.  Open **XrmToolBox**.
2.  Open the **Tool Library**.
3.  Search for **"Desktop Flow Dependency Fixer"**.
4.  Click **Install**.

## How to Use

### Scenario: Fixing a Missing Base Layer

**Goal:** You need to add a missing binary component to your unmanaged solution in Dev so you can deploy an update.

1.  **Get the Component ID:**
    - In your target environment (e.g., Prod) or via error logs, find the GUID of the missing component (e.g., `d5362aa0-9661-4c73-afda-e54a791834eb`).
2.  **Open the Tool:**
    - Connect to your **Development** environment.
3.  **Configure:**
    - **Component ID:** Paste the GUID.
    - **Component Type:** Select "Desktop Flow Binary" (or Process if the Flow itself is missing).
    - **Target Solution:** Select the unmanaged solution you plan to export.
4.  **Execute:**
    - Click **Move Component to Solution**.
5.  **Finish:**
    - Open your solution in the Power Apps Maker Portal. Verify the component is now present.
    - Export as Managed and deploy to Production to resolve the dependency error.

## ⚠️ Requirements

- **XrmToolBox:** Latest version recommended.
- **Permissions:** System Customizer or System Administrator role in the Dataverse environment.

## 🤝 Contributing

Contributions are welcome! Please submit a Pull Request or open an Issue if you encounter edge cases with specific environment configurations.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

_Author: Oluwafemi Tosin Ajigbayi_
