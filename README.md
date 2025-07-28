# Inventor API Scripts

This repository is a collection of iLogic and Visual Basic macros to accelerate common modeling tasks in Autodesk Inventor.

Each script is self-contained and focuses on automating a specific repetitive action, such as applying constraints, exporting files, or cleaning up metadata.

## Contents

| Script Name | Description |
|------------|-------------|
| `ilogic/OriginConstraintToAssembly.vb` | Applies an origin constraint to every component in an open assembly. |


> ✅ Scripts are intended for use through the iLogic editor (`Manage > iLogic > Rules`) or the Inventor VBA environment.

## Usage

1. Open Autodesk Inventor.
2. Go to **Manage > iLogic > Rules** or press `Alt + F11` to access the iLogic editor.
3. Create a new rule and paste the code from any `.vb` file in this repository.
4. Run the rule to perform the automation task.

## Folder Structure

inventor-api-scripts/
├── ilogic/ # All VB iLogic macros/scripts
│ ├── OriginConstraintToAssembly.vb
│ └── ... more coming soon
├── README.md
├── LICENSE

## Contributing

If you’d like to suggest a script or share your own automation, feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License.
