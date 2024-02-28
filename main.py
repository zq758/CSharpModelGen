import pandas as pd
import argparse
import os
import json


# Function to load configuration from a JSON file
def load_config(config_file):
    with open(config_file, "r", encoding="utf-8") as file:
        return json.load(file)


# Attempt to load configuration and handle potential errors
try:
    config = load_config("config.json")
except FileNotFoundError:
    print("Configuration file not found. Please ensure config.json is present.")
    exit(1)
except json.JSONDecodeError:
    print("Error parsing configuration file. Please check the format of config.json.")
    exit(1)

# Extract mappings from the configuration
excel_columns = config["excel_columns"]
data_type_mapping = config["data_type_mapping"]


def generate_csharp_property(name, type, required, explanations, remarks, indent=""):
    csharp_type = data_type_mapping.get(type, "string") + "?"
    explanation_str = "".join(explanations)  # Combine explanations into a single line
    remarks_str = ("\n" + indent + "/// ").join(
        remarks
    )  # Format remarks for C# multiline comment
    comment_str = (
        f"{explanation_str}\n" + indent + f"/// {remarks_str}"
        if remarks_str
        else explanation_str
    )
    property_str = (
        indent
        + f"/// <summary>\n"
        + indent
        + f"/// {comment_str}\n"
        + indent
        + "/// </summary>\n"
    )
    # if required == "Y":
    #     property_str += indent + "[Required]\n"
    property_str += indent + f"public {csharp_type} {name} {{ get; set; }}\n\n"
    return property_str


# Function to aggregate field data from dataframes
def aggregate_field_data(dfs, field_name):
    types, required_values, explanations, remarks = set(), set(), [], []
    for df in dfs:
        field_data = df[df[excel_columns["name"]] == field_name]
        types.update(field_data[excel_columns["type"]].dropna().unique())
        required_values.update(field_data[excel_columns["required"]].dropna().unique())
        explanations.extend(field_data[excel_columns["description"]].dropna().unique())
        remarks.extend(field_data[excel_columns["remarks"]].dropna().unique())
    type = list(types)[0] if types else "string"
    required = "Y" if "Y" in required_values else "N"
    return type, required, explanations, remarks


# Function to find common fields in a list of dataframes
def find_common_fields(dfs):
    common_fields = set(dfs[0][excel_columns["name"]].dropna())
    for df in dfs[1:]:
        sheet_fields = set(df[excel_columns["name"]].dropna())
        common_fields &= sheet_fields
    return common_fields


# Function to write a C# class string to a file
def write_class_to_file(class_name, class_str, output_dir):
    with open(
        os.path.join(output_dir, f"{class_name}.cs"), "w", encoding="utf-8"
    ) as file:
        file.write(class_str)


# Function to generate a C# class from data
def generate_class(namespace, class_name, fields, dfs, base_class_name=None):
    indent = ""
    class_str = ""

    # Namespace
    if namespace:
        class_str += f"namespace {namespace}\n{{\n"
        indent += "    "

    # Class declaration
    if base_class_name:
        class_str += f"{indent}public class {class_name} : {base_class_name}\n"
    else:
        class_str += f"{indent}public class {class_name}\n"
    class_str += f"{indent}{{\n"

    # Generating properties for each field
    for field_name in fields:
        type, required, explanations, remarks = aggregate_field_data(dfs, field_name)
        class_str += generate_csharp_property(
            field_name, type, required, explanations, remarks, indent + "    "
        )

    # Closing class and namespace
    class_str += f"{indent}}}\n"
    if namespace:
        class_str += "}\n"

    return class_str


# Function to generate C# classes from an Excel file
def generate_csharp_classes_from_excel(excel_file, group_names, namespace, output_dir):
    xl = pd.ExcelFile(excel_file)

    # Check if output directory exists, if not create it
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Process each group
    for group_name in group_names:
        # Identify sheets in the group
        group_sheets = [sheet for sheet in xl.sheet_names if group_name in sheet]
        dfs = [xl.parse(sheet_name=sheet) for sheet in group_sheets]

        # Forward fill to handle merged cells
        for df in dfs:
            df[excel_columns["name"]].ffill(inplace=True)

        # Find common fields across all sheets in the group
        common_fields = find_common_fields(dfs)

        # Generate base class for common fields
        base_class_name = f"Base{group_name}"
        base_class_str = generate_class(
            namespace, base_class_name, common_fields, [dfs[0]]
        )
        write_class_to_file(base_class_name, base_class_str, output_dir)

        # Generate child classes for each sheet
        for sheet_name, df in zip(group_sheets, dfs):
            unique_fields = set(df[excel_columns["name"]]) - common_fields
            child_class_name = sheet_name
            child_class_str = generate_class(
                namespace, child_class_name, unique_fields, [df], base_class_name
            )
            write_class_to_file(child_class_name, child_class_str, output_dir)


# Main function to parse arguments and generate classes
def main():
    parser = argparse.ArgumentParser(description="Generate C# classes from Excel data.")
    parser.add_argument("-file", help="Path to the Excel file.", default="1.xlsx")
    parser.add_argument(
        "-g",
        "--groups",
        nargs="+",
        default=["InData", "OutData"],
        help="List of group names to process.",
    )
    parser.add_argument(
        "-n",
        "--namespace",
        default="Models",
        help="The namespace to use for the generated classes.",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="output",
        help="The output directory to write the classes to.",
    )

    args = parser.parse_args()

    generate_csharp_classes_from_excel(
        args.file, args.groups, args.namespace, args.output
    )


if __name__ == "__main__":
    main()
