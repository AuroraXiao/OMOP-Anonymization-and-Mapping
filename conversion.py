import os
import pandas as pd

BASE_DIRECTORY = "R://Code//Python_Code//Result_Tables"


def process_files(base_directory):

    for file_name in os.listdir(base_directory):
        file_path = os.path.join(base_directory, file_name)
        if not os.path.isfile(file_path):
            continue

        output_directory = os.path.join(base_directory, "OMOP_CDM_tables")
        os.makedirs(output_directory, exist_ok=True)

        if file_name.startswith("mapped_demo_patient"):
            process_person_and_death(file_path, output_directory)
        elif file_name.startswith("mapped_surgical_patient"):
            process_procedure_occurrence(file_path, output_directory)
        elif file_name.startswith("mapped_lab_results"):
            process_measurement(file_path, output_directory)
        elif file_name.startswith("mapped_case_visit"):
            process_visit_occurrence(file_path, output_directory)
            process_condition_occurrence(file_path, output_directory)
        elif file_name.startswith("mapped_medication_order"):
            process_drug_exposure(file_path, output_directory)




def process_person_and_death(file_path, output_directory):
    df_source = pd.read_excel(file_path)

    # person
    df_person = pd.DataFrame()
    df_person["person_id"] = df_source["Patient_No"]
    df_person["gender_concept_id"] = df_source["Gender"]
    df_person["year_of_birth"] = df_source["Patient_DOB"].astype(int)
    df_person["race_concept_id"] = df_source["Race"]
    df_person["ethnicity_concept_id"] = df_source["ethnicity_concept_id"]# 38003564
    df_person.to_excel(os.path.join(output_directory, "person.xlsx"), index=False)
    print(f"Person table processed for {file_path}.")

    # death
    df_death = pd.DataFrame()
    df_death["person_id"] = df_source["Patient_No"]
    df_death["death_date"] = df_source["Death_Date"]
    df_death.to_excel(os.path.join(output_directory, "death.xlsx"), index=False)
    print(f"Death table processed for {file_path}.")


# procedure_occurrence
def process_procedure_occurrence(file_path, output_directory):
    df_source = pd.read_excel(file_path)

    df_procedure_occurrence = pd.DataFrame()
    df_procedure_occurrence["procedure_occurrence_id"] = df_source["Surgical_Code"]
    df_procedure_occurrence["procedure_concept_id"] = df_source["Surgical_Desc"]
    df_procedure_occurrence["procedure_date"] = df_source["operation_date"]
    df_procedure_occurrence["person_id"] = df_source["Patient_No"]
    df_procedure_occurrence["procedure_type_concept_id"] = "32827"
    df_procedure_occurrence["visit_occurrence_id"] = df_source["Case_No"]
    df_procedure_occurrence.to_excel(
        os.path.join(output_directory, "procedure_occurrence.xlsx"), index=False
    )
    print(f"Procedure Occurrence table processed for {file_path}.")


# measurement
def process_measurement(file_path, output_directory):
    df_source = pd.read_excel(file_path)

    df_measurement = pd.DataFrame()
    df_measurement["person_id"] = df_source["Patient_No"]
    df_measurement["measurement_concept_id"] = df_source["Test_Code_ID"]
    df_measurement["measurement_date"] = pd.to_datetime(
        df_source["Result_Test_Date"]
    ).dt.date
    df_measurement["value_as_number"] = df_source["Test_Result"]
    df_measurement["range_low"] = df_source["Minimum_Range"]
    df_measurement["range_high"] = df_source["Maximum_Range"]
    df_measurement["measurement_source_value"] = df_source["Short_Text"]
    df_measurement["measurement_type_concept_id"] = "EHR"
    df_measurement.to_excel(
        os.path.join(output_directory, "measurement.xlsx"), index=False
    )
    print(f"Measurement table processed for {file_path}.")


# visit_occurrence
def process_visit_occurrence(file_path, output_directory):
    df_source = pd.read_excel(file_path)

    df_visit_occurrence = pd.DataFrame()
    df_visit_occurrence["visit_occurrence_id"] = df_source["Case_No"]
    df_visit_occurrence["person_id"] = df_source["Patient_No"]
    # df_visit_occurrence["visit_concept_id"] = "38004307"
    df_visit_occurrence["visit_start_datetime"] = pd.to_datetime(
        df_source["Adm_DateTime"]
    )
    df_visit_occurrence["visit_end_datetime"] = pd.to_datetime(
        df_source["Dis_DateTime"]
    )
    # df_visit_occurrence["visit_type_concept_id"] = "32817"
    df_visit_occurrence["visit_source_value"] = df_source["Adm_Type_Desc"]
    df_visit_occurrence["admitted_from_source_value"] = df_source[
        "Adm_Dept_Description"
    ]
    df_visit_occurrence.to_excel(
        os.path.join(output_directory, "visit_occurrence.xlsx"), index=False
    )
    print(f"Visit Occurrence table processed for {file_path}.")


# condition_occurrence
def process_condition_occurrence(file_path, output_directory):
    df_source = pd.read_excel(file_path)

    df_condition_occurrence = pd.DataFrame()
    df_condition_occurrence["condition_occurrence_id"] = df_source["Case_No"]
    df_condition_occurrence["person_id"] = df_source["Patient_No"]
    df_condition_occurrence["condition_concept_id"] = df_source["condition_concept_id"]
    df_condition_occurrence["condition_start_date"] = pd.to_datetime(
        df_source["Adm_DateTime"]
    ).dt.date
    df_condition_occurrence["condition_end_date"] = pd.to_datetime(
        df_source["Dis_DateTime"]
    ).dt.date
    # df_condition_occurrence["condition_type_concept_id"] = "32817" 
    df_condition_occurrence["condition_status_concept_id"] = df_source[
        "Discharge_Type_Desc"
    ]
    df_condition_occurrence["stop_reason"] = df_source["Dis_Reason"]
    df_condition_occurrence["visit_occurrence_id"] = df_source["Case_No"]
    df_condition_occurrence["condition_source_value"] = df_source[
        "Primary_Diagnosis_Description_Mediclaim"
    ]
    # df_condition_occurrence["condition_status_source_value"] = "32890"
    df_condition_occurrence.to_excel(
        os.path.join(output_directory, "condition_occurrence.xlsx"), index=False
    )
    print(f"Condition Occurrence table processed for {file_path}.")


def process_drug_exposure(file_path, output_directory):
    df_source = pd.read_excel(file_path)

    df_drug_exposure = pd.DataFrame()
    df_drug_exposure["person_id"] = df_source["Patient_No"]
    df_drug_exposure["drug_concept_id"] = df_source["Cluster_Preferred_Name_Code"]
    df_drug_exposure["drug_exposure_start_date"] = pd.to_datetime(
        df_source["Order_Creation_Date"]
    ).dt.date
    df_drug_exposure["drug_exposure_end_date"] = df_source["Drug_Exposure_End_Date"]
    # df_drug_exposure["drug_type_concept_id"] = "32838"
    df_drug_exposure["quantity"] = df_source["Dosage_Ordered"]
    df_drug_exposure["dose_unit_concept_id"] = df_source["Dosage_Ordered_Unit"]
    df_drug_exposure["route_concept_id"] = df_source["Dosage_Form"]
    df_drug_exposure["drug_source_value"] = df_source["Medication_Order_Component_Text"]
    df_drug_exposure["visit_occurrence_id"] = df_source["Case_No"]
    df_drug_exposure.to_excel(
        os.path.join(output_directory, "drug_exposure.xlsx"), index=False
    )
    print(f"Drug Exposure table processed for {file_path}.")


if __name__ == "__main__":
    if os.path.isdir(BASE_DIRECTORY):
        process_files(BASE_DIRECTORY)
        print("All files processed.")
    else:
        print("Invalid directory. Please check the path and try again.")