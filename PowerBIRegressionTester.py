import os
import sys
from sys import path
import json
import glob
import hashlib
import pandas as pd
import importlib.util
from enum import Enum
from tabulate import tabulate

class PowerBIRegressionTester:
    """
    A class to perform regression testing on Power BI DAX queries by comparing query results
    between a baseline and new test runs. Handles event loading, query execution, result hashing,
    and comparison, as well as mapping visuals to their page names.
    """

    PROJECT_FOLDER_BASE = "Projects"
    QUERIES_BASE_FOLDER = "Query Files"
    BASELINE_FOLDER_NAME = "baseline"
    INSTANCE_FOLDER_NAME = "instance"
    BASELINE_CSV_FILE = "baseline.csv"
    BASELINE_PARQUET_FILE = "baseline.parquet"
    DAX_STUDIO_QUERY_FOLDER_NAME = "DAX Studio"
    POWER_BI_PERF_ANALYZER_QUERY_FOLDER_NAME = "PBI Performance Analyzer"

    class QueryType(Enum):
        DAX_STUDIO = "DAXSTUDIO"
        PERFORMANCE_ANALYZER = "PERFORMANCEANALYZER"

    def __init__(self, project_folder, connection_string, pbi_report_folder):
        """
        Initialize the regression tester with project and environment details.

        Args:
            project_name (str): Name of the Power BI project.
            working_directory (str): Path to the working directory.
            connection_string (str): Connection string for the Power BI semantic model.
            pbi_report_folder (str): Path to the folder containing the Power BI report definition.
        """
        self.project_name = project_folder
        self.working_directory = os.getcwd()
        # self.server = server
        # self.catalog = model
        # self.connection_string = f"Provider=MSOLAP;Data Source={server};Initial Catalog={model}"
        self.connection_string = connection_string
        self.pbi_report_folder = pbi_report_folder
        self.project_folder = os.path.join(self.working_directory, self.PROJECT_FOLDER_BASE, project_folder)
        # self.pbi_pa_folder = os.path.join(self.project_folder, self.QUERIES_BASE_FOLDER)
        
        self.baseline_folder = os.path.join(self.project_folder, self.BASELINE_FOLDER_NAME)
        self.baseline_csv_file = os.path.join(self.baseline_folder, self.BASELINE_CSV_FILE)
        self.baseline_parquet_file = os.path.join(self.baseline_folder, self.BASELINE_PARQUET_FILE)
        self.instance_folder_base = os.path.join(self.project_folder, self.INSTANCE_FOLDER_NAME)
        # self.query_type = query_type

        # self.query_subfolder = None
        # if self.query_type == self.QueryType.DAX_STUDIO:
        #     self.query_subfolder = "DAX Studio"
        # elif self.query_type == self.QueryType.PERFORMANCE_ANALYZER:
        #     self.query_subfolder = "PBI Performance Analyzer"

        # self.query_folder = os.path.join(self.project_folder, self.QUERIES_BASE_FOLDER, self.query_subfolder)
        self.dax_studio_query_folder = os.path.join(self.project_folder, self.QUERIES_BASE_FOLDER, self.DAX_STUDIO_QUERY_FOLDER_NAME)
        self.power_bi_perf_analyzer_query_folder = os.path.join(self.project_folder, self.QUERIES_BASE_FOLDER, self.POWER_BI_PERF_ANALYZER_QUERY_FOLDER_NAME)


        adomd_path = r'C:\Program Files\Microsoft.NET\ADOMD.NET\160'

        if not os.path.isdir(adomd_path):
            print("Folder does not exist.")
            sys.exit(1)

        # Check if the ADOMD.NET path is already in the system path
        if adomd_path not in path:
            path.append(adomd_path)

    @classmethod
    def for_compare_only(cls, project_folder):
        """
        Alternate constructor for compare-only usage (no connection string or ADOMD.NET dependency).

        Args:
            project_folder (str): Name of the Power BI project.
            working_directory (str): Path to the working directory.
            pbi_report_folder (str): Path to the folder containing the Power BI report definition.

        Returns:
            PowerBIRegressionTester: An instance with only file paths set.
        """
        obj = cls.__new__(cls)
        obj.project_folder = project_folder
        obj.working_directory = os.getcwd()
        obj.project_folder = os.path.join(obj.working_directory, obj.PROJECT_FOLDER_BASE, project_folder)
        obj.pbi_report_folder = ""
        obj.baseline_folder = os.path.join(obj.project_folder, cls.BASELINE_FOLDER_NAME)
        obj.baseline_csv_file = os.path.join(obj.baseline_folder, cls.BASELINE_CSV_FILE)
        obj.baseline_parquet_file = os.path.join(obj.baseline_folder, cls.BASELINE_PARQUET_FILE)
        obj.instance_folder_base = os.path.join(obj.project_folder, cls.INSTANCE_FOLDER_NAME)
        obj.connection_string = None  # Not needed for compare-only

        # obj.query_type = query_type

        # obj.query_subfolder = None
        # if obj.query_type == obj.QueryType.DAX_STUDIO:
        #     obj.query_subfolder = "DAX Studio"
        # elif obj.query_type == obj.QueryType.PERFORMANCE_ANALYZER:
        #     obj.query_subfolder = "PBI Performance Analyzer"

        # obj.query_folder = os.path.join(obj.project_folder, obj.QUERIES_BASE_FOLDER, obj.query_subfolder)
        obj.dax_studio_query_folder = os.path.join(obj.project_folder, obj.QUERIES_BASE_FOLDER, obj.DAX_STUDIO_QUERY_FOLDER_NAME)
        obj.power_bi_perf_analyzer_query_folder = os.path.join(obj.project_folder, obj.QUERIES_BASE_FOLDER, obj.POWER_BI_PERF_ANALYZER_QUERY_FOLDER_NAME)

        return obj

    def baseline_exists(self):
        """
        Check if the baseline parquet file exists.

        Returns:
            bool: True if the baseline parquet file exists, False otherwise.
        """
        return os.path.isfile(self.baseline_parquet_file)

    def instance_exists(self, instance_name):
        """
        Check if the parquet file for a given instance exists.

        Args:
            instance_name (str): The name of the instance to check.

        Returns:
            bool: True if the instance parquet file exists, False otherwise.
        """
        instance_folder = os.path.join(self.instance_folder_base, instance_name)
        instance_parquet_file = os.path.join(instance_folder, f"{instance_name}.parquet")
        return os.path.isfile(instance_parquet_file)
        
    def find_visual_id(self, id_map, start_id):
        """
        Traverse the event id_map to find the visualId for a given event id.

        Args:
            id_map (dict): Mapping of event ids to event data.
            start_id (str): The starting event id.

        Returns:
            str or None: The visualId if found, else None.
        """
        current_id = start_id
        while current_id:
            node = id_map.get(current_id)
            if not node:
                break
            metrics = node.get('metrics', {})
            visual_id = metrics.get('visualId')
            if visual_id:
                return visual_id
            current_id = node.get('parentId')
        return None

    def flatten_event(self, event):
        """
        Flatten the 'metrics' dictionary into the top-level event dictionary.

        Args:
            event (dict): The event dictionary.

        Returns:
            dict: The flattened event dictionary.
        """
        flat = event.copy()
        metrics = flat.pop("metrics", {})
        flat.update(metrics)
        return flat

    def row_hash(self, row):
        """
        Generate a SHA-256 hash for a DataFrame row by concatenating its values.

        Args:
            row (pd.Series): The row to hash.

        Returns:
            str: The SHA-256 hash string.
        """
        row_str = '|'.join(str(val) for val in row.values)
        return hashlib.sha256(row_str.encode('utf-8')).hexdigest()

    def find_visual_json_by_id(self, report_folder, visual_id):
        """
        Recursively search for a visual.json file whose root 'name' matches the given visual_id.

        Args:
            report_folder (str): The root folder to search.
            visual_id (str): The visualId to match.

        Returns:
            str or None: The path to the matching visual.json file, or None if not found.
        """
        for root, dirs, files in os.walk(report_folder):
            for file in files:
                if file == "visual.json":
                    file_path = os.path.join(root, file)
                    try:
                        with open(file_path, "r", encoding="utf-8") as f:
                            data = json.load(f)
                            if data.get("name") == visual_id:
                                return file_path
                    except Exception as e:
                        print(f"Error reading {file_path}: {e}")
        return None

    # def get_page_display_name(self, visual_json_path):
    #     """
    #     Go up three directories from a visual.json file to find the corresponding page.json,
    #     and extract the displayName.

    #     Args:
    #         visual_json_path (str): Path to the visual.json file.

    #     Returns:
    #         str or None: The displayName from page.json, or None if not found.
    #     """
    #     page_dir = os.path.dirname(os.path.dirname(os.path.dirname(visual_json_path)))
    #     page_json_path = os.path.join(page_dir, "page.json")
    #     if os.path.isfile(page_json_path):
    #         with open(page_json_path, "r", encoding="utf-8") as f:
    #             page_data = json.load(f)
    #             return page_data.get("displayName")
    #     return None

    def build_visualid_to_pagename_map(self, report_folder):
        visualid_to_pagename = {}
        for root, dirs, files in os.walk(report_folder):
            if "visual.json" in files:
                visual_json_path = os.path.join(root, "visual.json")
                try:
                    with open(visual_json_path, "r", encoding="utf-8") as f:
                        visual_data = json.load(f)
                    visual_id = visual_data.get("name")
                    # Go up three directories to find page.json
                    page_dir = os.path.dirname(os.path.dirname(os.path.dirname(visual_json_path)))
                    page_json_path = os.path.join(page_dir, "page.json")
                    page_name = None
                    if os.path.isfile(page_json_path):
                        with open(page_json_path, "r", encoding="utf-8") as pf:
                            page_data = json.load(pf)
                            page_name = page_data.get("displayName")
                    if visual_id:
                        visualid_to_pagename[visual_id] = page_name
                except Exception as e:
                    print(f"Error reading {visual_json_path}: {e}")
        return visualid_to_pagename

    # def add_page_names_to_df(self, df, report_folder):
    #     """
    #     Add a 'pageName' column to the DataFrame by mapping each visualId to its page display name.

    #     Args:
    #         df (pd.DataFrame): DataFrame with a 'visualId' column.
    #         report_folder (str): Folder to search for visual.json and page.json files.

    #     Returns:
    #         pd.DataFrame: DataFrame with an added 'pageName' column.
    #     """
    #     page_names = []
    #     for visual_id in df['visualId']:
    #         visual_json_path = self.find_visual_json_by_id(report_folder, visual_id)
    #         if visual_json_path:
    #             page_name = self.get_page_display_name(visual_json_path)
    #         else:
    #             page_name = None
    #         page_names.append(page_name)
    #     df['pageName'] = page_names
    #     return df

    # def load_events(self):
    #     """
    #     Load and flatten all events from JSON files in the PBI PA Files folder.

    #     Returns:
    #         pd.DataFrame: Combined DataFrame of all events.
    #     """
    #     all_dfs = []

    #     if self.pbi_report_folder:
    #         visualid_to_pagename = self.build_visualid_to_pagename_map(self.pbi_report_folder)

    #     for json_path in glob.glob(os.path.join(self.power_bi_perf_analyzer_query_folder, r"*.json")):
    #         with open(json_path, "r", encoding="utf-8-sig") as f:
    #             data = json.load(f)
    #         events = data.get("events", [])
    #         flat_events = [self.flatten_event(e) for e in events]
    #         df = pd.DataFrame(flat_events)
    #         if not df.empty:
    #             id_map = {event['id']: event for event in events if 'id' in event}
    #             df['VisualID'] = df['id'].apply(lambda x: self.find_visual_id(id_map, x))
    #             if self.pbi_report_folder:
    #                 df['PageName'] = df['VisualID'].map(visualid_to_pagename)
    #             else:
    #                 df['PageName'] = ""

    #             df['ResultSets'] = 0

    #             all_dfs.append(df)
    #     if all_dfs:
    #         concat_df = pd.concat(all_dfs, ignore_index=True)
    #         filtered_df = concat_df[concat_df['name'] == 'Execute DAX Query'].copy()
    #         filtered_df.drop_duplicates(inplace=True)

    #         # filtered_df = filtered_df.rename(columns={
    #         #     "id": "ID", 
    #         #     "QueryText": "Query"
    #         # })
    #         # filtered_df = filtered_df.drop([
    #         #     'name', 'component', 'start', 'end', 'sourceLabel', 'status',
    #         #     'visualTitle', 'visualType', 'visualId', 'initialLoad', 'parentId'
    #         # ], axis=1)
    #         # desired_order = ['ID', 'Query', 'PageName', 'VisualID', 'ResultSets', 'RowCount']
    #         # filtered_df = filtered_df[desired_order]

    #         filtered_df = filtered_df.rename(columns={"id": "ID", "QueryText": "Query"})
    #         desired_order = ['ID', 'Query', 'PageName', 'VisualID', 'ResultSets', 'RowCount']
    #         filtered_df = filtered_df[desired_order]

    #         return filtered_df
    #     return pd.DataFrame()

    def load_events(self, query_type):
        """
        Load and flatten all events from JSON files in the PBI PA Files folder.

        Returns:
            pd.DataFrame: Combined DataFrame of all events.
        """
        # query_folder = None
        # if query_type == self.QueryType.PERFORMANCE_ANALYZER:
        #     query_folder = self.power_bi_perf_analyzer_query_folder
        # elif query_type == self.QueryType.DAX_STUDIO:
        #     query_folder = self.dax_studio_query_folder

        all_dfs = []
        filtered_df = pd.DataFrame()

        if query_type == self.QueryType.DAX_STUDIO:
            for json_path in glob.glob(os.path.join(self.dax_studio_query_folder, r"*.json")):
                with open(json_path, "r", encoding="utf-8-sig") as f:
                    data = json.load(f)
                events = data
                # id_map = {event['id']: event for event in events if 'id' in event}
                df = pd.DataFrame(events)
                if not df.empty:
                    # df['visualId'] = df['id'].apply(lambda x: self.find_visual_id(id_map, x))
                    # df = self.add_page_names_to_df(df, self.pbi_report_folder)
                    df['VisualID'] = ""
                    df['PageName'] = ""
                    df['RowCount'] = "0"
                    df['ResultSets'] = 0

                    all_dfs.append(df)

            if all_dfs:
                concat_df = pd.concat(all_dfs, ignore_index=True)
                filtered_df = concat_df[concat_df['QueryType'] == 'DAX'].copy()
                filtered_df = filtered_df.rename(columns={"RequestID": "ID"})

        if query_type == self.QueryType.PERFORMANCE_ANALYZER:
            if self.pbi_report_folder:
                visualid_to_pagename = self.build_visualid_to_pagename_map(self.pbi_report_folder)

            for json_path in glob.glob(os.path.join(self.power_bi_perf_analyzer_query_folder, r"*.json")):
                with open(json_path, "r", encoding="utf-8-sig") as f:
                    data = json.load(f)

                events = data.get("events", [])
                flat_events = [self.flatten_event(e) for e in events]
                df = pd.DataFrame(flat_events)
                if not df.empty:
                    id_map = {event['id']: event for event in events if 'id' in event}
                    df['VisualID'] = df['id'].apply(lambda x: self.find_visual_id(id_map, x))

                    if self.pbi_report_folder:
                        df['PageName'] = df['VisualID'].map(visualid_to_pagename)
                    else:
                        df['PageName'] = ""

                    df = df.drop(['RowCount'], axis=1) # Drop the numeric field from the PBI PA file
                    df['RowCount'] = ""
                    df['ResultSets'] = 0

                    all_dfs.append(df)

            if all_dfs:
                concat_df = pd.concat(all_dfs, ignore_index=True)
                filtered_df = concat_df[concat_df['name'] == 'Execute DAX Query'].copy()
                filtered_df = filtered_df.rename(columns={"id": "ID", "QueryText": "Query"})

        if filtered_df is not None and not filtered_df.empty:
            filtered_df.drop_duplicates(inplace=True)

            # Rename columns and select only the relevant ones for clarity
            # filtered_df = filtered_df.rename(columns={"RequestID": "ID"})
            # filtered_df = filtered_df.drop(columns=[
            #     'Duration', 'StartTime', 'EndTime', 'Username', 'DatabaseName', 'QueryType',
            #     'AggregationMatchCount', 'AggregationMissCount', 'RequestProperties',
            #     'RequestParameters', 'ActivityID'
            # ])
            # filtered_df = filtered_df[['ID', 'Query', 'PageName', 'VisualID', 'ResultSets', 'RowCount']]

            
            desired_columns = ['ID', 'Query', 'PageName', 'VisualID', 'ResultSets', 'RowCount']
            filtered_df = filtered_df[desired_columns]

            return filtered_df
        
        return pd.DataFrame()

    def execute_queries(self, filtered_df):
        """
        Execute DAX queries from the DataFrame, hash the results, and store hashes in the DataFrame.

        Args:
            filtered_df (pd.DataFrame): DataFrame with DAX queries to execute.

        Returns:
            pd.DataFrame: DataFrame with result set and query hashes.
        """
        from pyadomd import Pyadomd
        filtered_df['ResultSets'] = 0
        filtered_df['CombinedQueryHash'] = ""
        with Pyadomd(self.connection_string) as conn:
            for idx, row in filtered_df.iterrows():
                query_id = row['ID']
                query_text = row['Query']
                try:
                    with conn.cursor().execute(query_text) as cur:
                        result_set_index = 0
                        row_count = ""
                        resultset_hashes = []
                        has_next = True
                        while has_next:
                            result_set_index += 1
                            columns = [col[0] for col in cur.description]
                            result_rows = cur.fetchall()
                            if result_rows:
                                result_df = pd.DataFrame(result_rows, columns=columns)
                                if not result_df.empty:
                                    if row_count:
                                        row_count += f", {len(result_df)}"
                                    else:
                                        row_count = str(len(result_df))
                                    result_df['row_hash'] = result_df.apply(self.row_hash, axis=1)
                                    result_df = result_df.sort_values('row_hash').reset_index(drop=True)
                                    row_hashes = '|'.join(result_df['row_hash'].tolist())
                                    combined_row_hash = hashlib.sha256(row_hashes.encode('utf-8')).hexdigest()
                                    resultset_hashes.append(combined_row_hash)
                            has_next = cur.nextresult()
                        filtered_df.loc[idx, 'ResultSets'] = result_set_index
                        filtered_df.loc[idx, 'RowCount'] = row_count
                        if resultset_hashes:
                            resultset_hashes.sort()
                            combined = '|'.join(resultset_hashes)
                            combined_query_hash = hashlib.sha256(combined.encode('utf-8')).hexdigest()
                            filtered_df.loc[idx, 'CombinedQueryHash'] = combined_query_hash
                        else:
                            filtered_df.loc[idx, 'CombinedQueryHash'] = None
                except Exception as e:
                    print(f"Error executing query {query_id}: {e}\n")
        return filtered_df

    def compare_with_baseline(self, instance):
        """
        Compare the current run's DataFrame with the baseline, returning only differing rows.

        Args:
            filtered_df (pd.DataFrame): DataFrame from the current run.
            instance_name (str): Name of the current instance.

        Returns:
            pd.DataFrame: DataFrame of value differences, with page names added.
        """

        if isinstance(instance, pd.DataFrame):
            instance_df = instance
        if isinstance(instance, str):
            if not self.instance_exists(instance):
                print(f"Instance '{instance}' does not exist. Please run the instance first.")
                sys.exit(1)

            instance_one_folder = os.path.join(self.instance_folder_base, instance)
            instance_one_parquet_file = os.path.join(instance_one_folder, f"{instance}.parquet")
            instance_df = pd.read_parquet(instance_one_parquet_file)

        baseline_exists = os.path.isfile(self.baseline_parquet_file)
        if not baseline_exists:
            print("Baseline file does not exist. Please create a baseline first.")
            sys.exit(1)
        baseline_df = pd.read_parquet(self.baseline_parquet_file)

        value_diffs = self.compare_internal(baseline_df, instance_df)

        # comparison_df = filtered_df.merge(
        #     baseline_df, on='ID', suffixes=('', '_baseline'), how='outer', indicator=True
        # )
        # cols_to_compare = ['CombinedQueryHash']
        # diff_mask = (comparison_df['_merge'] == 'both') & (
        #     comparison_df[[col for col in cols_to_compare]].ne(
        #         comparison_df[[f"{col}_baseline" for col in cols_to_compare]].values
        #     ).any(axis=1)
        # )
        # value_diffs = comparison_df[diff_mask]

        # desired_columns = ['ID', 'Query', 'PageName', 'VisualID', 'ResultSets', 'RowCount']
        # value_diffs = value_diffs[desired_columns]
        
        return value_diffs
    
    def compare_instances(self, instance_one, instance_two):
        if not self.instance_exists(instance_one):
            print(f"Instance '{instance_one}' does not exist. Please run the instance first.")
            sys.exit(1)

        instance_one_folder = os.path.join(self.instance_folder_base, instance_one)
        instance_one_parquet_file = os.path.join(instance_one_folder, f"{instance_one}.parquet")
        instance_one_df = pd.read_parquet(instance_one_parquet_file)

        if not self.instance_exists(instance_one):
            print(f"Instance '{instance_one}' does not exist. Please run the instance first.")
            sys.exit(1)

        instance_two_folder = os.path.join(self.instance_folder_base, instance_two)
        instance_two_parquet_file = os.path.join(instance_two_folder, f"{instance_two}.parquet")
        instance_two_df = pd.read_parquet(instance_two_parquet_file)

        value_diffs = self.compare_internal(instance_one_df, instance_two_df)
        
        return value_diffs
    
    def compare_internal(self, instance_one_df, instance_two_df):
        comparison_df = instance_one_df.merge(
            instance_two_df, on='ID', suffixes=('', '_baseline'), how='outer', indicator=True
        )
        cols_to_compare = ['CombinedQueryHash']
        diff_mask = (comparison_df['_merge'] == 'both') & (
            comparison_df[[col for col in cols_to_compare]].ne(
                comparison_df[[f"{col}_baseline" for col in cols_to_compare]].values
            ).any(axis=1)
        )
        value_diffs = comparison_df[diff_mask]

        desired_columns = ['ID', 'Query', 'PageName', 'VisualID', 'ResultSets', 'RowCount']
        value_diffs = value_diffs[desired_columns]
        
        return value_diffs

    def prepare_df(self):
        """
        Shared logic to load, filter, and process events, execute queries, and compute hashes.

        Returns:
            pd.DataFrame: The processed DataFrame ready for baseline or comparison.
        """

        # Columns from a PA PBI file
        # ['name', 'component', 'start', 'id', 'sourceLabel', 'end', 'status', 'visualTitle', 'visualId', 'visualType', 'initialLoad', 'parentId', 'QueryText', 'RowCount']

        # if self.query_type == self.QueryType.DAX_STUDIO:
        #     final_df = self.load_events_DAX_Studio()
        # elif self.query_type == self.QueryType.PERFORMANCE_ANALYZER:
        #     final_df = self.load_events()
        # else:
        #     raise ValueError("Unsupported query type. Use DAX_STUDIO or PERFORMANCE_ANALYZER.")

        # Check if the Power BI Performance Analyzer query folder has any .json files
        has_perf_analyzer_files = (
            os.path.isdir(self.power_bi_perf_analyzer_query_folder) and
            any(fname.endswith('.json') for fname in os.listdir(self.power_bi_perf_analyzer_query_folder))
        )

        # Check if the DAX Studio query folder has any .json files
        has_dax_studio_files = (
            os.path.isdir(self.dax_studio_query_folder) and
            any(fname.endswith('.json') for fname in os.listdir(self.dax_studio_query_folder))
        )

        power_bi_perf_analyzer = pd.DataFrame()
        dax_studio_df = pd.DataFrame()

        if has_perf_analyzer_files:
            power_bi_perf_analyzer = self.load_events(self.QueryType.PERFORMANCE_ANALYZER)

        if has_dax_studio_files:
            dax_studio_df = self.load_events(self.QueryType.DAX_STUDIO)

        combined_df = pd.concat([dax_studio_df, power_bi_perf_analyzer], ignore_index=True)

        if not combined_df.empty and 'Query' in combined_df.columns:
            combined_df['Query'] = combined_df['Query'].apply(self.normalize_line_endings)
        # final_df['QueryHash'] = final_df['Query'].apply(lambda x: hashlib.sha256(str(x).encode('utf-8')).hexdigest())

        combined_df = self.execute_queries(combined_df)
        all_hashes = sorted(combined_df['CombinedQueryHash'].dropna().astype(str).tolist())
        combined = '|'.join(all_hashes)
        final_overall_hash = hashlib.sha256(combined.encode('utf-8')).hexdigest()
        combined_df['FinalOverallHash'] = final_overall_hash

        return combined_df


    # def prepare_df_DAX_Studio(self):
    #     """
    #     Shared logic to load, filter, and process events, execute queries, and compute hashes.

    #     Returns:
    #         pd.DataFrame: The processed DataFrame ready for baseline or comparison.
    #     """

    #     # Columns from a PA PBI file
    #     # ['name', 'component', 'start', 'id', 'sourceLabel', 'end', 'status', 'visualTitle', 'visualId', 'visualType', 'initialLoad', 'parentId', 'QueryText', 'RowCount']

    #     final_df = self.load_events_DAX_Studio()
    #     pd.set_option('display.max_columns', None)
    #     pd.set_option('display.max_rows', 20)
    #     filtered_df = final_df
    #     # filtered_df = final_df[final_df['name'] == 'Execute DAX Query'].copy()
    #     # filtered_df.drop_duplicates(inplace=True)
    #     filtered_df = self.execute_queries(filtered_df)
    #     all_hashes = filtered_df['CombinedQueryHash'].dropna().astype(str).tolist()
    #     combined = '|'.join(all_hashes)
    #     final_overall_hash = hashlib.sha256(combined.encode('utf-8')).hexdigest()
    #     filtered_df['FinalOverallHash'] = final_overall_hash
    #     # filtered_df = self.add_page_names_to_df(filtered_df, self.pbi_report_folder)
    #     # filtered_df = filtered_df[
    #     #     ['id', 'parentId', 'visualId', 'QueryText', 'RowCount', 'ResultSets',
    #     #      'CombinedQueryHash', 'final_overall_hash', 'pageName']
    #     # ]
    #     return filtered_df
    
    def run_baseline(self):
        """
        Create and save a new baseline DataFrame.

        Returns:
            pd.DataFrame: The baseline DataFrame.
        """
        # if self.query_type == self.QueryType.DAX_STUDIO:
        #     filtered_df = self.prepare_df_DAX_Studio()
        # elif self.query_type == self.QueryType.PERFORMANCE_ANALYZER:
        #     filtered_df = self.prepare_df()
        # else:
        #     raise ValueError("Unsupported query type. Use DAX_STUDIO or PERFORMANCE_ANALYZER.")
        
        filtered_df = self.prepare_df()
        os.makedirs(self.baseline_folder, exist_ok=True)
        filtered_df.to_csv(self.baseline_csv_file, index=False)
        filtered_df.to_parquet(self.baseline_parquet_file, index=False)
        return filtered_df

    def run_instance(self, instance_name):
        """
        Run an instance, save its results, and compare to the baseline.

        Args:
            instance_name (str): Name of the test instance.

        Returns:
            pd.DataFrame: DataFrame of value differences.
        """
        instance_df = self.prepare_df()
        instance_folder = os.path.join(self.instance_folder_base, instance_name)
        os.makedirs(instance_folder, exist_ok=True)
        instance_csv_file = os.path.join(instance_folder, f"{instance_name}.csv")
        instance_parquet_file = os.path.join(instance_folder, f"{instance_name}.parquet")

        # # convert to \r\n for Windows compatibility in  CSV and Parquet files
        # # Step 1: Normalize everything to \n
        # instance_df['Query'] = instance_df['Query'].str.replace('\r\n', '\n').str.replace('\r', '\n')
        # # Step 2: Convert to Windows-style CRLF
        # instance_df['Query'] = instance_df['Query'].str.replace('\n', '\r\n')

        instance_df.to_csv(instance_csv_file, index=False)
        instance_df.to_parquet(instance_parquet_file, index=False)
        value_diffs = self.compare_with_baseline(instance_df)
        return value_diffs
    
    def compare(self, instance_name):
        """
        Load the baseline and specified instance from parquet files and compare them.

        Args:
            instance_name (str): The name of the instance to compare against the baseline.

        Returns:
            pd.DataFrame: DataFrame of value differences, with page names added.
        """
        # # Load baseline and instance DataFrames
        # if not self.baseline_exists():
        #     print("Baseline file does not exist. Please create a baseline first.")
        #     sys.exit(1)
        # if not self.instance_exists(instance_name):
        #     print(f"Instance '{instance_name}' does not exist. Please run the instance first.")
        #     sys.exit(1)

        # instance_folder = os.path.join(self.instance_folder_base, instance_name)
        # instance_parquet_file = os.path.join(instance_folder, f"{instance_name}.parquet")
        # instance_df = pd.read_parquet(instance_parquet_file)

        # value_diffs = self.compare_with_baseline(instance_df)

        value_diffs = self.compare_with_baseline(instance_name)
        return value_diffs
    
        # # Merge and compare
        # comparison_df = instance_df.merge(
        #     baseline_df, on='id', suffixes=('', '_baseline'), how='outer', indicator=True
        # )
        # cols_to_compare = ['combined_query_hash']
        # diff_mask = (comparison_df['_merge'] == 'both') & (
        #     comparison_df[[col for col in cols_to_compare]].ne(
        #         comparison_df[[f"{col}_baseline" for col in cols_to_compare]].values
        #     ).any(axis=1)
        # )
        # value_diffs = comparison_df[diff_mask]
        # value_diffs = value_diffs[
        #     ['id', 'parentId', 'visualId', 'QueryText', 'RowCount', 'result_sets',
        #      'combined_query_hash', 'combined_query_hash_baseline',
        #      'final_overall_hash', 'final_overall_hash_baseline']
        # ]
        # value_diffs = self.add_page_names_to_df(value_diffs, self.pbi_report_folder)
        # return value_diffs

    def load_baseline_df(self):
        """Load the baseline DataFrame from the parquet file."""
        if os.path.isfile(self.baseline_parquet_file):
            return pd.read_parquet(self.baseline_parquet_file)
        return pd.DataFrame()

    def load_instance_df(self, instance_name):
        """Load the instance DataFrame from the parquet file."""
        instance_folder = os.path.join(self.instance_folder_base, instance_name)
        instance_parquet_file = os.path.join(instance_folder, f"{instance_name}.parquet")
        if os.path.isfile(instance_parquet_file):
            return pd.read_parquet(instance_parquet_file)
        return pd.DataFrame()

    def normalize_line_endings(self, text: str) -> str:
        return text.replace('\r\n', '\n').replace('\r', '\n')

    def normalize_to_crlf(self, s):
        if pd.isna(s):
            return s
        return s.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')

    @staticmethod
    def create_project_skeleton(project_name):
        """
        Create a skeleton folder structure for a new Power BI regression test project.

        Structure:
        Projects/
            <project_name>/
                baseline/
                instance/
                Query Files/
                    DAX Studio/
                    PBI Performance Analyzer/
        """
        base_dir = os.path.join(os.getcwd(), "Projects", project_name)
        folders = [
            base_dir,
            os.path.join(base_dir, "baseline"),
            os.path.join(base_dir, "instance"),
            os.path.join(base_dir, "Query Files"),
            os.path.join(base_dir, "Query Files", "DAX Studio"),
            os.path.join(base_dir, "Query Files", "PBI Performance Analyzer"),
        ]
        for folder in folders:
            os.makedirs(folder, exist_ok=True)