from __future__ import annotations
from typing import TypeVar, NamedTuple, Iterator, Tuple, List
from pyadomd._type_code import adomd_type_map, convert
import clr

"""
Copyright 2020 SCOUT

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
"""

# Types
T = TypeVar('T')

class Description(NamedTuple):
    """
    :param [name]: Column name
    :param [type_code]: The column data type
    """
    name: str
    type_code: str

try:
    clr.AddReference('Microsoft.AnalysisServices.AdomdClient')
    from Microsoft.AnalysisServices.AdomdClient import AdomdConnection, AdomdCommand # type: ignore
except Exception as e:  # Catch both Python and .NET exceptions
    print('========================================================================================')
    print(e.ToString())
    print()
    print('This error is raised when Pyadomd is not able to find the AdomdClient.dll file')
    print('The error might be solved by adding the dll to your path. ')
    print('Make sure that the dll is added, to the path, before you import Pyadomd.')
    print()
    print('This can also be resolved by copying the custom.pth to .venv/Lib/site-packages/')
    print('Ensure the custom.pth points to the location that Microsoft.AnalysisServices.AdomdClient.dll is located.')
    print('This is often located in C:\Program Files\Microsoft.NET\ADOMD.NET\170 or similar path')
    print()
    print('If in doubt how to do that, please have a look at Getting Stated in the docs.')
    print('========================================================================================')
    raise

class Cursor:
    
    def __init__(self, connection:AdomdConnection):
        self._conn = connection
        self._description: List[Description] = []
        # Initialize the _row_count to 0
        self._row_count: int = 0

    def close(self) -> None:
        """
        Closes the cursor
        """
        if self.is_closed:
            return
        self._reader.Close()

    def execute(self, query:str) -> Cursor:
        """
        Executes a query against the data source

        :params [query]: The query to be executed
        """
        self._cmd = AdomdCommand(query, self._conn)
        self._reader = self._cmd.ExecuteReader()

        # Set the row count back to 0 so it is not retained from a prior query
        # This could be misleading upon a failure if an old count is retained
        self._row_count = 0

        # Update the columns based on the first result set
        self._update_columns(self._reader)

        # Moved to a function so it can be updated for each result set of a query
        # self._field_count = self._reader.FieldCount
        # self._description = []        
        # for i in range(self._field_count):
        #     self._description.append(Description(
        #             self._reader.GetName(i), 
        #             adomd_type_map[self._reader.GetFieldType(i).ToString()].type_name
        #             ))

        return self

    def executenonquery(self, command:str) -> Cursor:
        """
            Executes a Analysis Services Command agains the database
            
            :params [command]: The command to be executed
        """
        self._cmd = AdomdCommand(command, self._conn)
        # ExecuteNonQuery does not return a reader, but returns an int
        # The int result is not well defined in the Microsoft documentation so don't save it
        # https://learn.microsoft.com/en-us/dotnet/api/microsoft.analysisservices.adomdclient.adomdcommand.executenonquery
        self._cmd.ExecuteNonQuery()
    
        return self

    def _fetchone(self) -> Iterator[Tuple[T, ...]]:
        """
        Fetches the current line from the last executed query
        """
        # Setting the row count to 0 to provide a consistent row count as it would otherwise be the number of rows fetched from the last query executed
        # The row count doesn't really apply to fetchone, but could also be set to 1 instead of 0
        # fetchall and fetchmany will appropriately overwrite the row count after calling this method
        self._row_count = 0

        while(self._reader.Read()):
            # Modified to use the current reader being iterated to support multiple result sets
            yield tuple(convert(self._reader.GetFieldType(i).ToString(), self._reader[i], adomd_type_map) for i in range(self._reader.FieldCount))

    def fetchone(self) -> Tuple[T, ...]:
        """
        Fetches the current line from the last executed query
        """

        result = next(self._fetchone())

        return result
     

    def fetchmany(self, size=1) -> List[Tuple[T, ...]]:
        """
        Fetches one or more lines from the last executed query

        :params [size]: The number of rows to fetch. 
                        If the size parameter exceeds the number of rows returned from the last executed query then fetchmany will return all rows from that query.
        """
        rows: List[Tuple[T, ...]] = []
        try:
            for i in range(size):
                rows.append(next(self._fetchone()))
        except StopIteration:
            pass

        # Update row count for the current result set. This should be equal to or less than the size passed in.
        self._row_count = len(rows)

        # Update the columns as this could be from a subsequent result set
        self._update_columns(self._reader)

        return rows

    def fetchall(self) -> List[Tuple[T, ...]]:
        """
        Fetches all the rows from the last executed query
        """
        # mypy issues with list comprehension :-( 
        _last_results = [i for i in self._fetchone()]

        # Update row count and columns for the current result set. 
        # This is needed for each result set when NextResult is used as the rows and columns can change.
        self._row_count = len(_last_results)
        self._update_columns(self._reader)

        return _last_results # type: ignore
    
    def nextresult(self) -> bool:
        """
        DAX queries can return multiple result sets. These can be iterated through using this method.
        This method moves to the next result set in the current query
        Returns True if there is another result set, False if not
        """
        try:
            return self._reader.NextResult()
        except AttributeError:
            return False

    def _update_columns(self, reader) -> None:
        """
        Updates the column description based on the current reader state
        """
        self._field_count = reader.FieldCount
        self._description = []

        for i in range(reader.FieldCount):
            self._description.append(Description(
            reader.GetName(i), 
            adomd_type_map[reader.GetFieldType(i).ToString()].type_name
            ))

    @property
    def rowcount(self) -> int:
        """
        Returns the number of rows fetched
        """
        return self._row_count
            
    @property
    def has_rows(self) -> bool:
        return self._reader.HasRows
    
    @property
    def is_closed(self) -> bool:
        try:
            state = self._reader.IsClosed
        except AttributeError:
            return True        
        return state

    @property
    def description(self) -> List[Description]:
        return self._description

    def __enter__(self) -> Cursor:
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()

class Pyadomd:

    def __init__(self, conn_str:str):
        self.conn = AdomdConnection()
        self.conn.ConnectionString = conn_str

    def close(self) -> None:
        """
        Closes the connection
        """
        self.conn.Close()
    
    def open(self) -> None:
        """
        Opens the connection
        """
        self.conn.Open()

    def cursor(self) -> Cursor:
        """
        Creates a cursor object
        """
        c = Cursor(self.conn)
        return c
    
    @property
    def state(self) -> int:
        """
        1 = Open
        0 = Closed
        """
        return self.conn.State
    
    def __enter__(self) -> Pyadomd:
        self.open()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.close()