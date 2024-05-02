#!/usr/bin/env python3
import dxpy
import re

class DXFile():
    '''
    Process and handle DNAnexus files, as this script can be run both on 
    DNAnexus and locally.
    '''
    def read_dx_file(self, file):
        '''
        read a dx file
        '''
        print(f"Reading from {file}")

        if isinstance(file, dict):
            # provided as {'$dnanexus_link': '[project-xxx:]file-xxx'}
            file = file.get('$dnanexus_link')

        if re.match(r'^file-[\d\w]+$', file):
            # just file-xxx provided => find a project context to use
            file_details = self.get_file_project_context(file)
            project = file_details.get('project')
            file_id = file_details.get('id')
        elif re.match(r'^project-[\d\w]+:file-[\d\w]+', file):
            # nicely provided as project-xxx:file-xxx
            project, file_id = file.split(':')
        else:
            # who knows what's happened, not for me to deal with
            raise RuntimeError(
                f"DXFile not in an expected format: {file}"
            )
        
        return dxpy.DXFile(
            project=project, dxid=file_id).read().rstrip('\n').split('\n')

    def get_file_project_context(self, file) -> dxpy.DXObject:
        '''
        Get project ID for a given file ID, used where only file ID is
        provided as DXFile().read() requires both, will ensure that
        only a live version of a project context is returned.
        Inputs:
            file: a dnanexus file ID

        Outputs:
            DXObject: DXObject file handler object
        '''
        print(f"Searching all projects for: {file}")

        # find projects where file exists and get DXFile objects for
        # each to check archivalState, list_projects() returns dict
        # where key is the project ID and value is permission level
        projects = dxpy.DXFile(dxid=file).list_projects()
        print(f"Found file in {len(projects)} project(s)")

        files = [
            dxpy.DXFile(dxid=file, project=id).describe()
            for id in projects.keys()
        ]

        # filter out any archived files or those resolving
        # to the current job container context
        files = [
            x for x in files
            if x['archivalState'] == 'live'
            and not re.match(r"^container-[\d\w]+$", x['project'])
        ]
        assert files, f"No live files could be found for the ID: {file}"

        print(
            f"Found {file} in {len(files)} projects, "
            f"using {files[0]['project']} as project context"
        )

        return files[0]
