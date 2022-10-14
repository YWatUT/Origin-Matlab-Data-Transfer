classdef OriginObjClass < handle
      % OriginObjClass help doc:
      % This is for wrapping the Matlab to Origin data communication
      % Origin = OriginObjClass;                             % Open a new Origin instance or just work on the currently open instance
      % 
      % A sample script looks like the following:
      %
      % Origin = OriginObjClass;                             % Initializing handle to origin session
      % % Origin = OriginObjClass(ProjectLocation);  % You can also load from a specific project path, indicated by string ProjectLocation
      %                                                                          % e.g., ProjectLocation = 'C:\Users\SomeonesEID\SomeFolder\SomeProject.opj';
      % 
      % Origin.Send(WorksheetName, data, 1);      % Send data to worksheet 'WorksheetName', starting from (row 1, ignored by default) column 1
      % Origin.SetCol(1, 'Name', 'Xlabel');                               % Set the name of the current column, which is the one just sent into Origin for data, to 'Xlabel'
      % 
      % Origin.Send(WorksheetName, data, 5, 2);                   % Send data to worksheet 'WorksheetName', starting from row 5, column 2;
      % Origin.SetCol(2, 'Name', 'Ylabel', 'Unit', 'Y Unit');       % You can also set multiple parameters in one command. 
      %                                                                                      % Supported are 'Name', 'Unit', and I think 'Comment', all are case sensitive.
      % 
      % Origin.Release;                                             % Release handle so the Origin session can be properly closed
      % 
      % Programming manuals can be found on Origin's knowledge base website, although it's pretty crappy and you have to dig pretty hard 
      % Language used in this class includes both Origin's Class and Methods (e.g.,   invoke(SomeKindOfHandle, 'Method', InputParameters))
      % and their own LabTalk script (e.g.  invoke(SomeKindOfHandle, 'Execute', 'script stuff here');
      
      properties
            h                            % h for ApplicationHandle
                                          % obj.h=actxserver('Origin.ApplicationSI');
            CurrentSheet         % Current sheet handle
      end
            
      methods        
          
        function obj = OriginObjClass(varargin)
              obj.h=actxserver('Origin.ApplicationSI');
              invoke(obj.h, 'Execute', 'doc -mc 3;'); % Make the Origin session visible

              % invoke(OriginObj, 'Execute', 'window -z;');
              % Clear "dirty" flag in Origin to suppress prompt for saving current project
              % You may want to replace this with code to handling of saving current project
              invoke(obj.h, 'IsModified', 'false');

              if ~isempty(varargin) % load a specific Origin project file
                    invoke(obj.h, 'Load', varargin{1});
              end

              ActivePage = invoke(obj.h, 'ActivePage');
              if  ~isempty(ActivePage) % if an origin instance is already open with an active sheet
                    obj.CurrentSheet  = invoke(obj.h, 'FindWorksheet', invoke(ActivePage,'Name'));
              else
                    obj.CurrentSheet = []; % initialize to be blank
              end
        end   
        
        Send(obj, WorksheetName, data, varargin)
        
        CreateWorksheet(obj, WorksheetName)
        
        SetCol(obj, Col, varargin)  
        SetColUserParam(obj, Col, up_name, up_value);

        Release(obj)
        close(obj)

      end
end