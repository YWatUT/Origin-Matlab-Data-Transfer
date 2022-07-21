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

      % Make the Origin session visible
      invoke(obj.h, 'Execute', 'doc -mc 3;');

      % invoke(OriginObj, 'Execute', 'window -z;');

      % Clear "dirty" flag in Origin to suppress prompt for saving current project
      % You may want to replace this with code to handling of saving current project
      invoke(obj.h, 'IsModified', 'false');

      if ~isempty(varargin)
            invoke(obj.h, 'Load', varargin{1});
      end
      
      
      ActivePage = invoke(obj.h, 'ActivePage');
      if  ~isempty(ActivePage) % if an origin instance is already open with an active sheet
            obj.CurrentSheet  = invoke(obj.h, 'FindWorksheet', invoke(ActivePage,'Name'));
      else
            obj.CurrentSheet = []; % initialize to be blank
      end
end

function Send(obj, WorksheetName, data, varargin)
      % Send data into origin
%       Worksheet = invoke(obj.h,'FindWorksheet',WorksheetName);
%       if isempty(Worksheet)
            CreateWorksheet(obj, WorksheetName);
%       end
      
      if ~isempty(varargin)
            largin = length(varargin);
            if largin == 1
                  Row = 0;
                  Col = varargin{1}-1;
            elseif largin ==2
                  Row = varargin{1} - 1;
                  Col = varargin{2} -1;
            else
                  Row = 0;
                  Col = 0;
            end
      else
            Row = 0;
            Col = 0;
      end
      
      invoke(obj.h, 'PutWorksheet', WorksheetName, data, Row, Col);
end
            
function CreateWorksheet(obj, WorksheetName)
      % find a worksheet by name and set it to be the active sheet
      % https://www.originlab.com/doc/en/COM/Classes/Application/CreatePage
      
      WorksheetExist = invoke(obj.h, 'FindWorksheet',WorksheetName);
      if isempty(WorksheetExist) % if the specified worksheet does not exist

            if WorksheetName(1)=='['  % It's '[book]sheet' style
                  temp = split(WorksheetName,']');
                  SheetName = char(temp(2)); % Get sheet
                  temp = split(temp(1),'[');
                  BookName = char(temp(2));  % Get book

                  WorkbookExist = invoke(obj.h,'FindWorksheet',BookName);
                  if isempty(WorkbookExist)  % The book doesn't exist yet
                        BookName = invoke(obj.h, 'CreatePage', 2, BookName, 'Origin');    % This creates a new Book and also auto-focuses to Sheet1
                  else  % The book exist, but sheet doesn't
                        CurrentBook = invoke(obj.h, 'FindWorksheet', BookName);
                        invoke(CurrentBook,'Activate');                                                                  % Before this point the activate sheet is not the one desired
                        invoke(CurrentBook,'Execute','newsheet');                                                % creates a new sheet
                  end
            % Converges to the part where you rename the current sheet
            newSheet = invoke(obj.h, 'FindWorksheet', BookName);
            invoke(newSheet,'Name',SheetName);
            newSheet = invoke(obj.h, 'FindWorksheet', WorksheetName);
            invoke(newSheet,'Activate');
            
            obj.CurrentSheet = newSheet;

            else  % It's a simple 'book' format
                  BookName = invoke(obj.h, 'CreatePage', 2, WorksheetName, 'Origin');
                  newSheet = invoke(obj.h, 'FindWorksheet', WorksheetName);
                  invoke(newSheet,'Activate');

                  obj.CurrentSheet = newSheet;
            end
      else
            % if the specified worksheet already exist, set it to be the active sheet
            newSheet = invoke(obj.h, 'FindWorksheet', WorksheetName);
            invoke(newSheet,'Activate');
            obj.CurrentSheet = newSheet;
      end 
end

function SetCol(obj, Col, varargin)
      
      % Parse varargin into ['Property', 'String'] pairs
      ParsePrep = varargin;
      PropFlag = 1;
      
      Properties = {};
      Values = {};
      
      while ~isempty(ParsePrep)
            ReadIn = ParsePrep{1};
            ParsePrep = ParsePrep(2:end);
            if PropFlag && isColProperty(ReadIn)               % Should be reading in a property string
                  Properties = [Properties, ColProperty(ReadIn)];
                  PropFlag = 0;
            elseif ~PropFlag              % Should be reading in a string for the previous property string
                  Values = [Values, ReadIn];
                  PropFlag = 1;
            else 
                  disp('Bad input, cannot set column info');
                  return 
            end
      end
      
      if length(Properties) ~= length(Values)
            disp('Bad input, cannot set column info');
            return
      end
      
      % Select the current column to input column name, units and etc)
      if ~isempty(obj.CurrentSheet)
            for i = 1:length(Properties)
                  Command = ['Col(',num2str(Col),')[',Properties{i},']$ = ',Values{i}];
                  invoke(obj.CurrentSheet, 'Execute', Command);
            end
      else 
            return
      end
end

function Release(obj)
    if ~isempty(obj.h)
      release(obj.h);
    end
    if ~isempty(obj.CurrentSheet)
      release(obj.CurrentSheet);
    end
end
      end
end


%% Utility functions

function Bool = isColProperty(string)
      PropertyList = {'Name','Unit','Comment'};
      Bool = any(strcmp(PropertyList,string));
end

function Property = ColProperty(string)
      InputList = {'Name','Unit','Comment'};
      % https://www.originlab.com/doc/en/LabTalk/ref/Column-Label-Row-Characters
      OutputSymbol = {'L','U','C'};
      Property = OutputSymbol{strcmp(InputList,string)};
end