function [] = STO(WorksheetName, data, varargin)
% 2022-02 new version of SendToOrigin for sending data from Matlat to Origin in an automated fashion

% This time trying to leverage different sheets in one worksheet(page) so that I don't have to
% withstand the stupid short name limitations of Origin anymore. 

if ~isempty(varargin)
    largin = length(varargin);
    if largin == 1
        StartingColumn = varargin{1} - 1;
        StartingRow = 0;
        ProjectLocation=[];
    elseif largin==2
        if iscellstr(varargin(2))
            StartingRow = varargin{1} - 1 ;
            StartingColumn=0;
            ProjectLocation = char(varargin(2));
        else
            StartingRow = varargin{1} - 1 ;
            StartingColumn = varargin{2} - 1;
            ProjectLocation=[];
        end
    elseif largin==3
        StartingRow = varargin{1} - 1 ;
        StartingColumn = varargin{2} - 1;
        ProjectLocation = varargin{3};
    end
else
    StartingRow=0;
    StartingColumn=0;
    ProjectLocation=[];
end

% Associate with either a new or existing Origin instance
OriginObj = StartOrigin(ProjectLocation);

CreateWorksheet(OriginObj, WorksheetName);

PutWorksheet(OriginObj,WorksheetName, data, StartingRow, StartingColumn);

release(OriginObj);



end

function [OriginObj] = StartOrigin(ProjectLocation)

OriginObj=actxserver('Origin.ApplicationSI');

% Make the Origin session visible
invoke(OriginObj, 'Execute', 'doc -mc 3;');

% invoke(OriginObj, 'Execute', 'window -z;');
       
% Clear "dirty" flag in Origin to suppress prompt for saving current project
% You may want to replace this with code to handling of saving current project
invoke(OriginObj, 'IsModified', 'false');

if ~isempty(ProjectLocation)
     invoke(OriginObj, 'Load', ProjectLocation);
end


end

function [] = CreateWorksheet(OriginObj, WorksheetName)
% find a worksheet by name

%{
%     Empty string -- this means the active sheet from the active book
%     Book name only -- like "Book1", will get the active sheet from named book
%     Sheet name only with ! at the end -- like "Sheet2!", will get the named sheet from the active book
%}
WorksheetExist = invoke(OriginObj, 'FindWorksheet',WorksheetName);
% If the named worksheet is found then an Origin Worksheet is returned. 
% so isempty can be used to determine stuff

if isempty(WorksheetExist) % if the specified worksheet does not exist
    
    if WorksheetName(1)=='['  % It's '[book]sheet' style
        temp = split(WorksheetName,']');
        SheetName = char(temp(2)); % Get sheet
        temp = split(temp(1),'[');
        BookName = char(temp(2));  % Get book
        
        WorkbookExist = invoke(OriginObj,'FindWorksheet',BookName);
        if isempty(WorkbookExist)  % The book doesn't exist yet
            BookName = invoke(OriginObj, 'CreatePage', 2, BookName, 'Origin');
            % This creates a new Book and also auto-focuses to Sheet1
        else  % The book exist, but sheet doesn't
            CurrentBook = invoke(OriginObj, 'FindWorksheet', BookName);
            invoke(CurrentBook,'Activate');
            % At this point the activate sheet is not the one desired
            invoke(CurrentBook,'Execute','newsheet'); % creates a new sheet
        end
        % Converges to the part where you rename the current sheet
        newSheet = invoke(OriginObj, 'FindWorksheet', BookName);
        invoke(newSheet,'Name',SheetName);
        
    else  % It's a simple 'book' format
        BookName = invoke(OriginObj, 'CreatePage', 2, WorksheetName, 'Origin');
    end
    
    %{
    % Create a workbook
    % The CreatePage method creates an Origin page of the specified type, such 
    % as a worksheet page or graph page.
    % ReturnName = invoke(originObj, 'CreatePage', type, name, template);
    % type (integer) 1: Matrix, 2: Worksheet, 3: Graph, 4: Layout, 5: Notes
    % name (string)
    % template (string) I don't know, the example uses 'Origin'   
    % Returns a string containing the name of the page created. 
    %}
    
end % if the specified worksheet already exist, just ignore

% at the end of this, a worksheet with the name Worksheet_name will exist
% so no output is needed

end

function [] = PutWorksheet(OriginObj, WorksheetName, data, StartingRow, StartingColumn)
% Put data into a worksheet
% return = PutWorksheet(name, data, first row, first column);
% default first row/column is 1
% if not specified, it's gonna overwrite every time
invoke(OriginObj, 'PutWorksheet', WorksheetName, data, StartingRow, StartingColumn);

end