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