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