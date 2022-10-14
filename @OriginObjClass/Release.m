function Release(obj)
% .Release and .close do the same thing
    if ~isempty(obj.CurrentSheet)
      release(obj.CurrentSheet);
    end
    
    if ~isempty(obj.h)
      release(obj.h);
    end
end