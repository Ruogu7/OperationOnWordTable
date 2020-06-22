%% open the doc
word = actxserver('Word.Application');
% Open/Select Word file
[filename_in, pathname] = uigetfile({'*.doc;*.docx;*.docm','Word Files (*.doc,*.docx,*.docm)'; ...
    '*.*', 'All Files (*.*)'}, 'Select a file')
document = word.documents.Open(fullfile(pathname,filename_in));
r=0;
tbl_cnt = document.Tables.Count;
% Loop through each table in the document
for tbl = 1 : tbl_cnt
    row_cnt = document.Tables.Item(tbl).Rows.Count;
    col_cnt = document.Tables.Item(tbl).Columns.Count;
    
    for row = 1 : row_cnt
        r = r+1;
        for col = 1 : col_cnt
            % Pull the values from the table
            cell_txt = strtrim(document.Tables.Item(tbl).Cell(row, col).Range.Text);
            % Add each value to its own cell
            tblVals{r,col} = cell_txt;
        end
    end
end
% Close the document.
document.Close;
% Close Word.
word.Quit;
% delete server object
delete(word);
clear word document

