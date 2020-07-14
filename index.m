clear all
close all
clc

[imageName, imagePath] = uigetfile({'*.png'; '*.jpg'; '*.jpeg'},  'Choose an image')
rawImage = imread(strcat(imagePath, imageName));

[a, b, c] = size(rawImage);
red = rawImage(:,:,1);
green = rawImage(:,:,2); 
blue = rawImage(:,:,3);

fileName = strsplit(imageName, '.');
outname = strcat(imagePath, strcat(fileName{1}, '.xlsx'));
xlswrite(outname, ' ');
excel=actxserver('Excel.application');
wb=excel.Workbooks.Open(outname);

for i = 1 : a
    for j = 1 : b
        cellName = strcat(idx2col(j), num2str(i));
        rgb_val = @(r,g,b) double(r)*1+double(g)*256+double(b)*256^2;
        colorValue = rgb_val(red(i, j), green(i, j), blue(i, j));
        wb.Worksheets.Item(1).Range(cellName).Interior.Color  = colorValue;
        disp(strcat(num2str(i), 'x', num2str(j)));
    end
end

        
        
        
wb.Save
wb.Close
winopen(outname)


clear all
close all



