function xlcol_num=col2idx(xlcol_addr)
% xlcol_addr - upper case character
    if ischar(xlcol_addr) && ~any(~isstrprop(xlcol_addr,"upper"))
        xlcol_num=0;
        n=length(xlcol_addr);
        for k=1:n
            xlcol_num=xlcol_num+(double(xlcol_addr(k)-64))*26^(n-k);
        end
    else
        error('not a valid character')
    end
end