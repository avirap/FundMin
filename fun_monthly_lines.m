function fun_monthly_lines(DATES,EoM)

YLIM=ylim;


%% Vertical Lines
% Beg Day
% line([DATES(1) DATES(1)],YLIM,'color','k','LineStyle','-','LineWidth',2.5);

% EoM | EoY
for m=1:length(EoM)
    [Mnum,Mstr] =month(DATES(EoM(m)));
    if   Mnum==12;  LineStyle='-'; LineWidth=1.00;
    else            LineStyle='-'; LineWidth=0.20;
    end        
    line([DATES(EoM(m)) DATES(EoM(m))],YLIM,'color','k','LineStyle',LineStyle,'LineWidth',LineWidth);
end

%% 'mmm'
for m=2:length(EoM)
    [Mnum,Mstr] =month(DATES(EoM(m)));
    dy=floor(day(DATES(EoM(m)))/2);
    mn=month(DATES(EoM(m)));
    yr=year(DATES(EoM(m)));
    dtnm=datenum(yr,mn,dy);
    fnd=find(DATES==dtnm);
    if 37 < length(EoM) && length(EoM) <= 49; text(DATES(fnd),YLIM(2),{Mstr(1)},'color','k','HorizontalAlignment', 'center','VerticalAlignment', 'middle','FontWeight','normal','FontSize',5.5); end
    if 25 < length(EoM) && length(EoM) <= 37; text(DATES(fnd),YLIM(2),{Mstr(1)},'color','k','HorizontalAlignment', 'center','VerticalAlignment', 'middle','FontWeight','normal','FontSize',8);   end
    if 13 < length(EoM) && length(EoM) <= 25; text(DATES(fnd),YLIM(2),{Mstr}   ,'color','k','HorizontalAlignment', 'center','VerticalAlignment', 'middle','FontWeight','normal','FontSize',8);   end
    if 0  < length(EoM) && length(EoM) <= 13; text(DATES(fnd),YLIM(1),{Mstr}   ,'color','k','HorizontalAlignment', 'center','VerticalAlignment', 'top','FontWeight','normal','FontSize',11);  end
%     text(MoM(m),YLIM(1)-1.75,{Mstr},'color','k','FontWeight','normal','FontSize',9);
end

end