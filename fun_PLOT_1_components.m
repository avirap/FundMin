function ...
fun_PLOT_1_components(DATES,...
                      PLOT_main,...
                      PLOT_bench,...
                      TITLE,LEGEND,YLABEL)

%% PREAMBLE
    DATES              = [DATES(1)-1;DATES];
    T=length(DATES);

        % EoM
    DATES_temp = fun_last_day_of_month(DATES);
           EoM = find( DATES_temp(:,2)==1 );
           EoY = find( DATES_temp(:,3)==1 );
                   
    PLOT_main_Return   = [ 0 ;100*(PLOT_main(:,2)-1)];
    PLOT_main_Annlzd   =  100*(PLOT_main(end,3)-1);
    
        B=size(PLOT_bench,3);
for b=1:B
    PLOT_bench_Return(:,b) = [ 0 ;100*(PLOT_bench(:,2,b))];
end

hold on
%% BENCH
color{1}=[0,0,1];
color{2}=[1,0,0];

b=1;

            text(DATES(end)+15,PLOT_bench_Return(end,b),['\fontsize{12}','$',num2str(PLOT_bench_Return(end,b)/100,2),'\fontsize{10.5} for each','\fontsize{12} $1'],'FontWeight','bold','HorizontalAlignment','left','VerticalAlignment','middle','color',color{b});
Pbench(1,b)=plot(DATES,PLOT_bench_Return(:,b),    'LineWidth',1.75,'color',color{b});
Pbench(2,b)=plot(DATES,PLOT_bench_Return(:,b),'o','LineWidth',1.20,'MarkerFaceColor',color{b},'color','k'); Pbench(2,b).MarkerIndices=EoM;

b=2;

            text(DATES(end)+15,PLOT_bench_Return(end,b),['\fontsize{12}','$',num2str(PLOT_bench_Return(end,b)/100,1),'\fontsize{10.5} on each','\fontsize{12} $1'],'FontWeight','bold','HorizontalAlignment','left','VerticalAlignment','middle','color',color{b});
Pbench(1,b)=plot(DATES,PLOT_bench_Return(:,b),    'LineWidth',1.75,'color',color{b});
Pbench(2,b)=plot(DATES,PLOT_bench_Return(:,b),'o','LineWidth',1.20,'MarkerFaceColor',color{b},'color','k'); Pbench(2,b).MarkerIndices=EoM;

clear color
%% MAIN
color=[103/255,142/255,47/255];

%    text(DATES(end),PLOT_main_Return(end),{['\fontsize{12}',num2str(PLOT_main_Return(end),'%.1f'),'%'];['\fontsize{10}','(',num2str(PLOT_main_Annlzd,'%.1f'),'%*)']},'FontWeight','bold','HorizontalAlignment','left','VerticalAlignment','middle','color',color);
     text(DATES(end)+15,PLOT_main_Return(end),{['\fontsize{12} $1','\fontsize{10.5} turned to \fontsize{12}$',num2str(PLOT_main_Return(end)/100+1,3)];['\fontsize{12} *',num2str(PLOT_main_Annlzd,'%.1f'),'%','\fontsize{10.5} annualized']},'FontWeight','bold','HorizontalAlignment','left','VerticalAlignment','middle','color',color);

Pmain(1,1)=plot(DATES,PLOT_main_Return,    'LineWidth',2.50,'color',color);
Pmain(2,1)=plot(DATES,PLOT_main_Return,'o','LineWidth',1.50,'MarkerFaceColor',color,'color','k');                        Pmain(2).MarkerIndices=EoM;

plot(DATES,zeros(T,1),'k');

%% Y limits
Ymin=min(min([PLOT_main_Return,PLOT_bench_Return]));
Ymax=max(max([PLOT_main_Return,PLOT_bench_Return]));
set(gca,'Ylim',[Ymin,Ymax]);

%% VERTICALS
% fun_event_dates(DATES);
fun_monthly_lines(DATES,EoM);

%% Xticks
Y=length(EoY);
M=length(EoM);

if Y<=2
   for m=2:M
   yr=year(DATES(EoM(m)));
   mnth=month(DATES(EoM(m)));
   dy=day(DATES(EoM(m)));
%    dy=floor(dy/2);
   dtnm=datenum(yr,mnth,dy);
   TICKS(m-1)=find(DATES==dtnm);
   end
   set(gca,'XTick',DATES(TICKS));
   datetick('x', '', 'keepticks','keeplimits');
end
if Y>=3
    for y=2:Y
        yr=year(DATES(EoY(y)));
           mnth=month(DATES(EoY(y)));
           mnth=floor(mnth/2);
        if mnth==0; mnth=1; end
        dy=eomday(yr,mnth);
        dtnm=datenum(yr,mnth,dy);
        TICKS(y-1)=find(DATES==dtnm);
    end
    set(gca,'XTick',DATES(TICKS));
    datetick('x','yyyy','keepticks','keeplimits');
end

%% Ytick
Ydiff = Ymax - Ymin;
if 100 < Ydiff && Ydiff < 9999; x=20; set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',12); end
if 50  < Ydiff && Ydiff < 100;  x=10; set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',12); end
if 20  < Ydiff && Ydiff < 50;   x=5;  set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',12); end
if 10  < Ydiff && Ydiff < 20;   x=2;  set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',10); end
if  0  < Ydiff && Ydiff < 10;   x=1;  set(gca,'YTick',[x*(fix(Ymin/x):1:fix(Ymax/x))],'fontsize',10); end
ytickformat('%.f') 
%% HORIZONTALS
grid on
ax = gca; ax.GridLineStyle = '-'; ax.GridAlpha = 0.2;

axis tight

%% NO resizing of Font Size
set(gcf, 'PaperPositionMode', 'auto')
%%
title(TITLE,'color',[103/255,142/255,47/255],'FontSize',14,'FontWeight','bold','FontName','Fredoka One');
if B<=2; legend([Pmain(1),Pbench(1,:)],LEGEND,'location','northwest','FontSize',14,'FontWeight','bold','FontName','Roboto'); end
if B>=3; legend([Pmain(1),Pbench(1,:)],LEGEND,'location','northwest','FontSize',09,'FontWeight','bold','FontName','Roboto'); end
ylabel(YLABEL    ,'FontSize',14,'fontweight','bold');

hold off

% xlswrite('\reports\FIG_1.xlsx',[str2num(datestr(DATES,'YYYYmmDD')),PLOT_main_Return,PLOT_bench_Return,[zeros(length(DATES)-1,1);PLOT_main_Annlzd]])

end