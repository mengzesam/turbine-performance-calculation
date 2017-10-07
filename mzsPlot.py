# -*- coding: utf-8 -*-
"""
Copyright (C) 2014-2016, Zoomer Analytics LLC.
All rights reserved.

License: BSD 3-clause (see LICENSE.txt for details)
"""
import numpy as np
import pandas
import matplotlib.pyplot as plt
import xlwings as xw
from matplotlib.figure import Figure
from matplotlib.backends.backend_qt4agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.backends.backend_qt4agg import NavigationToolbar2QT as NavigationToolbar
from PyQt4 import QtCore, QtGui
import timer
import matplotlib
matplotlib.rcParams['font.sans-serif'] = ['SimHei']

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

class Ui_MainWindow(QtGui.QMainWindow):
    def __init__(self,drawStyle=0):#drawStyle==0普通趋势；drawStyle==1，增加振动相位
        super(Ui_MainWindow, self).__init__()
        self.Members_init(drawStyle)
        self.setupUi(self)
        self.Canvas_init()
        self.Tableview_init()
        self.set_fig_handle()

    def set_fig_handle(self):
        if self.fig is None:
            pass
        else:
            self.fig.canvas.mpl_connect('button_press_event',self.fig_on_press)
            self.fig.canvas.mpl_connect('motion_notify_event',self.fig_on_motion)
            self.fig.canvas.mpl_connect('button_release_event',self.fig_on_release)
            self.fig.canvas.mpl_connect('pick_event',self.fig_on_pick)

    def fig_on_press(self,event):
        i=0
        if (event.button==1 and event.dblclick):
            if not (event.xdata is None or event.ydata is None):
                if self.cursor[0]==-1 and self.cursor[1]==-1:
                    self.curCursorIdx=0
                elif self.cursor[1]==-1:
                    self.curCursorIdx=1
                else:
                    if np.abs(event.xdata-self.cursor[0])<=np.abs(event.xdata-self.cursor[1]):
                        self.curCursorIdx = 0
                    else:
                        self.curCursorIdx =1
                self.cursor[self.curCursorIdx]=self.fig_setVline(event)
        elif event.button==1:
            self.press=True

    def fig_on_motion(self,event):
        if not self.pickedline==None:
            self.pickedline.set_alpha(0.5)
            #self.canvas.draw() 不更新
        if self.cursor[0]!=-1 and self.press:
            if self.cursor[1]==-1:
                self.curCursorIdx = 0
            else:
                if np.abs(event.xdata - self.cursor[0]) <= np.abs(event.xdata - self.cursor[1]):
                    self.curCursorIdx=0
                else:
                    self.curCursorIdx=1
            self.cursor[self.curCursorIdx] = self.fig_setVline(event)
	    self.ax.set_title(u'')
        if self.drawStyle==1:
	        self.ax_angle.set_title(u'')

    def fig_on_release(self,event):
        self.press=False

    def fig_on_pick(self,event):
        if not self.pickedline==None:
            self.pickedline.set_alpha(0.5)
        self.pickedline=event.artist
        self.pickedline.set_alpha(1.5-self.pickedline.get_alpha())
        event.artist.axes.set_title(u'Picked Line:%s' % (self.pickedline.get_label()))
        if event.artist.axes==self.ax_angle:
            self.ax.set_title(u"")
        else:
	    if self.drawStyle==1:
            	self.ax_angle.set_title(u"")
        self.canvas.draw()

    def fig_setVline(self,event):
        if self.tagCount<1: return -1
        if (event.xdata is None or event.ydata is None):return -1
        colors=['red','blue']
        if self.myVline[self.curCursorIdx]==None:
            xs=np.array([event.xdata,event.xdata])
            ys=np.array([0,100])
            if self.drawStyle==1:
                ymin,ymax=self.ax_angle.get_ylim()
                self.myVline_angle[self.curCursorIdx],=self.ax_angle.plot(xs,[ymin,ymax],color=colors[self.curCursorIdx])
            self.myVline[self.curCursorIdx],=self.ax.plot(xs,ys,color=colors[self.curCursorIdx])
        else:
            xs=np.array([event.xdata,event.xdata])
            ys=np.array([0,100])
            if self.drawStyle==1:
                ymin,ymax=self.ax_angle.get_ylim()
                self.myVline_angle[self.curCursorIdx].set_data(xs,[ymin,ymax])
            self.myVline[self.curCursorIdx].set_data(xs,ys)
        if self.XisTime:
            ss = u"游标%d:%s" % (self.curCursorIdx+1,("%s"%(matplotlib.dates.num2date(event.xdata,tz=None)))[0:22])
        else:
            ss=u"游标%d:%.3f"%(self.curCursorIdx+1,event.xdata)
        if self.Cursor_TextObjs[self.curCursorIdx] is None:
            mycolor='red'
            XPos=0.01
            HAlig='left'
            if self.curCursorIdx==1:
                mycolor='blue'
                XPos=0.99
                HAlig = 'right'
            self.Cursor_TextObjs[self.curCursorIdx]=self.ax.text(XPos, 0.022, ss,
                       verticalalignment='bottom', horizontalalignment=HAlig,
                       transform=self.ax.transAxes,
                       color=mycolor, fontsize=15,
                       bbox={'facecolor': '#88FF00', 'alpha': 0.2, 'pad': 10})
        else:
            self.Cursor_TextObjs[self.curCursorIdx].set_text(ss)
        self.canvas.draw()
        curdatas=np.zeros(self.tagCount)
        line=self.ax.lines[0]
        xdatas=line.get_xdata()
        if self.XisTime: #dtype is datetime
            xdatas=matplotlib.dates.date2num(xdatas)
        if(xdatas.size<1):
            self.Tableview_update(curdatas, self.curCursorIdx)
            return -1
        if event.xdata<xdatas[0]:
            idx1 = 0
            idx2 = 1
        elif event.xdata>xdatas[xdatas.size-1]:
            idx1 = xdatas.size - 2
            idx2 = xdatas.size - 1
        else:
            k=0
            for xx in xdatas:
                if xx>event.xdata:
                    idx1=k-1
                    idx2=k
                    break
                k+=1
        xs=np.array([xdatas[idx1],xdatas[idx2]])
        k=0
        phidatas=[]
        for i in range(self.tagCount):
            line=self.ax.lines[i]
            ys=np.array([(line.get_ydata())[idx1],(line.get_ydata())[idx2]])
            if self.drawStyle==1 and self.dat_yphiflag[i]==1:
                line2=self.ax_angle.lines[k]
                k=k+1
                ys2=np.array([(line2.get_ydata())[idx1],(line2.get_ydata())[idx2]])
            if np.abs(xs[0]-xs[1])<0.000001:
                curdatas[i] =ys[0]/100*(self.dat_ymax[i]-self.dat_ymin[i])+self.dat_ymin[i]
                if self.drawStyle==1 and self.dat_yphiflag[i] == 1:
                    phidatas.append(ys2[0])
            else:
                curdatas[i]=ys[0]+(event.xdata-xs[0])*(ys[1]-ys[0])/(xs[1]-xs[0])
                curdatas[i]=curdatas[i]/100*(self.dat_ymax[i]-self.dat_ymin[i])+self.dat_ymin[i]
                if self.drawStyle==1 and self.dat_yphiflag[i] == 1:
                    phidatas.append(ys2[0]+(event.xdata-xs[0])*(ys2[1]-ys2[0])/(xs[1]-xs[0]))
        phidatas=np.asarray(phidatas)
        self.Tableview_update(curdatas,self.curCursorIdx,phidatas)
        return event.xdata

    def Tableview_init(self):
        self.model_ColCount=9
        self.model_RowCount=1
        self.model=QtGui.QStandardItemModel(self.tableView)
        self.model.setColumnCount(self.model_ColCount)
        self.model.setHeaderData(0,QtCore.Qt.Horizontal,_fromUtf8(u"名称"))
        self.model.setHeaderData(1,QtCore.Qt.Horizontal,_fromUtf8(u"下限"))
        self.model.setHeaderData(2,QtCore.Qt.Horizontal,_fromUtf8(u"上限"))
        self.model.setHeaderData(3,QtCore.Qt.Horizontal,_fromUtf8(u"颜色"))
        self.model.setHeaderData(4,QtCore.Qt.Horizontal,_fromUtf8(u"单位"))
        self.model.setHeaderData(5,QtCore.Qt.Horizontal,_fromUtf8(u"游标1值"))
        self.model.setHeaderData(6,QtCore.Qt.Horizontal,_fromUtf8(u"游标2值"))
        self.model.setHeaderData(7,QtCore.Qt.Horizontal,_fromUtf8(u"相位游标1值"))
        self.model.setHeaderData(8,QtCore.Qt.Horizontal,_fromUtf8(u"相位游标2值"))
        self.tableView.setModel(self.model)
        self.tableView.setColumnWidth(0,300)

    def Tableview_setData(self):
        rowCount=self.tagCount
        if rowCount>0:
            self.model_RowCount =rowCount
            self.model.setRowCount(rowCount)
            for i in range(rowCount):
                self.tableView.setRowHeight(i, 20)
                self.model.setData(self.model.index(i, 0, QtCore.QModelIndex()), QtCore.QVariant(_fromUtf8(u"{}".format(self.dat_tags[i]))))
                self.model.setData(self.model.index(i, 1, QtCore.QModelIndex()), QtCore.QVariant(_fromUtf8(u"{}".format(self.dat_ymin[i]))))
                self.model.setData(self.model.index(i, 2, QtCore.QModelIndex()), QtCore.QVariant(_fromUtf8(u"{}".format(self.dat_ymax[i]))))
                self.model.setItem(i,3,QtGui.QStandardItem(""))
                (r,g,b)=matplotlib.colors.ColorConverter().to_rgb(self.dat_ycolor[i])
                self.model.item(i, 3).setBackground(QtGui.QBrush(QtGui.QColor(255*r, 255*g, 255*b)))
                self.model.setData(self.model.index(i, 5, QtCore.QModelIndex()), QtCore.QVariant(_fromUtf8(u"")))
                self.model.setData(self.model.index(i, 6, QtCore.QModelIndex()), QtCore.QVariant(_fromUtf8(u"")))
                self.model.setData(self.model.index(i, 7, QtCore.QModelIndex()), QtCore.QVariant(_fromUtf8(u"")))
                self.model.setData(self.model.index(i, 8, QtCore.QModelIndex()), QtCore.QVariant(_fromUtf8(u"")))

            self.tableView.setSelectionBehavior(1)  # 0  Selecting single items.选中单个单元格
                                                    # 1  Selecting only rows.选中一行
                                                     # 2  Selecting only columns.选中一列
            # self.tableView.setItemDelegateForColumn(0,11)
            # self.tableView.openPersistentEditor()

    def Tableview_update(self,curdatas,cursorIdx,phidatas):
        rowCount=self.tagCount
        k=0
        if rowCount>0 and rowCount==curdatas.size:
            for i in range(rowCount):
                self.model.setData(self.model.index(i, 5+cursorIdx, QtCore.QModelIndex()),QtCore.QVariant(_fromUtf8(u"{0:.2f}".format(curdatas[i]))))
                if self.drawStyle==1 and self.dat_yphiflag[i]==1:
                    self.model.setData(self.model.index(i, 7 + cursorIdx, QtCore.QModelIndex()),
                                      QtCore.QVariant(_fromUtf8(u"{0:.2f}".format(phidatas[k]))))#增加相位值
                    k=k+1

    def Canvas_init(self):
        self.fig = Figure(figsize=(7, 7), dpi=72, facecolor=(1, 1, 1), edgecolor=(0, 0, 0))
        if self.drawStyle==0:
            self.ax = self.fig.add_subplot(111)
        elif self.drawStyle==1:
            self.ax = self.fig.add_subplot(2,1,2)
            self.ax_angle = self.fig.add_subplot(2,1,1,sharex=self.ax)
        self.canvas=FigureCanvas(self.fig)
        self.canvas.setSizePolicy(QtGui.QSizePolicy.Expanding,QtGui.QSizePolicy.Expanding,)
        self.canvas.setParent(self.frame)
        self.layout = QtGui.QVBoxLayout(self.frame)
        self.toolbar = NavigationToolbar(self.canvas, self.frame)
        self.layout.addWidget(self.canvas)
        #self.layout.addWidget(self.toolbar)

    def setupUi(self, MainWindow):
        MainWindow.setObjectName(_fromUtf8("MainWindow"))
        MainWindow.resize(1350,690)
        self.centralwidget = QtGui.QWidget(MainWindow)
        self.centralwidget.setObjectName(_fromUtf8("centralwidget"))
        self.tableView = QtGui.QTableView(self.centralwidget)
        self.tableView.setGeometry(QtCore.QRect(0, 480, 1350,210))
        self.tableView.setObjectName(_fromUtf8("tableView"))
        self.frame = QtGui.QFrame(self.centralwidget)
        self.frame.setGeometry(QtCore.QRect(0, 0,1350,480))
        self.frame.setFrameShape(QtGui.QFrame.StyledPanel)
        self.frame.setFrameShadow(QtGui.QFrame.Raised)
        self.frame.setObjectName(_fromUtf8("frame"))
        MainWindow.setCentralWidget(self.centralwidget)
        self.statusBar = QtGui.QStatusBar(MainWindow)
        self.statusBar.setObjectName(_fromUtf8("statusBar"))
        MainWindow.setStatusBar(self.statusBar)

    def Members_init(self,drawStyle):
        self.drawStyle=drawStyle
        self.dat_tags=np.array([])
        self.dat_ymin=np.array([])
        self.dat_ymax=np.array([])
        self.dat_ycolor=np.array([])
        self.dat_yphiflag=np.array([])
        self.tagCount=0
        self.ax_angle=None
        self.XisTime=True
        self.myVline=[None,None]
        self.myVline_angle=[None,None]
        self.cursor = [-1, -1]
        self.curCursorIdx = 0
        self.press=False
        self.pickedline = None
        self.Cursor_TextObjs=[None,None]

def get_qt_app():
    """
    returns the global QtGui.QApplication instance and starts
    the event loop if necessary.
    """
    app = QtCore.QCoreApplication.instance()
    if app is None:
        # create a new application
        app = QtGui.QApplication([])
        # use timer to process events periodically
        processing_events = {}
        def qt_timer_callback(timer_id, time):
            if timer_id in processing_events:
                return
            processing_events[timer_id] = True
            try:
                app = QtCore.QCoreApplication.instance()
                if app is not None:
                    app.processEvents(QtCore.QEventLoop.AllEvents, 300)
            finally:
                del processing_events[timer_id]
        timer.set_timer(100, qt_timer_callback)
    return app

_plot_windows = {}

@xw.func
@xw.arg('figname')
@xw.arg('x', np.array,ndim=1)
@xw.arg('ys', np.array,ndim=2)
@xw.arg('Tags', np.array,ndim=1)
@xw.arg('ysMin', np.array,ndim=1)
@xw.arg('ysMax', np.array,ndim=1)
@xw.arg('ysColor', np.array,ndim=1)
@xw.arg('ysPhiFlag', np.array,ndim=1)
@xw.arg('drawStyle')
def mzsplot_Trend(figname, x, ys,Tags, ysMin,ysMax, ysColor,ysPhiFlag,drawStyle=0):
    app = get_qt_app()
    if figname in _plot_windows:
        window = _plot_windows[figname]
    else:
        window = Ui_MainWindow(drawStyle)
        _plot_windows[figname] = window
        window.setWindowTitle(figname)
    # plot the data
    window.drawStyle=drawStyle
    if x.dtype.name!='float64':
        window.XisTime=True
    else:
        window.XisTime=False
    window.tagCount=Tags.size
    window.dat_tags=Tags
    window.dat_ymin=ysMin
    window.dat_ymax=ysMax
    window.dat_ycolor=ysColor
    window.dat_yphiflag=ysPhiFlag
    window.Tableview_setData()
    n,m=ys.shape
    window.ax.cla()
    if drawStyle==1:
        window.ax_angle.cla()
    window.myVline_angle=[None,None]
    window.myVline=[None,None]
    window.cursor=[-1,-1]
    window.curCursorIdx=0
    window.press=False
    window.pickedline=None
    window.Cursor_TextObjs=[None,None]
    for i in range(m):
        y=ys[:,i]
        ymin=ysMin[i]
        ymax=ysMax[i]
        ycolor=ysColor[i]
        tag=u""+_fromUtf8(Tags[i])
        if window.drawStyle == 1 and ysPhiFlag[i] == 1:
            y = np.asarray([yi.split(u'∠') for yi in y])
            yA = (y[:, 0]).astype(np.float64)
            yPhi = (np.asarray([yi.split(u'°') for yi in y[:, 1]])[:, 0]).astype(np.float64)
            yA=(yA - ymin) / (ymax - ymin) * 100
            line, = window.ax.plot(x, yA, alpha=0.5, label=tag, color=ycolor, picker=5)
            window.ax_angle.plot(x, yPhi, alpha=0.5, label=tag, color=ycolor, picker=5)
        else:
            if y.dtype.name != 'float64':
                y=y.astype(np.float64)
            y = (y - ymin) / (ymax - ymin) * 100
            line,=window.ax.plot(x, y, alpha=0.5, label=tag,color=ycolor,picker=5)
    window.ax.set_ylim(0,100)
    window.ax.set_ylabel('%',fontsize=24)
    if drawStyle==1:
        window.ax_angle.set_ylabel(u'相位', fontsize=16)
    for label in window.ax.xaxis.get_ticklabels():
        label.set_rotation(0)
    window.fig.tight_layout()
    window.canvas.draw()
    window.show()
    return 0

if __name__=='__main__':
    import sys
    app = QtGui.QApplication(sys.argv)
    x=np.linspace(0,2*np.pi,100)
    ys=np.transpose(np.asarray([np.sin(x), np.cos(x)]))
    Tags=np.asarray(["sin","cos"])
    ysMin=np.asarray([-2,-2])
    ysMax=np.asarray([2,2])
    ysColor=np.asarray(["#FF00FF","#FF0000"])
    ysPhiFlag=np.asarray([0,0])
    drawStyle=0
    mzsplot_Trend("aaa", x, ys ,Tags, ysMin,ysMax, ysColor,ysPhiFlag,drawStyle)
    sys.exit(app.exec_())
