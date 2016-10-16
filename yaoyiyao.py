#coding=utf-8
import win32api
import time
import win32con
import ImageGrab
import win32clipboard as w


def paste():
    win32api.keybd_event(17,0,0,0)
    win32api.keybd_event(86,0,0,0)
    win32api.keybd_event(86,0,win32con.KEYEVENTF_KEYUP,0)
    win32api.keybd_event(17,0,win32con.KEYEVENTF_KEYUP,0)
        
def getText():
    w.OpenClipboard()
    d = w.GetClipboardData(win32con.CF_TEXT)
    w.CloseClipboard()
    return d
 
def setText(aString):
    w.OpenClipboard()
    w.EmptyClipboard()
    w.SetClipboardData(win32con.CF_TEXT, aString.encode('gb18030'))
    w.CloseClipboard()


def color(pos):
    im = ImageGrab.grab()
    pixel = im.getpixel(pos)
    r, g, b = pixel
    print r,g,b #113
    count=r+g+b
    print count
    return count

#475 85 坐标
time.sleep(10)
for one in range(10000):
    #点击摇一摇
    win32api.SetCursorPos((396,81))
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
    #检测摇一摇结果
    time.sleep(4)
    win32api.SetCursorPos((53,545))
    numm=0
    while True:
        x,y=win32api.GetCursorPos()
        print x,y
        numm=numm+1
        print numm
        if color((x,y))==130:
            print "没有头像"
            if numm>1002:
                break
            continue
        else:
            print '有头像，开始下一步'
            break
    if  numm>1000:
        print "下一次循环"
    else:
        #点击摇一摇结果
        time.sleep(1)
        win32api.SetCursorPos((205,545))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        
        #点击打招呼
        time.sleep(1)
        win32api.SetCursorPos((159,358))
        while True:
            x,y=win32api.GetCursorPos()
            print x,y
            color((x,y))
            if color((x,y))!=224:
                print "没有打招呼按钮"
                
                continue
            else:
                print '有打招呼按钮，下一步'
                break
        win32api.SetCursorPos((159,358))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        
        #输入广告语
        time.sleep(5)
        win32api.SetCursorPos((364,75))
        while True:
            x,y=win32api.GetCursorPos()
            print x,y
            color((x,y))
            if color((x,y))!=224:
                print "没有发送按钮"
                
                continue
            else:
                print '有打发送按钮，下一步'
                break
            
        win32api.SetCursorPos((39,140))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        time.sleep(2)
        setText(u"【微信搜索关注】【种子BT资源】10000部日本小电影随便看，总有你要的口味，你懂的")
        paste()
                
        #点击发送 
        time.sleep(2)
        win32api.SetCursorPos((346,75))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        
        #发送完成，检测完成页面
        time.sleep(5)
        win32api.SetCursorPos((155,356))
        while True:
            x,y=win32api.GetCursorPos()
            print x,y
            color((x,y))
            if color((x,y))!=224:
                print "没有打招呼按钮"
                
                continue
            else:
                print '有打招呼按钮'
                break
        #点击返回
        time.sleep(2)
        win32api.SetCursorPos((17,73))
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0,0,0)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0,0,0)
        print "完成第%s次摇一摇" % one
        time.sleep(10)
    
        
