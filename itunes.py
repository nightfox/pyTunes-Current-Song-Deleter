#Python 2.7
#Author n!ghtf0x
#email anirvan.mandal@gmail.com
#Script for Deleting the Current playing song in iTunes from HDD and iTunes Music Library

import win32com.client
import xml.dom.minidom
import string
import os,urllib
import wx
import shelve



class refreshDone(wx.Frame):
		def __init__(self):
				wx.Frame.__init__(self, None, title='iTunes', pos=(600,350), size=(160,100), style = wx.DEFAULT_MINIFRAME_STYLE)
				self.pan = wx.Panel(self,-1)
				self.text3 = wx.StaticText(self.pan,5,'Hashfile updated !',pos = (25,7))
				self.button1 = wx.Button(self.pan,6,'Ok',(30,30))
				self.Bind(wx.EVT_BUTTON, self.close, self.button1)
				self.Show()
		def close(self,event):
				self.Close()
				f = Window()
				
				
class WindowE(wx.Frame):
		def __init__(self):
				wx.Frame.__init__(self, None, title='Confirmation Box', pos=(500,300), size=(350,200))
				panel = wx.Panel(self, -1)
				self.text = wx.StaticText(panel,2,'iTunes is currently not running',pos = (30,20))
				button = wx.Button(panel, 1, 'Exit', (230, 120))
				self.Bind(wx.EVT_BUTTON, self.close, button)
				button.Show()
				
				self.Show()
		def close(self,event):
				self.Close()

class Window(wx.Frame):
		def __init__(self):
				wx.Frame.__init__(self, None, title='Confirmation Box', pos=(500,300), size=(350,200))
				self.SetSizeHints(350,200,350,200) 
				panel = wx.Panel(self, -1)
				self.text = wx.StaticText(panel,2,'iTunes is not running but the process is still active ! Please restart script',pos = (30,20))
				panel.SetBackgroundColour('White')
				iconFile = "it.ico"
				icon1 = wx.Icon(iconFile, wx.BITMAP_TYPE_ICO)
				self.SetIcon(icon1)
				self.button = wx.Button(panel, 1, 'Delete', (230, 120))
				self.Bind(wx.EVT_BUTTON, self.delSong, self.button)
				self.button2 = wx.Button(panel, 3, 'Refresh', (130, 120))
				self.Bind(wx.EVT_BUTTON, self.shelving, self.button2)
				self.button3 = wx.Button(panel, 4, 'Exit', (230, 120))
				self.Bind(wx.EVT_BUTTON, self.close, self.button3)
				self.button3.Disable()
				self.button.Disable()
				self.button2.Enable()
				
				self.Show()
				self.iTunes = win32com.client.gencache.EnsureDispatch("iTunes.Application")
				try:
						self.curTrackName = self.iTunes.CurrentTrack.Name
						self.Artist = self.iTunes.CurrentTrack.Artist
						self.Album = self.iTunes.CurrentTrack.Album
						self.file = shelve.open('hashdata')
						self.locationfile = self.file[self.curTrackName.encode('ascii','ignore')]
						self.a = urllib.unquote(self.locationfile)
						print self.a
						self.b = self.a[17:len(self.a)]
						self.text.Hide()
						
						self.msg = 'Song : '+ self.curTrackName + '\nArtist : '+self.Artist+'\nAlbum : '+self.Album+'\n\nDo You Really want to delete this song from the hard disk and iTunes library ?'
						self.text1 = wx.StaticText(panel,2,self.msg,pos = (30,10),size=(290,300))
						self.text1.Show()
						self.button.Enable()
						panel.Update()
						
						
				except AttributeError:
						self.text.Hide()
						self.msg = 'iTunes is currently is stopped State. Try Again Later'
						self.text1 = wx.StaticText(panel,2,self.msg,pos = (30,20),size=(290,300))
						self.text1.Show()
						self.button.Destroy()
						self.button3.Enable()
										
				except	KeyError:
						self.text.Hide()
						self.msg = 'The Hash Table is Not updated. Please Click Refresh'
						self.text1 = wx.StaticText(panel,2,self.msg,pos = (30,20),size=(290,300))
						self.text1.Show()
						self.button.Destroy()
						self.button3.Enable()
		def delSong(self,event):
				self.iTunes.CurrentTrack.Delete()
				print 'a'
				self.iTunes.Play()
				os.remove(self.b)
				self.Close()
		
		def shelving(self,event):
				self.file1 = shelve.open('hashdata')
				self.path = self.iTunes.LibraryXMLPath
				self.file = open(self.path,'r')
				#print self.curTrackName.Delete()
				self.dom = xml.dom.minidom.parse(self.file)
				for dct in self.dom.getElementsByTagName('dict'):
						keys=dct.getElementsByTagName('key')
						vals=[key.nextSibling.firstChild for key in keys]
						keys=[key.firstChild.data for key in keys]
						vals=[val.data if val else None for val in vals]
						data=dict(zip(keys,vals))
						try :
								self.file1[data['Name'].encode('ascii','ignore')]= data['Location'].encode('ascii','ignore')
						except KeyError:
								continue
				print 'done'
				box = refreshDone()
				self.Close()
				self.file1.close()
		
		def close(self,event):
				self.Close()
				

if __name__ == '__main__' :		
		cmd = os.popen('query process')
		x = cmd.readlines()
		p = 0
		for y in x:
				p = y.find('itunes.exe')
				if p > 0:
						break
				
		if p > 0 :		
				app = wx.App(False)
				#x = curSong()
				f = Window()
				#f.Show()
				app.MainLoop()
		else :
				app = wx.App(False)
				f =WindowE()
				app.MainLoop()