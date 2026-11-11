import json, math, os, queue, re, threading, tkinter as tk, urllib.parse, urllib.request, webbrowser
from tkinter import ttk, messagebox

def parse_coordinates(text):
    parts = re.split(r'[,;\s]+', text.strip())
    if len(parts) != 2:
        raise ValueError('Enter latitude, longitude in decimal degrees, for example 13.7563, 100.5018.')
    lat, lon = map(float, parts)
    if not math.isfinite(lat) or not math.isfinite(lon) or not -90 <= lat <= 90 or not -180 <= lon <= 180:
        raise ValueError('Latitude must be -90 to 90 and longitude -180 to 180.')
    return lat, lon

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Coordinates to Address')
        left, top, right, bottom = 0, 0, self.winfo_screenwidth(), self.winfo_screenheight()
        if os.name == 'nt':
            import ctypes
            from ctypes import wintypes
            area = wintypes.RECT()
            if ctypes.windll.user32.SystemParametersInfoW(48, 0, ctypes.byref(area), 0):
                left, top, right, bottom = area.left, area.top, area.right, area.bottom
        width = min(760, right-left-32); height = min(640, bottom-top-64)
        self.geometry(f'{width}x{height}+{left+(right-left-width)//2}+{top+(bottom-top-height)//2}')
        self.minsize(min(580,width), min(460,height))
        self.configure(bg='#111827')
        self.events = queue.Queue(); self.coords = None; self.address = ''; self.report = ''
        style = ttk.Style(self); style.theme_use('clam')
        style.configure('TFrame', background='#111827')
        style.configure('TLabel', background='#111827', foreground='#e5e7eb', font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10), padding=(10, 5), background='#374151', foreground='#f3f4f6', bordercolor='#4b5563')
        style.map('TButton', background=[('active', '#4b5563'), ('disabled', '#1f2937')], foreground=[('disabled', '#6b7280')])
        style.configure('Vertical.TScrollbar', background='#374151', troughcolor='#111827', arrowcolor='#d1d5db')
        style.configure('Accent.TButton', background='#2563eb', foreground='white')
        style.map('Accent.TButton', background=[('active', '#2563eb')])
        style.configure('TEntry', padding=5, fieldbackground='#1f2937', foreground='#f3f4f6', insertcolor='#f3f4f6')
        panel = ttk.Frame(self, padding=16); panel.pack(fill='both', expand=True)
        ttk.Label(panel, text='Coordinates to Address', font=('Segoe UI', 19, 'bold')).pack(anchor='w', pady=(0, 10))
        ttk.Label(panel, text='Coordinates (latitude, longitude)', font=('Segoe UI',10,'bold')).pack(anchor='w')
        row = ttk.Frame(panel); row.pack(fill='x', pady=(6,4))
        self.entry = ttk.Entry(row, font=('Segoe UI',12)); self.entry.pack(side='left',fill='x',expand=True); self.entry.focus()
        self.button = ttk.Button(row, text='Find address', style='Accent.TButton', command=self.lookup); self.button.pack(side='left',padx=(12,0))
        settings = ttk.Frame(panel); settings.pack(fill='x',pady=(4,0))
        ttk.Label(settings, text='API key').pack(side='left',padx=(0,8))
        self.key = ttk.Entry(settings, show='*', takefocus=False); self.key.insert(0, os.getenv('GEOCODING_KEY','')); self.key.pack(side='left',fill='x',expand=True)
        self.status = tk.StringVar(value='Ready to look up a location')
        ttk.Label(panel,textvariable=self.status,foreground='#9ca3af').pack(anchor='w',pady=(6,6))
        card = tk.Frame(panel,bg='#1f2937',highlightbackground='#374151',highlightthickness=1,padx=12,pady=10); card.pack(fill='x')
        tk.Label(card,text='PROBABLE ADDRESS',bg='#1f2937',fg='#60a5fa',font=('Segoe UI',10,'bold')).pack(anchor='w')
        self.address_var = tk.StringVar(value='')
        self.address_label = tk.Label(card,textvariable=self.address_var,bg='#1f2937',fg='#f3f4f6',font=('Segoe UI',14,'bold'),justify='left',anchor='w',wraplength=760)
        def update_address_area(*_):
            if self.address_var.get():
                self.address_label.pack(fill='x',pady=(4,6),before=actions)
            else:
                self.address_label.pack_forget()
        self.address_var.trace_add('write', update_address_area)
        card.bind('<Configure>',lambda e:self.address_label.configure(wraplength=max(300,e.width-40)))
        actions=tk.Frame(card,bg='#1f2937'); actions.pack(fill='x',pady=(6,0))
        self.copy_address=ttk.Button(actions,text='Copy address',command=lambda:self.copy(True),state='disabled'); self.copy_address.pack(side='left')
        self.copy_details=ttk.Button(actions,text='Copy all details',command=self.copy,state='disabled'); self.copy_details.pack(side='left',padx=8)
        ttk.Button(actions,text='Open in Maps',command=self.maps).pack(side='left')
        ttk.Label(panel,text='Address & location details',font=('Segoe UI',11,'bold')).pack(anchor='w',pady=(10,6))
        detail=ttk.Frame(panel); detail.pack(fill='both',expand=True)
        self.output = tk.Text(detail, height=6, wrap='word', font=('Segoe UI',11), padx=18,pady=14,bg='#1f2937',fg='#d1d5db',insertbackground='#f3f4f6',selectbackground='#2563eb',selectforeground='white',relief='flat',state='disabled',spacing1=3,spacing3=4)
        self.output.pack(side='left',fill='both',expand=True)
        scroll=ttk.Scrollbar(detail,command=self.output.yview); scroll.pack(side='right',fill='y'); self.output.configure(yscrollcommand=scroll.set)
        self.output.tag_configure('heading',font=('Segoe UI',10,'bold'),foreground='#60a5fa',spacing1=12)
        ttk.Label(panel,text='Google Maps | Address matches may be approximate.',foreground='#9ca3af').pack(anchor='w',pady=(10,0))
        def focus_next(widget):
            widget.focus_set()
            return 'break'
        def activate(command):
            command()
            return 'break'
        self.entry.bind('<Tab>', lambda e: focus_next(self.button))
        self.button.bind('<Tab>', lambda e: focus_next(self.copy_address) if self.copy_address.instate(['!disabled']) else focus_next(self.entry))
        self.copy_address.bind('<Shift-Tab>', lambda e: focus_next(self.button))
        self.entry.bind('<Return>', lambda e: activate(self.lookup))
        self.button.bind('<Return>', lambda e: activate(self.lookup))
        self.copy_address.bind('<Return>', lambda e: activate(lambda: self.copy(True)))
        self.copy_details.bind('<Return>', lambda e: activate(self.copy))
        self.after(100,self.poll)
    def copy(self, address_only=False):
        text=self.address if address_only else self.report
        if text:
            self.clipboard_clear(); self.clipboard_append(text)
            self.status.set('Address copied' if address_only else 'Full details copied')
    def maps(self):
        try: lat,lon=parse_coordinates(self.entry.get())
        except ValueError as e: messagebox.showerror('Coordinates',str(e)); return
        webbrowser.open(f'https://www.google.com/maps/search/?api=1&query={lat},{lon}')
    def lookup(self):
        if str(self.button['state']) == 'disabled': return
        try:
            coords=parse_coordinates(self.entry.get()); key=self.key.get().strip()
            if not key: raise ValueError('Enter a Google Geocoding API key, or configure GEOCODING_KEY as in Image Analyzer.')
        except ValueError as e: messagebox.showerror('Check input',str(e)); return
        self.button.config(state='disabled'); self.status.set('Looking up address...')
        self.address_var.set('Searching...'); self.address = ''; self.report = ''
        self.copy_address.config(state='disabled'); self.copy_details.config(state='disabled')
        threading.Thread(target=self.fetch,args=(coords,key),daemon=True).start()
    def fetch(self,coords,key):
        try:
            params=urllib.parse.urlencode({'latlng':f'{coords[0]},{coords[1]}','key':key,'language':'en'})
            with urllib.request.urlopen('https://maps.googleapis.com/maps/api/geocode/json?'+params,timeout=20) as response: data=json.load(response)
            if data.get('status') == 'ZERO_RESULTS': raise ValueError('No address found for these coordinates.')
            if data.get('status') != 'OK': raise ValueError('Google lookup failed: '+data.get('status','Unknown status')+'. Check API enablement, billing, and key restrictions.')
            results=data.get('results',[]); best=results[0]
            lines=[f'Coordinates: {coords[0]}, {coords[1]}','', 'PROBABLE ADDRESS',best.get('formatted_address','Unavailable'),'','ADDRESS DETAILS']
            for component in best.get('address_components',[]):
                label=next((t for t in component.get('types',[]) if t != 'political'), 'Address detail').replace('_',' ').title()
                lines.append(f"{label}: {component['long_name']}")
            lines += ['', 'LOCATION DETAILS', 'Accuracy: '+best.get('geometry',{}).get('location_type','Unknown'), 'Place ID: '+best.get('place_id','Unavailable')]
            plus=best.get('plus_code') or data.get('plus_code',{})
            if plus: lines.append('Plus code: '+plus.get('compound_code',plus.get('global_code','')))
            lines += ['', 'OTHER MATCHES']+[r.get('formatted_address','') for r in results[1:]]
            self.events.put((True,(best.get('formatted_address','Unavailable'),'\n'.join(lines))))
        except ValueError as e: self.events.put((False,str(e)))
        except Exception: self.events.put((False,'Could not reach Google. Check your internet connection and try again.'))
    def poll(self):
        try:
            ok,text=self.events.get_nowait(); self.button.config(state='normal'); self.status.set('Lookup complete' if ok else 'Lookup failed')
            if ok:
                self.address,self.report=text; self.address_var.set(self.address)
                self.copy_address.config(state='normal'); self.copy_details.config(state='normal')
                text=self.report.split('ADDRESS DETAILS',1)[-1].lstrip()
                text='ADDRESS DETAILS\n'+text
            else: self.address_var.set('Address unavailable')
            self.output.config(state='normal'); self.output.delete('1.0','end')
            for line in text.splitlines():
                self.output.insert('end',line+'\n', 'heading' if line in ('ADDRESS DETAILS','LOCATION DETAILS','OTHER MATCHES') else ())
            self.output.config(state='disabled')
        except queue.Empty: pass
        self.after(100,self.poll)

if __name__ == '__main__': App().mainloop()
