# 桌面待办
# 双击 DeskTodo.vbs 启动（不会出现黑窗口）。数据保存在 %APPDATA%\DeskTodo\todos.json

$ErrorActionPreference = 'Stop'
$AppVersion = '2026.11.13'

Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase, System.Windows.Forms, System.Drawing

$DataDir  = Join-Path $env:APPDATA 'DeskTodo'
$SetFile  = Join-Path $DataDir 'settings.json'
$LogFile  = Join-Path $DataDir 'error.log'
$StartupLnk = Join-Path ([Environment]::GetFolderPath('Startup')) 'DeskTodo.lnk'
$AppFile  = $PSCommandPath
$AppDir   = Split-Path -Parent $PSCommandPath
$MirrorDir = Join-Path $DataDir 'mirror'   # 开着同步时，本机另存一份，网盘覆盖了也能合并回来
$Utf8 = New-Object System.Text.UTF8Encoding($false)
$Fmt  = 'yyyy-MM-dd HH:mm'
$Inv  = [Globalization.CultureInfo]::InvariantCulture
$WeekNames = @('周日','周一','周二','周三','周四','周五','周六')

if (-not (Test-Path $DataDir)) { New-Item -ItemType Directory -Force -Path $DataDir | Out-Null }

function Write-Log([string]$msg) {
    try { [IO.File]::AppendAllText($LogFile, ((Get-Date).ToString('yyyy-MM-dd HH:mm:ss') + '  ' + $msg + "`r`n"), $Utf8) } catch {}
}

# ---------- 只允许运行一个 ----------
$created = $false
$script:Mutex = [System.Threading.Mutex]::new($true, 'DeskTodo_SingleInstance_v1', [ref]$created)
if (-not $created) {
    try { $ev0 = [System.Threading.EventWaitHandle]::OpenExisting('DeskTodo_Show_v1'); [void]$ev0.Set(); exit } catch {}
    [System.Windows.MessageBox]::Show('桌面待办已经在运行了，按 Ctrl+Alt+N 呼出，或点右下角托盘里的蓝色勾图标。', '桌面待办') | Out-Null
    exit
}
$script:ShowEvent = [System.Threading.EventWaitHandle]::new($false, [System.Threading.EventResetMode]::AutoReset, 'DeskTodo_Show_v1')

# ---------- 全局热键 ----------
if (-not ('DeskTodoHotKey' -as [type])) {
    Add-Type -ReferencedAssemblies System.Windows.Forms -TypeDefinition @'
using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
public class DeskTodoHotKey : NativeWindow, IDisposable {
    [DllImport("user32.dll")] static extern bool RegisterHotKey(IntPtr hWnd, int id, uint fsModifiers, uint vk);
    [DllImport("user32.dll")] static extern bool UnregisterHotKey(IntPtr hWnd, int id);
    public event EventHandler Pressed;
    public bool Ok;
    public DeskTodoHotKey(uint mod, uint vk) {
        CreateHandle(new CreateParams());
        Ok = RegisterHotKey(Handle, 1, mod, vk);
    }
    protected override void WndProc(ref Message m) {
        if (m.Msg == 0x0312 && Pressed != null) Pressed(this, EventArgs.Empty);
        base.WndProc(ref m);
    }
    public void Dispose() {
        UnregisterHotKey(Handle, 1);
        DestroyHandle();
    }
}
'@
}

# ---------- 数据 ----------
$script:Items = New-Object System.Collections.ArrayList

function Parse-Dt([string]$s) {
    if ([string]::IsNullOrWhiteSpace($s)) { return $null }
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, $Fmt, $Inv, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function New-Todo([string]$text) {
    [pscustomobject]@{
        id = [guid]::NewGuid().ToString('N'); text = $text
        created = (Get-Date).ToString($Fmt); due = ''
        times = @(); fired = @(); warned = $false
        done = $false; doneAt = ''; star = $false; notes = @(); subs = @(); upd = [long]0
        project = ''; start = ''; amount = ''; rstate = ''; copied = $false
    }
}

function ConvertTo-Notes($raw, [string]$created) {
    if ($null -eq $raw) { return }
    if ($raw -is [string]) {
        # 旧版备注是一整段文字
        if ($raw.Trim()) { [pscustomobject]@{ id = [guid]::NewGuid().ToString('N'); kind = 'text'; text = $raw.Trim(); file = ''; at = $created } }
        return
    }
    foreach ($n in @($raw)) {
        if (-not $n) { continue }
        $kind = if ([string]$n.kind -eq 'image') { 'image' } else { 'text' }
        if ($kind -eq 'text' -and -not [string]$n.text) { continue }
        if ($kind -eq 'image' -and -not [string]$n.file) { continue }
        $nid = [string]$n.id; if (-not $nid) { $nid = [guid]::NewGuid().ToString('N') }
        [pscustomobject]@{ id = $nid; kind = $kind; text = [string]$n.text; file = [string]$n.file; at = [string]$n.at }
    }
}

function ConvertTo-Todo($o) {
    $id = [string]$o.id
    if (-not $id) { $id = [guid]::NewGuid().ToString('N') }
    $cr = [string]$o.created
    if (-not $cr) { $cr = (Get-Date).ToString($Fmt) }
    $times = @(@($o.times) | Where-Object { $_ } | ForEach-Object { [string]$_ })
    $fired = @(@($o.fired) | Where-Object { $_ } | ForEach-Object { [string]$_ })
    if ($times.Count -eq 0 -and $o.due) {
        # 旧版只有一个截止时间
        $times = @([string]$o.due)
        if ([bool]$o.n2) { $fired = @([string]$o.due) }
    }
    $t = [pscustomobject]@{
        id = $id; text = [string]$o.text; created = $cr; due = [string]$o.due
        times = $times; fired = $fired; warned = ([bool]$o.warned -or [bool]$o.n1)
        done = [bool]$o.done; doneAt = [string]$o.doneAt; star = [bool]$o.star
        notes = @(ConvertTo-Notes $o.notes $cr)
        upd = [long]$o.upd
        project = [string]$o.project; start = [string]$o.start; amount = [string]$o.amount; rstate = [string]$o.rstate
        copied = [bool]$o.copied
        subs = @(@($o.subs) | Where-Object { $_ -and ($_.text -or $_.file) } | ForEach-Object {
            $sid = [string]$_.id; if (-not $sid) { $sid = [guid]::NewGuid().ToString('N') }
            $sk = if ([string]$_.kind -eq 'image' -and $_.file) { 'image' } else { 'text' }
            [pscustomobject]@{ id = $sid; kind = $sk; text = [string]$_.text; file = [string]$_.file; done = [bool]$_.done; upd = [long]$_.upd } })
    }
    # 备注功能去掉了：旧备注并到具体事项后面，内容不丢
    if (@($t.notes).Count -gt 0) {
        foreach ($n in @($t.notes)) {
            $t.subs = @($t.subs) + [pscustomobject]@{ id = $n.id; kind = $n.kind; text = $n.text; file = $n.file; done = $false }
        }
        $t.notes = @()
    }
    Sync-Times $t
    return $t
}

# 时间点排序去重，截止时间 = 最后一个时间点
function Sync-Times($it) {
    $list = @(@($it.times) | Where-Object { Parse-Dt $_ } | Sort-Object -Unique | Sort-Object { Parse-Dt $_ })
    $it.times = $list
    $it.fired = @(@($it.fired) | Where-Object { $list -contains $_ })
    $newDue = if ($list.Count -gt 0) { [string]$list[-1] } else { '' }
    if ($newDue -ne $it.due) { $it.warned = $false }
    $it.due = $newDue
}

function Get-NextTime($it, [datetime]$now) {
    foreach ($x in @($it.times)) { $d = Parse-Dt $x; if ($d -and $d -ge $now) { return $d } }
    return $null
}

# ---------- 存储与多台电脑同步 ----------
# 数据文件放在 $script:ShareDir（默认本机 %APPDATA%\DeskTodo；设置了同步文件夹就放那里）
# 每条记录带 upd（最后修改时间），删除的记录记在 deleted 里，两台电脑的改动按条合并
$script:ShareDir = $DataDir
$script:SyncWarned = $false
$script:PendingWrite = @{}
$script:FailSince = $null

$script:Recv = @{}       # "电脑id|文件名" -> 收到的那台电脑改动截到的时间
$script:MyChange = @{}   # 文件名 -> 本机最后一次写入的时间
$script:DevDirty = $false
$script:DevWritten = [datetime]::MinValue

function ConvertTo-Hash($o) {
    $h = @{}
    if ($null -eq $o) { return $h }
    if ($o -is [hashtable]) { foreach ($k in $o.Keys) { $h[[string]$k] = $o[$k] }; return $h }
    foreach ($p in $o.PSObject.Properties) { $h[$p.Name] = $p.Value }
    return $h
}
function Get-DeviceId {
    if (-not $script:S.deviceId) { $script:S.deviceId = [guid]::NewGuid().ToString('N').Substring(0, 12) }
    return [string]$script:S.deviceId
}
function Get-DeviceLabel { if ($script:S.deviceLabel) { return [string]$script:S.deviceLabel }; return $env:COMPUTERNAME }
function Note-Recv([string]$file, $r) {
    if (-not $r -or -not $r.vec) { return }
    $me = Get-DeviceId
    foreach ($k in @($r.vec.Keys)) {
        if ($k -eq $me) { continue }
        $key = $k + '|' + $file
        if (-not $script:Recv.ContainsKey($key) -or [long]$script:Recv[$key] -lt [long]$r.vec[$k]) { $script:Recv[$key] = [long]$r.vec[$k] }
    }
}

# 用一小段 C# 算"这条记录的内容指纹"，比 ConvertTo-Json 快很多，记录多了也不卡
$script:FastSig = $false
try {
    if (-not ('DeskTodoSig' -as [type])) {
        Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections;
using System.Globalization;
using System.Management.Automation;
public static class DeskTodoSig {
    public static string Of(object o) { StringBuilder sb = new StringBuilder(); W(sb, o, true); return sb.ToString(); }
    // 每个字段各自的指纹（不含 id、修改时间、子项）
    public static Hashtable Fields(object o) {
        Hashtable h = new Hashtable();
        PSObject pso = o as PSObject;
        if (pso == null) return h;
        foreach (PSPropertyInfo p in pso.Properties) {
            string n = p.Name;
            if (n == "id" || n == "upd" || n == "fu" || n == "subDel" || n == "subs") continue;
            StringBuilder sb = new StringBuilder(); W(sb, p.Value, false); h[n] = sb.ToString();
        }
        return h;
    }
    public static string Json(object o) { StringBuilder sb = new StringBuilder(); J(sb, o); return sb.ToString(); }
    static void Str(StringBuilder sb, string s) {
        sb.Append('"');
        foreach (char c in s) {
            switch (c) {
                case '"': sb.Append("\\\""); break;
                case '\\': sb.Append("\\\\"); break;
                case '\n': sb.Append("\\n"); break;
                case '\r': sb.Append("\\r"); break;
                case '\t': sb.Append("\\t"); break;
                case '\b': sb.Append("\\b"); break;
                case '\f': sb.Append("\\f"); break;
                default:
                    if (c < (char)0x20) sb.Append("\\u").Append(((int)c).ToString("x4"));
                    else sb.Append(c);
                    break;
            }
        }
        sb.Append('"');
    }
    static void J(StringBuilder sb, object o) {
        if (o == null) { sb.Append("null"); return; }
        PSObject pso = o as PSObject;
        if (pso != null) {
            object b = pso.BaseObject;
            if (b is PSCustomObject) {
                sb.Append('{'); bool first = true;
                foreach (PSPropertyInfo p in pso.Properties) {
                    if (!first) sb.Append(','); first = false;
                    Str(sb, p.Name); sb.Append(':'); J(sb, p.Value);
                }
                sb.Append('}'); return;
            }
            o = b;
        }
        if (o is string) { Str(sb, (string)o); return; }
        if (o is bool) { sb.Append((bool)o ? "true" : "false"); return; }
        if (o is int || o is long || o is short || o is byte || o is uint || o is ulong || o is ushort || o is sbyte) { sb.Append(Convert.ToString(o, CultureInfo.InvariantCulture)); return; }
        if (o is double || o is float) { double d = Convert.ToDouble(o, CultureInfo.InvariantCulture); if (double.IsNaN(d) || double.IsInfinity(d)) sb.Append("null"); else sb.Append(d.ToString("R", CultureInfo.InvariantCulture)); return; }
        if (o is decimal) { sb.Append(((decimal)o).ToString(CultureInfo.InvariantCulture)); return; }
        if (o is DateTime) { Str(sb, ((DateTime)o).ToString("o", CultureInfo.InvariantCulture)); return; }
        if (o is Enum) { Str(sb, o.ToString()); return; }
        IDictionary dict = o as IDictionary;
        if (dict != null) {
            sb.Append('{'); bool first = true;
            foreach (DictionaryEntry e in dict) {
                if (!first) sb.Append(','); first = false;
                Str(sb, Convert.ToString(e.Key, CultureInfo.InvariantCulture)); sb.Append(':'); J(sb, e.Value);
            }
            sb.Append('}'); return;
        }
        IEnumerable en = o as IEnumerable;
        if (en != null) {
            sb.Append('['); bool first = true;
            foreach (object x in en) { if (!first) sb.Append(','); first = false; J(sb, x); }
            sb.Append(']'); return;
        }
        Str(sb, Convert.ToString(o, CultureInfo.InvariantCulture));
    }
    static void W(StringBuilder sb, object o, bool top) {
        if (o == null) { sb.Append('~'); return; }
        PSObject pso = o as PSObject;
        if (pso != null) {
            object b = pso.BaseObject;
            if (b is PSCustomObject) {
                sb.Append('{');
                // 按名字排序：字段顺序不同（新建的和从文件读回来的）也算同一个内容
                System.Collections.Generic.List<PSPropertyInfo> ps = new System.Collections.Generic.List<PSPropertyInfo>();
                foreach (PSPropertyInfo p0 in pso.Properties) ps.Add(p0);
                ps.Sort(delegate(PSPropertyInfo x1, PSPropertyInfo x2) { return string.CompareOrdinal(x1.Name, x2.Name); });
                foreach (PSPropertyInfo p in ps) {
                    if (p.Name == "upd") continue;
                    if (top && (p.Name == "fu" || p.Name == "subDel")) continue;
                    sb.Append(p.Name).Append('=');
                    W(sb, p.Value, false);
                    sb.Append('\u0001');
                }
                sb.Append('}');
                return;
            }
            o = b;
        }
        string s = o as string;
        if (s != null) { sb.Append('"').Append(s).Append('"'); return; }
        IEnumerable en = o as IEnumerable;
        if (en != null) {
            sb.Append('[');
            foreach (object x in en) { W(sb, x, false); sb.Append('\u0002'); }
            sb.Append(']');
            return;
        }
        sb.Append(Convert.ToString(o, CultureInfo.InvariantCulture));
    }
}
'@
    }
    $script:FastSig = $true
} catch { Write-Log ('指纹组件编译失败，改用慢速方式：' + $_.Exception.Message) }

function Test-Typing {
    $fe = [System.Windows.Input.Keyboard]::FocusedElement
    if (-not ($fe -is [System.Windows.Controls.TextBox]) -or $fe.IsReadOnly) { return $false }
    if ($List.IsKeyboardFocusWithin) { return $true }
    if ($script:CalWin -and $script:CalWin.IsKeyboardFocusWithin) { return $true }
    return $false
}

function Get-ItemSig($o) {
    if ($script:FastSig) { return [DeskTodoSig]::Of($o) }
    return ($o | Select-Object -Property * -ExcludeProperty upd, fu, subDel | ConvertTo-Json -Compress -Depth 6)
}
function Get-FieldSigs($o) {
    if ($script:FastSig) { return [DeskTodoSig]::Fields($o) }
    $h = @{}
    foreach ($p in $o.PSObject.Properties) { if ('id', 'upd', 'fu', 'subDel', 'subs' -notcontains $p.Name) { $h[$p.Name] = (ConvertTo-Json -InputObject $p.Value -Compress -Depth 4) } }
    return $h
}
function Get-SubSig($x) { return ([string]$x.kind + '|' + [string]$x.text + '|' + [string]$x.file + '|' + [string][bool]$x.done) }
function ConvertTo-JsonFast($o) {
    if ($script:FastSig) { return [DeskTodoSig]::Json($o) }
    return (ConvertTo-Json -InputObject $o -Depth 8)
}

function New-Store([string]$fileName, $list, [scriptblock]$conv) {
    return @{ FileName = $fileName; List = $list; Conv = $conv; Base = @{}; Dead = @{}; Stamp = $null; FBase = @{}; SubBase = @{} }
}

# 每条记录除了整条的修改时间 upd，还记：
#   fu      每个字段最后改的时间（题目、完成、时间、项目……各算各的）
#   subs[].upd  每个具体事项最后改的时间
#   subDel  删掉的具体事项和删的时间
# 两台电脑改了同一条的不同地方，合并时都能保住；只有同一个字段、同一个具体事项两边都改了，才以后改的为准
function Set-Prop($o, [string]$name, $value) {
    if ($o.PSObject.Properties[$name]) { $o.$name = $value } else { Add-Member -InputObject $o -NotePropertyName $name -NotePropertyValue $value -Force }
}
function Get-HashProp($o, [string]$name) {
    $v = if ($o.PSObject.Properties[$name]) { $o.$name } else { $null }
    if ($v -is [hashtable]) { return $v }
    $h = ConvertTo-Hash $v
    Set-Prop $o $name $h
    return $h
}
function Copy-Fine($raw, $x) {
    if ($raw.PSObject.Properties['fu'] -and $raw.fu) { Set-Prop $x 'fu' (ConvertTo-Hash $raw.fu) }
    if ($raw.PSObject.Properties['subDel'] -and $raw.subDel) { Set-Prop $x 'subDel' (ConvertTo-Hash $raw.subDel) }
}
function Remember-Fine($st, $x) {
    $st.FBase[$x.id] = Get-FieldSigs $x
    if ($x.PSObject.Properties['subs']) {
        $h = @{}; foreach ($y in @($x.subs)) { if ($y) { $h[[string]$y.id] = Get-SubSig $y } }
        $st.SubBase[$x.id] = $h
    }
}
function Mark-Fine($st, $x, [long]$now) {
    $fs = Get-FieldSigs $x
    $old = $st.FBase[$x.id]
    if ($old) {
        $fu = $null
        foreach ($k in @($fs.Keys)) { if ($old[$k] -ne $fs[$k]) { if (-not $fu) { $fu = Get-HashProp $x 'fu' }; $fu[$k] = $now } }
        foreach ($k in @($old.Keys)) { if (-not $fs.ContainsKey($k)) { if (-not $fu) { $fu = Get-HashProp $x 'fu' }; $fu[$k] = $now } }
    }
    $st.FBase[$x.id] = $fs
    if ($x.PSObject.Properties['subs']) {
        $sb = $st.SubBase[$x.id]
        $cur = @{}
        foreach ($y in @($x.subs)) {
            if (-not $y) { continue }
            $sg = Get-SubSig $y; $cur[[string]$y.id] = $sg
            if (-not $sb -or $sb[[string]$y.id] -ne $sg -or -not $y.PSObject.Properties['upd']) { Set-Prop $y 'upd' $now }
        }
        if ($sb) {
            foreach ($k in @($sb.Keys)) { if (-not $cur.ContainsKey($k)) { $sd = Get-HashProp $x 'subDel'; $sd[$k] = $now } }
        }
        $st.SubBase[$x.id] = $cur
    }
}
function Get-Ticks($h, [string]$k) { if ($h -and $h.ContainsKey($k)) { return [long]$h[$k] }; return [long]0 }

# 把 remote 合进 cur（cur 就是内存里那一条，直接改它）
function Merge-Pair($cur, $x) {
    $remoteNewer = [long]$x.upd -gt [long]$cur.upd
    $fc = Get-HashProp $cur 'fu'
    $fx = if ($x.PSObject.Properties['fu']) { ConvertTo-Hash $x.fu } else { @{} }
    $names = New-Object System.Collections.Generic.List[string]
    foreach ($p in $cur.PSObject.Properties) { $names.Add($p.Name) }
    foreach ($p in $x.PSObject.Properties) { if (-not $names.Contains($p.Name)) { $names.Add($p.Name) } }
    foreach ($n in $names) {
        if ('id', 'upd', 'fu', 'subDel', 'subs' -contains $n) { continue }
        $tc = Get-Ticks $fc $n; $tx = Get-Ticks $fx $n
        if ($tx -gt $tc -or ($tx -eq $tc -and $remoteNewer)) {
            $v = if ($x.PSObject.Properties[$n]) { $x.$n } else { $null }
            Set-Prop $cur $n $v
            if ($tx -gt 0) { $fc[$n] = $tx }
        }
    }
    if ($cur.PSObject.Properties['subs'] -or $x.PSObject.Properties['subs']) {
        $sd = Get-HashProp $cur 'subDel'
        if ($x.PSObject.Properties['subDel']) { $rx = ConvertTo-Hash $x.subDel; foreach ($k in @($rx.Keys)) { if ((Get-Ticks $sd $k) -lt [long]$rx[$k]) { $sd[$k] = [long]$rx[$k] } } }
        $a = @($cur.subs | Where-Object { $_ }); $b = @($x.subs | Where-Object { $_ })
        $ma = @{}; foreach ($y in $a) { $ma[[string]$y.id] = $y }
        $mb = @{}; foreach ($y in $b) { $mb[[string]$y.id] = $y }
        $first = if ($remoteNewer) { $b } else { $a }
        $second = if ($remoteNewer) { $a } else { $b }
        $out = New-Object System.Collections.ArrayList
        $done = @{}
        foreach ($y in @($first) + @($second)) {
            $k = [string]$y.id
            if ($done.ContainsKey($k)) { continue }
            $done[$k] = $true
            $pick = $y
            if ($ma.ContainsKey($k) -and $mb.ContainsKey($k)) {
                $ua = [long]$ma[$k].upd; $ub = [long]$mb[$k].upd
                $pick = if ($ub -gt $ua -or ($ub -eq $ua -and $remoteNewer)) { $mb[$k] } else { $ma[$k] }
            }
            if ((Get-Ticks $sd $k) -ge [long]$pick.upd -and (Get-Ticks $sd $k) -gt 0) { continue }
            [void]$out.Add($pick)
        }
        $cur.subs = @($out.ToArray())
    }
    if ($remoteNewer) { $cur.upd = [long]$x.upd }
}

function Prune-SubDel($x) {
    if (-not $x.PSObject.Properties['subDel'] -or -not $x.subDel) { return }
    $sd = Get-HashProp $x 'subDel'
    $limit = [DateTime]::UtcNow.AddDays(-90).Ticks
    foreach ($k in @($sd.Keys)) { if ([long]$sd[$k] -lt $limit) { $sd.Remove($k) } }
}

function Get-StorePath($st) { return (Join-Path $script:ShareDir $st.FileName) }

# 网盘客户端上传/下载时会短暂占住文件，读写失败就隔一会儿再试几次
function Invoke-Retry([scriptblock]$sb, [int]$times = 6, [int]$waitMs = 250) {
    for ($i = 1; $i -le $times; $i++) {
        try { return (& $sb) }
        catch { if ($i -eq $times) { throw }; Start-Sleep -Milliseconds $waitMs }
    }
}

# 允许别的程序同时打开（网盘正在上传时也能读）
function Read-SharedText([string]$path) {
    $fs = [IO.File]::Open($path, [IO.FileMode]::Open, [IO.FileAccess]::Read, ([IO.FileShare]::ReadWrite -bor [IO.FileShare]::Delete))
    try { $sr = New-Object IO.StreamReader($fs, $Utf8, $true); return $sr.ReadToEnd() } finally { $fs.Dispose() }
}

function Read-StoreFile([string]$path) {
    # 文件正在被网盘写入时可能只有半截，解析失败也算"暂时读不了"，重试
    return (Invoke-Retry { Read-StoreFileOnce $path })
}

function Read-StoreFileOnce([string]$path) {
    $res = @{ items = @(); deleted = @{}; vec = @{} }
    $raw = Read-SharedText $path
    if (-not $raw.Trim()) { return $res }
    $o = ConvertFrom-Json $raw
    if ($o -and $o.PSObject.Properties['items']) {
        $res.items = @($o.items | Where-Object { $_ })
        if ($o.PSObject.Properties['vec'] -and $o.vec) { foreach ($p in $o.vec.PSObject.Properties) { $res.vec[$p.Name] = [long]$p.Value } }
        if ($o.deleted) { foreach ($p in $o.deleted.PSObject.Properties) { $res.deleted[$p.Name] = [long]$p.Value } }
    } else {
        $res.items = @($o | Where-Object { $_ })   # 旧版：整个文件就是一个数组
    }
    return $res
}

function Mark-Store($st) {
    $now = [DateTime]::UtcNow.Ticks
    $seen = @{}
    $dirty = $false
    foreach ($x in $st.List) {
        $seen[$x.id] = $true
        $sig = Get-ItemSig $x
        if ($st.Base[$x.id] -ne $sig) { Mark-Fine $st $x $now; $x.upd = $now; $st.Base[$x.id] = $sig; $dirty = $true }
    }
    foreach ($k in @($st.Base.Keys)) {
        if (-not $seen.ContainsKey($k)) { $st.Dead[$k] = $now; $st.Base.Remove($k); $st.FBase.Remove($k); $st.SubBase.Remove($k); $dirty = $true }
    }
    return $dirty
}

# 把读到的文件内容合进内存；返回 changed（内存变了，要刷新界面）和 needWrite（本机有文件里没有的改动）
function Merge-Store($st, $r) {
    Note-Recv $st.FileName $r
    $changed = $false
    foreach ($k in @($r.deleted.Keys)) {
        $t = [long]$r.deleted[$k]
        if (-not $st.Dead.ContainsKey($k) -or [long]$st.Dead[$k] -lt $t) { $st.Dead[$k] = $t }
    }
    $idx = @{}
    for ($i = 0; $i -lt $st.List.Count; $i++) { $idx[$st.List[$i].id] = $i }
    $remoteUpd = @{}
    $diffRemote = $false
    foreach ($o in $r.items) {
        if (-not $o.text) { continue }
        $x = & $st.Conv $o
        Copy-Fine $o $x
        $remoteUpd[$x.id] = [long]$x.upd
        $dead = if ($st.Dead.ContainsKey($x.id)) { [long]$st.Dead[$x.id] } else { -1 }
        if ($dead -ge [long]$x.upd) { continue }
        if ($idx.ContainsKey($x.id)) {
            $cur = $st.List[$idx[$x.id]]
            $sx = Get-ItemSig $x; $sc = Get-ItemSig $cur
            if ($sx -ne $sc) {
                # 两边这一条不一样：按字段、按具体事项合并，谁后改用谁的
                Merge-Pair $cur $x
                $sm = Get-ItemSig $cur
                if ($sm -ne $sc) { $changed = $true }
                if ($sm -ne $sx) { $diffRemote = $true }
                $st.Base[$cur.id] = $sm
                Remember-Fine $st $cur
            } elseif ([long]$x.upd -gt [long]$cur.upd) {
                $cur.upd = [long]$x.upd
            }
        } else {
            [void]$st.List.Add($x)
            $idx[$x.id] = $st.List.Count - 1
            $st.Base[$x.id] = Get-ItemSig $x
            Remember-Fine $st $x
            $changed = $true
        }
    }
    for ($i = $st.List.Count - 1; $i -ge 0; $i--) {
        $x = $st.List[$i]
        if ($st.Dead.ContainsKey($x.id) -and [long]$st.Dead[$x.id] -ge [long]$x.upd) {
            $st.List.RemoveAt($i); $st.Base.Remove($x.id); $st.FBase.Remove($x.id); $st.SubBase.Remove($x.id); $changed = $true
        }
    }
    $needWrite = $diffRemote
    foreach ($x in $st.List) {
        if (-not $remoteUpd.ContainsKey($x.id) -or [long]$x.upd -gt $remoteUpd[$x.id]) { $needWrite = $true; break }
    }
    if (-not $needWrite) {
        foreach ($k in $st.Dead.Keys) {
            if (-not $r.deleted.ContainsKey($k) -or [long]$st.Dead[$k] -gt [long]$r.deleted[$k]) { $needWrite = $true; break }
        }
    }
    return @{ changed = $changed; needWrite = $needWrite }
}

function Prune-Dead($st) {
    $limit = [DateTime]::UtcNow.AddDays(-90).Ticks
    foreach ($k in @($st.Dead.Keys)) { if ([long]$st.Dead[$k] -lt $limit) { $st.Dead.Remove($k) } }
}

function Write-StoreFile($st) {
    $path = Get-StorePath $st
    Prune-Dead $st
    $at = [DateTime]::UtcNow.Ticks
    # vec：这份文件里包含了哪台电脑、截到什么时候的改动（用来判断"对方的改动收到没有"）
    $vec = @{}
    foreach ($k in @($script:Recv.Keys)) { $q = $k.Split('|'); if ($q.Count -eq 2 -and $q[1] -eq $st.FileName) { $vec[$q[0]] = [long]$script:Recv[$k] } }
    $vec[(Get-DeviceId)] = $at
    if (-not $st.Pruned -or ((Get-Date) - $st.Pruned).TotalHours -ge 6) { foreach ($x in $st.List) { Prune-SubDel $x }; $st.Pruned = Get-Date }
    $obj = [ordered]@{ version = 2; by = (Get-DeviceId); at = $at; vec = $vec; items = @($st.List.ToArray()); deleted = $st.Dead }
    $json = ConvertTo-JsonFast $obj
    $tmp = Join-Path $DataDir ($st.FileName + '.writing')
    [IO.File]::WriteAllText($tmp, $json, $Utf8)
    Invoke-Retry { [IO.File]::Copy($tmp, $path, $true) } 8 300
    if ($script:S.shareDir) {
        try {
            if (-not (Test-Path -LiteralPath $MirrorDir)) { New-Item -ItemType Directory -Force -Path $MirrorDir | Out-Null }
            [IO.File]::Copy($tmp, (Join-Path $MirrorDir $st.FileName), $true)
        } catch { Write-Log ('本机备份写入失败：' + $_.Exception.Message) }
    }
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    $st.Stamp = (Get-Item -LiteralPath $path).LastWriteTimeUtc
    $script:MyChange[$st.FileName] = $at
    $script:DevDirty = $true
}

function Report-SyncError([string]$msg) {
    Write-Log $msg
    # 偶尔一次失败不打扰，程序每 5 秒会自动重写；连续失败超过 2 分钟才提醒
    if (-not $script:FailSince) { $script:FailSince = Get-Date; return }
    if (-not $script:SyncWarned -and ((Get-Date) - $script:FailSince).TotalMinutes -ge 2) {
        $script:SyncWarned = $true
        Notify '数据一直存不进同步文件夹' ('已经连续 2 分钟写不进去，改动还在程序里。请看看网盘客户端是不是退出了或同步出错。详细原因记在：' + $LogFile)
    }
}

function Load-Store($st) {
    $loaded = Load-StoreMain $st
    # 同步文件夹里的文件可能被网盘用另一台电脑的旧版本覆盖了：把本机另存的那份合进来，按每条的修改时间取新的
    if ($script:S.shareDir) {
        $mp = Join-Path $MirrorDir $st.FileName
        if (Test-Path -LiteralPath $mp) {
            try {
                $m = Merge-Store $st (Read-StoreFile $mp)
                if ($m.changed -or -not $loaded) {
                    $script:PendingWrite[$st.FileName] = $true
                    Write-Log ('同步文件夹里的 ' + $st.FileName + ' 缺了本机的改动（多半是网盘用旧版本覆盖了），已从本机备份合并回来')
                }
                $loaded = $true
            } catch { Write-Log ('读取本机备份失败：' + $_.Exception.Message) }
        }
    }
    return $loaded
}

function Load-StoreMain($st) {
    $st.List.Clear(); $st.Base = @{}; $st.Dead = @{}; $st.Stamp = $null; $st.FBase = @{}; $st.SubBase = @{}
    $path = Get-StorePath $st
    if (-not (Test-Path -LiteralPath $path)) { return $false }
    try {
        $r = Read-StoreFile $path
        Note-Recv $st.FileName $r
        foreach ($o in $r.items) {
            if (-not $o.text) { continue }
            $x = & $st.Conv $o
            Copy-Fine $o $x
            [void]$st.List.Add($x)
            $st.Base[$x.id] = Get-ItemSig $x
            Remember-Fine $st $x
        }
        $st.Dead = $r.deleted
        $st.Stamp = (Get-Item -LiteralPath $path).LastWriteTimeUtc
    } catch {
        $bak = Join-Path $DataDir ($st.FileName + '.broken-' + (Get-Date).ToString('yyyyMMddHHmmss'))
        try { Copy-Item -LiteralPath $path $bak -Force } catch {}
        Write-Log ("读取 $path 失败，已备份到 $bak ：" + $_.Exception.Message)
    }
    return $true
}

$script:SaveQueue = @{}
$script:SaveTimer = [System.Windows.Threading.DispatcherTimer]::new([System.Windows.Threading.DispatcherPriority]::Background)
$script:SaveTimer.Interval = [TimeSpan]::FromMilliseconds(350)
$script:SaveTimer.add_Tick({ Flush-Saves })
function Request-Save($st) {
    $script:SaveQueue[$st.FileName] = $st
    $script:SaveTimer.Stop(); $script:SaveTimer.Start()
}
function Flush-Saves {
    $script:SaveTimer.Stop()
    if ($script:SaveQueue.Count -eq 0) { return }
    $q = @($script:SaveQueue.Values); $script:SaveQueue = @{}
    foreach ($st in $q) { Save-Store $st }
    if ($script:NeedRender -and -not (Test-Typing)) { $script:NeedRender = $false; Render }
}

function Save-Store($st, [bool]$force = $false) {
    # 没有真正改动就不写文件，省得网盘反复同步
    $dirty = (Mark-Store $st) -or $force -or $script:PendingWrite.ContainsKey($st.FileName)
    $path = Get-StorePath $st
    try {
        if (Test-Path -LiteralPath $path) {
            $lw = (Get-Item -LiteralPath $path).LastWriteTimeUtc
            if ($st.Stamp -ne $lw) {
                $m = Merge-Store $st (Read-StoreFile $path)
                $st.Stamp = $lw
                if ($m.changed) { $script:NeedRender = $true }
                if ($m.needWrite) { $dirty = $true }
            }
        } else { $dirty = $true }
        if (-not $dirty) { return }
        Write-StoreFile $st
        $script:PendingWrite.Remove($st.FileName)
        $script:SyncWarned = $false
        $script:FailSince = $null
    } catch {
        $script:PendingWrite[$st.FileName] = $true
        Report-SyncError ('保存失败：' + $_.Exception.Message)
    }
}

# 同步软件遇到两边同时改，会生成"todos (1).json""todos-冲突.json"之类的副本：合并进来后移到本机备份
function Merge-ConflictCopies($st) {
    if ($script:ShareDir -eq $DataDir) { return $false }
    $base = [IO.Path]::GetFileNameWithoutExtension($st.FileName)
    $any = $false
    foreach ($f in @(Get-ChildItem -LiteralPath $script:ShareDir -Filter ($base + '*.json') -File -ErrorAction SilentlyContinue)) {
        if ($f.Name -eq $st.FileName) { continue }
        try {
            $m = Merge-Store $st (Read-StoreFile $f.FullName)
            $bakDir = Join-Path $DataDir 'conflicts'
            if (-not (Test-Path $bakDir)) { New-Item -ItemType Directory -Force -Path $bakDir | Out-Null }
            Move-Item -LiteralPath $f.FullName -Destination (Join-Path $bakDir ((Get-Date).ToString('yyyyMMddHHmmss') + '_' + $f.Name)) -Force
            $any = $true
        } catch { Write-Log ('合并冲突副本失败：' + $f.FullName + ' ' + $_.Exception.Message) }
    }
    return $any
}

# 定时检查同步文件夹里的文件有没有被别的电脑改过
function Poll-Stores {
    Flush-Saves
    $changed = $false
    foreach ($st in $script:StoreList) {
        try {
            if ($script:PendingWrite.ContainsKey($st.FileName)) { Save-Store $st }
            $path = Get-StorePath $st
            $conf = Merge-ConflictCopies $st
            if ($conf) { $changed = $true; Save-Store $st; continue }
            if (-not (Test-Path -LiteralPath $path)) { continue }
            $lw = (Get-Item -LiteralPath $path).LastWriteTimeUtc
            if ($st.Stamp -eq $lw) { continue }
            [void](Mark-Store $st)
            $m = Merge-Store $st (Read-StoreFile $path)
            $st.Stamp = $lw
            if ($m.changed) { $changed = $true }
            if ($m.needWrite) { Write-StoreFile $st }
        } catch { Write-Log ('检查同步文件失败：' + $_.Exception.Message) }
    }
    if ($changed -or $script:NeedRender) {
        $script:NeedRender = $false
        if (-not (Test-Typing)) { Render } else { $script:NeedRender = $true }
    }
}

function Copy-MissingImages([string]$from, [string]$to) {
    $src = Join-Path $from 'images'
    if (-not (Test-Path -LiteralPath $src)) { return }
    $dst = Join-Path $to 'images'
    if (-not (Test-Path -LiteralPath $dst)) { New-Item -ItemType Directory -Force -Path $dst | Out-Null }
    foreach ($f in Get-ChildItem -LiteralPath $src -File) {
        $target = Join-Path $dst $f.Name
        if (-not (Test-Path -LiteralPath $target)) { Copy-Item -LiteralPath $f.FullName $target }
    }
}

function Set-ShareDir([string]$dir) {
    Flush-Saves
    $old = $script:ShareDir
    try { if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null } }
    catch { [System.Windows.MessageBox]::Show('这个文件夹建不了：' + $dir, '桌面待办') | Out-Null; return }
    $before = @{}
    foreach ($st in $script:StoreList) { $before[$st.FileName] = $st.List.Count }
    Copy-MissingImages $old $dir
    try { Copy-MissingReimbFiles $old $dir } catch { Write-Log ('复制报销附件失败：' + $_.Exception.Message) }
    $script:ShareDir = $dir
    foreach ($st in $script:StoreList) {
        if ($dir -eq $DataDir) {
            # 改回本机：直接用内存里的最新数据覆盖本机文件
            $p = Get-StorePath $st
            $st.Stamp = if (Test-Path -LiteralPath $p) { (Get-Item -LiteralPath $p).LastWriteTimeUtc } else { $null }
        } else {
            $st.Stamp = $null
        }
        Save-Store $st $true
    }
    $script:S.shareDir = if ($dir -eq $DataDir) { '' } else { $dir }
    if ($dir -eq $DataDir) { try { Remove-Item -LiteralPath $MirrorDir -Recurse -Force -ErrorAction SilentlyContinue } catch {} }
    else { foreach ($st in $script:StoreList) { $script:PendingWrite[$st.FileName] = $true } }
    Save-Settings
    Render-Chips
    Render
    $t = $script:StoreList[0]
    if ($dir -eq $DataDir) {
        Notify '已改回只存在本机' '数据已经写回本机。同步文件夹里的文件没有删，别的电脑还能继续用。'
    } else {
        Notify '同步已打开' ('待办现在 {0} 条（本机原来 {1} 条）。另一台电脑选同一个文件夹，两边就会自动合并。' -f $t.List.Count, $before[$t.FileName])
        if (-not $script:S.syncAppMode) { $script:AskSyncApp = $true }
    }
}

function Choose-ShareDir {
    $fb = [System.Windows.Forms.FolderBrowserDialog]::new()
    $fb.Description = '选一个会自动同步的文件夹（坚果云、百度网盘同步空间、OneDrive 等）。程序会在里面放 todos.json、worklog.json 和 images 文件夹。每台电脑都选同一个文件夹。'
    $fb.ShowNewFolderButton = $true
    if ($script:S.shareDir) { $fb.SelectedPath = [string]$script:S.shareDir }
    if ($fb.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $dir = $fb.SelectedPath
    if ($dir -eq $script:ShareDir) { return }
    if ([IO.Path]::GetFileName($dir) -ne '桌面待办数据' -and -not (Test-Path -LiteralPath (Join-Path $dir 'todos.json'))) {
        $sub = Join-Path $dir '桌面待办数据'
        $ans = [System.Windows.MessageBox]::Show("在这个文件夹里新建一个「桌面待办数据」子文件夹来放数据？`n`n$sub`n`n选「否」就直接放在你选的文件夹里。", '桌面待办', 'YesNoCancel', 'Question')
        if ($ans -eq 'Cancel') { return }
        if ($ans -eq 'Yes') { $dir = $sub }
    }
    Set-ShareDir $dir
}

function Load-Items {
    if (-not (Load-Store $script:TodoStore)) {
        [void]$script:Items.Add((New-Todo '示例：铅笔按钮修改，右键可以设提醒时间和重要，做完点左边打勾'))
        Save-Items
    }
}

function Save-Items { Request-Save $script:TodoStore }

function Find-Todo([string]$id) {
    foreach ($it in $script:Items) { if ($it.id -eq $id) { return $it } }
    return $null
}

# ---------- 设置 ----------
$script:S = @{
    left = $null; top = $null; width = 340; height = 540
    topmost = $true; collapsed = $false; showDone = $false
    asked = $false; trayTipShown = $false
    mode = 'todo'; logCols = '日期,时间,工作内容,项目'; projects = @()
    opacity = 1.0; hoverOpaque = $true; theme = '便签黄'; customColor = ''
    quick = $null; shareDir = ''; groupProj = $false; showTime = $true; lastProj = 'cat-daily'; lastLogProj = 'cat-daily'
}
function Load-Settings {
    if (-not (Test-Path $SetFile)) { return }
    try {
        $o = ConvertFrom-Json ([IO.File]::ReadAllText($SetFile, $Utf8))
        foreach ($p in $o.PSObject.Properties) { $script:S[$p.Name] = $p.Value }
    } catch { Write-Log ('读取设置失败：' + $_.Exception.Message) }
}
function Save-Settings {
    try { [IO.File]::WriteAllText($SetFile, (ConvertTo-Json $script:S -Depth 5), $Utf8) } catch {}
}

# ---------- 时间 ----------
function Format-When([datetime]$d) {
    $today = (Get-Date).Date
    $days = ($d.Date - $today).Days
    $hm = $d.ToString('HH:mm')
    if ($days -eq 0)  { return "今天 $hm" }
    if ($days -eq 1)  { return "明天 $hm" }
    if ($days -eq -1) { return "昨天 $hm" }
    if ($days -gt 1 -and $days -lt 7) { return ($WeekNames[[int]$d.DayOfWeek] + " $hm") }
    return ($d.ToString('M/d') + " $hm")
}

# 显示用的时间：关掉"显示几点几分"后只显示日期
function Format-Show([datetime]$d) {
    if ($script:S.showTime -ne $false) { return (Format-When $d) }
    $days = ($d.Date - (Get-Date).Date).Days
    if ($days -eq 0) { return '今天' }
    if ($days -eq 1) { return '明天' }
    if ($days -eq -1) { return '昨天' }
    if ($days -gt 1 -and $days -lt 7) { return $WeekNames[[int]$d.DayOfWeek] }
    return $d.ToString('M/d')
}
function Test-ShowTime { return ($script:S.showTime -ne $false) }

function Set-ShowTime([bool]$on) {
    $script:S.showTime = $on
    Save-Settings
    Render
}

function New-NumTitle([int]$no, $content) {
    if ($no -le 0) { return $content }
    $dp = [System.Windows.Controls.DockPanel]::new()
    $nl = New-NumLabel $no
    $nl.Margin = Th 0 1 4 0
    [System.Windows.Controls.DockPanel]::SetDock($nl, 'Left')
    [void]$dp.Children.Add($nl)
    [void]$dp.Children.Add($content)
    return $dp
}

$DefaultQuick = @(
    @{ name = '今天'; day = 'd0'; time = '18:00' },
    @{ name = '明天'; day = 'd1'; time = '18:00' },
    @{ name = '周五'; day = 'w5'; time = '18:00' },
    @{ name = '下周一'; day = 'n1'; time = '09:00' }
)
$DayChoices = [ordered]@{ 'd0' = '今天'; 'd1' = '明天'; 'd2' = '后天' }
$WeekShort = @('', '一', '二', '三', '四', '五', '六', '日')
for ($i = 1; $i -le 7; $i++) { $DayChoices['w' + $i] = '周' + $WeekShort[$i] }
for ($i = 1; $i -le 7; $i++) { $DayChoices['n' + $i] = '下周' + $WeekShort[$i] }

function Get-QuickTime($q) {
    $now = Get-Date
    $t = $now.Date
    $hm = ([string]$q.time).Split(':')
    $span = [TimeSpan]::new([int]$hm[0], [int]$hm[1], 0)
    $code = [string]$q.day
    $kind = $code.Substring(0, 1)
    $n = [int]$code.Substring(1)
    if ($kind -eq 'd') {
        $d = $t.AddDays($n) + $span
        if ($n -eq 0 -and $d -le $now) {
            # 今天的时间已经过了：改成今晚 23:30，太晚了就顺延到明天同一时间
            $d = $t.AddHours(23).AddMinutes(30)
            if ($d -le $now) { $d = $t.AddDays(1) + $span }
        }
        return $d
    }
    $target = $n % 7
    $dow = [int]$t.DayOfWeek
    if ($kind -eq 'w') {
        $diff = ($target - $dow + 7) % 7
        $d = $t.AddDays($diff) + $span
        if ($d -le $now) { $d = $d.AddDays(7) }
        return $d
    }
    $toMon = (8 - $dow) % 7
    if ($toMon -eq 0) { $toMon = 7 }
    return $t.AddDays($toMon + $n - 1) + $span
}

function Get-QuickName($q) {
    if ($q.name) { return [string]$q.name }
    return ($DayChoices[[string]$q.day] + ' ' + $q.time)
}

function Get-QuickList { return @(@($script:S.quick) | Where-Object { $_ -and $_.day -and $_.time }) }

# ---------- 界面 ----------
$app = [System.Windows.Application]::new()
$app.ShutdownMode = [System.Windows.ShutdownMode]::OnExplicitShutdown
$ThinScrollXaml = @'
<ResourceDictionary xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml">
  <Style x:Key="ThinThumb" TargetType="Thumb">
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="Thumb">
          <Border x:Name="tb" CornerRadius="3" Background="{DynamicResource MutedBrush}" Opacity="0.45"/>
          <ControlTemplate.Triggers>
            <Trigger Property="IsMouseOver" Value="True"><Setter TargetName="tb" Property="Opacity" Value="0.8"/></Trigger>
            <Trigger Property="IsDragging" Value="True"><Setter TargetName="tb" Property="Opacity" Value="0.9"/></Trigger>
          </ControlTemplate.Triggers>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
  </Style>
  <Style TargetType="ScrollBar">
    <Setter Property="Background" Value="Transparent"/>
    <Setter Property="Width" Value="8"/>
    <Setter Property="MinWidth" Value="8"/>
    <Setter Property="Template">
      <Setter.Value>
        <ControlTemplate TargetType="ScrollBar">
          <Grid Background="Transparent">
            <Track x:Name="PART_Track" Orientation="{TemplateBinding Orientation}" IsDirectionReversed="True">
              <Track.Thumb><Thumb Style="{StaticResource ThinThumb}" Margin="1,2"/></Track.Thumb>
            </Track>
          </Grid>
        </ControlTemplate>
      </Setter.Value>
    </Setter>
    <Style.Triggers>
      <Trigger Property="Orientation" Value="Horizontal">
        <Setter Property="Width" Value="Auto"/>
        <Setter Property="MinWidth" Value="0"/>
        <Setter Property="Height" Value="8"/>
        <Setter Property="MinHeight" Value="8"/>
        <Setter Property="Template">
          <Setter.Value>
            <ControlTemplate TargetType="ScrollBar">
              <Grid Background="Transparent">
                <Track x:Name="PART_Track" Orientation="{TemplateBinding Orientation}" IsDirectionReversed="False">
                  <Track.Thumb><Thumb Style="{StaticResource ThinThumb}" Margin="2,1"/></Track.Thumb>
                </Track>
              </Grid>
            </ControlTemplate>
          </Setter.Value>
        </Setter>
      </Trigger>
    </Style.Triggers>
  </Style>
</ResourceDictionary>
'@
try { $app.Resources.MergedDictionaries.Add([Windows.Markup.XamlReader]::Parse($ThinScrollXaml)) } catch { }

[xml]$xaml = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="桌面待办" Width="340" Height="540" MinWidth="280"
        WindowStyle="None" AllowsTransparency="True" Background="Transparent"
        ResizeMode="CanResizeWithGrip" ShowInTaskbar="False" Topmost="True"
        FontFamily="Microsoft YaHei UI" FontSize="13" Foreground="{DynamicResource TextBrush}">
  <Window.Resources>
    <SolidColorBrush x:Key="TextBrush" Color="#2B2B2B"/>
    <SolidColorBrush x:Key="SubBrush" Color="#8A7F68"/>
    <SolidColorBrush x:Key="MutedBrush" Color="#A39A87"/>
    <SolidColorBrush x:Key="AccentBrush" Color="#2F6FED"/>
    <SolidColorBrush x:Key="BgBrush" Color="#FFFBF2"/>
    <SolidColorBrush x:Key="HeaderBrush" Color="#F3EAD3"/>
    <SolidColorBrush x:Key="LineBrush" Color="#E3DAC6"/>
    <SolidColorBrush x:Key="HoverBrush" Color="#E8DEC6"/>
    <SolidColorBrush x:Key="InputLineBrush" Color="#D9CFB8"/>
    <SolidColorBrush x:Key="InputBgBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="HintBrush" Color="#B0A690"/>
    <SolidColorBrush x:Key="CardBrush" Color="#FFFFFF"/>
    <SolidColorBrush x:Key="CardLineBrush" Color="#ECE4D3"/>
    <SolidColorBrush x:Key="ChipBrush" Color="#EFE6D1"/>
    <SolidColorBrush x:Key="ChipTextBrush" Color="#6E6450"/>
    <SolidColorBrush x:Key="OverdueBrush" Color="#FFF2EF"/>
    <Style x:Key="IconBtn" TargetType="Button">
      <Setter Property="Width" Value="26"/>
      <Setter Property="Height" Value="24"/>
      <Setter Property="FontFamily" Value="Segoe MDL2 Assets"/>
      <Setter Property="FontSize" Value="11"/>
      <Setter Property="Foreground" Value="{DynamicResource SubBrush}"/>
      <Setter Property="Cursor" Value="Hand"/>
      <Setter Property="Focusable" Value="False"/>
      <Setter Property="Template">
        <Setter.Value>
          <ControlTemplate TargetType="Button">
            <Border x:Name="bd" Background="Transparent" CornerRadius="4">
              <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
            </Border>
            <ControlTemplate.Triggers>
              <Trigger Property="IsMouseOver" Value="True">
                <Setter TargetName="bd" Property="Background" Value="{DynamicResource HoverBrush}"/>
              </Trigger>
            </ControlTemplate.Triggers>
          </ControlTemplate>
        </Setter.Value>
      </Setter>
    </Style>
  </Window.Resources>

  <Border CornerRadius="10" Background="{DynamicResource BgBrush}" BorderBrush="{DynamicResource LineBrush}" BorderThickness="1" Margin="8">
    <Border.Effect>
      <DropShadowEffect BlurRadius="14" ShadowDepth="2" Opacity="0.22"/>
    </Border.Effect>
    <Grid>
      <Grid.RowDefinitions>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="Auto"/>
        <RowDefinition Height="*"/>
        <RowDefinition Height="Auto"/>
      </Grid.RowDefinitions>

      <Border x:Name="Header" Grid.Row="0" Background="{DynamicResource HeaderBrush}" CornerRadius="10,10,0,0" Padding="12,7,6,7">
        <DockPanel LastChildFill="True">
          <StackPanel DockPanel.Dock="Right" Orientation="Horizontal">
            <Button x:Name="BtnSync" Visibility="Collapsed" Style="{StaticResource IconBtn}" Content="&#xE895;"/>
            <Button x:Name="BtnTime" Visibility="Collapsed" Style="{StaticResource IconBtn}" Content="&#xE823;"/>
            <Button x:Name="BtnStyle" Visibility="Collapsed" Style="{StaticResource IconBtn}" Content="&#xE790;"/>
            <Button x:Name="BtnPin" Visibility="Collapsed" Style="{StaticResource IconBtn}"/>
            <Button x:Name="BtnCollapse" Visibility="Collapsed" Style="{StaticResource IconBtn}"/>
            <Button x:Name="BtnCal" Style="{StaticResource IconBtn}" Content="&#xE787;" ToolTip="工作台：周视图、月视图、报告、项目"/>
            <Button x:Name="BtnSearch" Style="{StaticResource IconBtn}" Content="&#xE721;" ToolTip="搜索（Ctrl+F）"/>
            <Button x:Name="BtnMore" Style="{StaticResource IconBtn}" Content="&#xE712;" ToolTip="更多：几点几分、置顶、外观、同步、更新"/>
            <Button x:Name="BtnHide" Style="{StaticResource IconBtn}" Content="&#xE711;" ToolTip="缩到托盘（Ctrl+Alt+N 呼出）"/>
          </StackPanel>
          <DockPanel VerticalAlignment="Center" LastChildFill="True">
            <TextBlock x:Name="TabTodo" DockPanel.Dock="Left" Text="待办" FontWeight="SemiBold" FontSize="14" Cursor="Hand"/>
            <TextBlock x:Name="TabLog" DockPanel.Dock="Left" Text="工作记录" FontSize="14" Margin="14,0,0,0" Cursor="Hand" Foreground="{DynamicResource MutedBrush}"/>
            <TextBlock x:Name="SyncDot" DockPanel.Dock="Left" Text="●" FontSize="10" Margin="9,1,0,0" VerticalAlignment="Center" Cursor="Hand" Visibility="Collapsed"/>
            <TextBlock x:Name="Count" Margin="10,1,4,0" FontSize="12" VerticalAlignment="Center" Foreground="{DynamicResource SubBrush}" TextTrimming="CharacterEllipsis"/>
          </DockPanel>
        </DockPanel>
      </Border>

      <StackPanel x:Name="InputArea" Grid.Row="1" Margin="10,10,10,6">
        <DockPanel x:Name="SearchBar" Margin="0,0,0,8" Visibility="Collapsed">
          <Button x:Name="BtnCloseSearch" DockPanel.Dock="Right" Style="{StaticResource IconBtn}" Content="&#xE711;" ToolTip="关闭搜索（Esc）"/>
          <Grid>
            <TextBox x:Name="SearchBox" Padding="6,4" BorderBrush="{DynamicResource AccentBrush}" Background="{DynamicResource InputBgBrush}"
                     Foreground="{DynamicResource TextBrush}" CaretBrush="{DynamicResource TextBrush}"/>
            <TextBlock x:Name="SearchHint" Text="搜索题目、具体事项、项目名；多个词用空格隔开" Foreground="{DynamicResource HintBrush}"
                       Margin="9,5,0,0" IsHitTestVisible="False" FontSize="12"/>
          </Grid>
        </DockPanel>
        <Grid>
          <TextBox x:Name="InputBox" Padding="6,5" BorderBrush="{DynamicResource InputLineBrush}" Background="{DynamicResource InputBgBrush}" Foreground="{DynamicResource TextBrush}" CaretBrush="{DynamicResource TextBrush}"
                   TextWrapping="Wrap" AcceptsReturn="False" MaxHeight="90"/>
          <TextBlock x:Name="Hint" Text="记一条，回车保存" Foreground="{DynamicResource HintBrush}"
                     Margin="9,6,0,0" IsHitTestVisible="False"/>
        </Grid>
        <TextBlock x:Name="ParseHint" FontSize="11" Margin="2,4,0,0" Foreground="{DynamicResource AccentBrush}" TextWrapping="Wrap" Visibility="Collapsed"/>
        <WrapPanel x:Name="Chips" Margin="0,7,0,0"/>
        <DockPanel x:Name="DayBar" Margin="0,2,0,0" Visibility="Collapsed">
          <TextBlock x:Name="BtnImport" DockPanel.Dock="Right" Foreground="{DynamicResource AccentBrush}" Cursor="Hand" FontSize="12"
                     VerticalAlignment="Center" ToolTip="把这一天打过勾的待办加进工作记录"/>
          <StackPanel Orientation="Horizontal">
            <Button x:Name="BtnPrevDay" Style="{StaticResource IconBtn}" Content="&#xE76B;" ToolTip="前一天"/>
            <TextBlock x:Name="DayLabel" VerticalAlignment="Center" FontWeight="SemiBold" Margin="2,0" Cursor="Hand" ToolTip="点击回到今天"/>
            <Button x:Name="BtnNextDay" Style="{StaticResource IconBtn}" Content="&#xE76C;" ToolTip="后一天"/>
          </StackPanel>
        </DockPanel>
        <Border x:Name="ProjBarBox" BorderThickness="0,1,0,0" BorderBrush="{DynamicResource LineBrush}" Margin="0,8,0,0" Padding="0,6,0,0" Visibility="Collapsed">
          <WrapPanel x:Name="ProjBar"/>
        </Border>
      </StackPanel>

      <ScrollViewer x:Name="ListArea" Grid.Row="2" VerticalScrollBarVisibility="Auto" Margin="10,2,4,0">
        <StackPanel x:Name="List" Margin="0,0,6,0"/>
      </ScrollViewer>

      <Border x:Name="Footer" Grid.Row="3" Padding="12,6,22,8">
        <Grid>
        <DockPanel x:Name="FootTodo">
          <TextBlock x:Name="BtnExport" DockPanel.Dock="Right" Text="复制周报" Foreground="{DynamicResource AccentBrush}" Cursor="Hand" FontSize="12"
                     ToolTip="把本周按项目整理成周报复制到剪贴板；月报、季报、导出 Word 在工作台的「报告」里"/>
          <TextBlock x:Name="BtnShowDone" DockPanel.Dock="Left" Foreground="{DynamicResource SubBrush}" Cursor="Hand" FontSize="12"/>
          <TextBlock x:Name="BtnGroup" DockPanel.Dock="Right" Margin="0,0,14,0" Text="按项目" Foreground="{DynamicResource SubBrush}" Cursor="Hand" FontSize="12"
                     ToolTip="按项目分组显示 / 按时间排序"/>
          <TextBlock x:Name="BtnExpandAll" Margin="14,0,0,0" Text="全部展开" Foreground="{DynamicResource AccentBrush}" Cursor="Hand" FontSize="12"
                     ToolTip="把列表里的待办全部展开 / 全部收起"/>
        </DockPanel>
        <DockPanel x:Name="FootLog" Visibility="Collapsed" LastChildFill="False">
          <TextBlock x:Name="BtnCols" DockPanel.Dock="Right" Text=" ▾" Foreground="{DynamicResource AccentBrush}" Cursor="Hand" FontSize="12" Margin="2,0,0,0"
                     ToolTip="表格列设置：按腾讯表格的列顺序设一次就行"/>
          <TextBlock x:Name="BtnCopyLog" DockPanel.Dock="Right" Text="复制到表格" Foreground="{DynamicResource AccentBrush}" Cursor="Hand" FontSize="12"
                     ToolTip="按表格列顺序复制这一天的记录（包括这天完成的待办），到腾讯表格里 Ctrl+V"/>
          <TextBlock x:Name="BtnLogExpand" DockPanel.Dock="Left" Text="全部展开" Foreground="{DynamicResource AccentBrush}" Cursor="Hand" FontSize="12" Visibility="Collapsed"
                     ToolTip="把这一天记录里待办的具体事项全部展开 / 收起"/>
        </DockPanel>
        <Border x:Name="UndoBar" Visibility="Collapsed" Background="{DynamicResource ChipBrush}" CornerRadius="4" Padding="8,2,8,3" Margin="-4,-3,-14,-2">
          <DockPanel>
            <TextBlock x:Name="UndoBtn" DockPanel.Dock="Right" Text="撤销" Foreground="{DynamicResource AccentBrush}" Cursor="Hand" FontSize="12" FontWeight="SemiBold" Margin="12,0,0,0"/>
            <TextBlock x:Name="UndoText" FontSize="12" Foreground="{DynamicResource TextBrush}" TextTrimming="CharacterEllipsis"/>
          </DockPanel>
        </Border>
        </Grid>
      </Border>
    </Grid>
  </Border>
</Window>
'@

$window = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $xaml))
foreach ($n in 'Header','Count','BtnPin','BtnCollapse','BtnHide','InputArea','InputBox','Hint','Chips','ListArea','List','Footer','BtnExport','BtnShowDone','TabTodo','TabLog','DayBar','BtnImport','BtnPrevDay','DayLabel','BtnNextDay','FootTodo','FootLog','BtnCopyLog','BtnCols','BtnStyle','BtnSearch','SearchBar','SearchBox','SearchHint','BtnCloseSearch','BtnExpandAll','BtnCal','BtnGroup','BtnLogExpand','BtnTime','ProjBar','BtnSync','BtnMore','SyncDot','ParseHint','ProjBarBox','UndoBar','UndoText','UndoBtn') {
    Set-Variable -Name $n -Value $window.FindName($n) -Scope Script
}
$IconBtnStyle = $window.FindResource('IconBtn')
$BC = [System.Windows.Media.BrushConverter]::new()
function Brush([string]$hex) { return $BC.ConvertFromString($hex) }
function Th([double]$l, [double]$t, [double]$r, [double]$b) { return [System.Windows.Thickness]::new($l, $t, $r, $b) }
$Hand = [System.Windows.Input.Cursors]::Hand
$Glyph = @{ Pin = [string][char]0xE718; Pinned = [string][char]0xE840; Up = [string][char]0xE70E; Down = [string][char]0xE70D; Del = [string][char]0xE74D; Copy = [string][char]0xE8C8; Done = [string][char]0xE73E; Close = [string][char]0xE711; Edit = [string][char]0xE70F }

# ---------- 托盘 ----------
if (-not ('DeskTodoNative' -as [type])) {
    Add-Type -TypeDefinition 'using System; using System.Runtime.InteropServices; public static class DeskTodoNative { [DllImport("user32.dll")] public static extern bool DestroyIcon(IntPtr h); }'
}
function New-RoundRectPath([single]$x, [single]$y, [single]$w, [single]$r) {
    $p = [System.Drawing.Drawing2D.GraphicsPath]::new()
    [single]$d = $r * 2
    [single]$x2 = $x + $w - $d; [single]$y2 = $y + $w - $d
    $p.AddArc($x, $y, $d, $d, [single]180, [single]90)
    $p.AddArc($x2, $y, $d, $d, [single]270, [single]90)
    $p.AddArc($x2, $y2, $d, $d, [single]0, [single]90)
    $p.AddArc($x, $y2, $d, $d, [single]90, [single]90)
    $p.CloseFigure()
    return $p
}
function Draw-AppIcon([int]$z, [bool]$alert = $false) {
    $bmp = [System.Drawing.Bitmap]::new($z, $z, [System.Drawing.Imaging.PixelFormat]::Format32bppArgb)
    $g = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality
    $g.Clear([System.Drawing.Color]::Transparent)
    $m = [single][math]::Max(0.5, $z * 0.06); $w = [single]($z - 2 * $m); $r = [single]($w * 0.26)
    $path = New-RoundRectPath $m $m $w $r
    $rect = [System.Drawing.RectangleF]::new($m, $m, $w, $w)
    $lg = [System.Drawing.Drawing2D.LinearGradientBrush]::new($rect, [System.Drawing.Color]::FromArgb(255, 92, 156, 255), [System.Drawing.Color]::FromArgb(255, 30, 94, 224), [single]90)
    $g.FillPath($lg, $path)
    if ($z -ge 32) {
        # 上半部分一点点高光，大图标才看得出来
        $hl = [System.Drawing.Drawing2D.LinearGradientBrush]::new($rect, [System.Drawing.Color]::FromArgb(46, 255, 255, 255), [System.Drawing.Color]::FromArgb(0, 255, 255, 255), [single]90)
        $g.FillPath($hl, $path); $hl.Dispose()
    }
    $pen = [System.Drawing.Pen]::new([System.Drawing.Color]::White, [single][math]::Max(1.7, $z * 0.115))
    $pen.StartCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.EndCap = [System.Drawing.Drawing2D.LineCap]::Round
    $pen.LineJoin = [System.Drawing.Drawing2D.LineJoin]::Round
    $pts = [System.Drawing.PointF[]]@(
        [System.Drawing.PointF]::new([single]($m + $w * 0.25), [single]($m + $w * 0.53)),
        [System.Drawing.PointF]::new([single]($m + $w * 0.43), [single]($m + $w * 0.71)),
        [System.Drawing.PointF]::new([single]($m + $w * 0.76), [single]($m + $w * 0.33)))
    $g.DrawLines($pen, $pts)
    if ($alert) {
        $dd = [single]($z * 0.44); $x0 = [single]($z - $dd)
        $g.FillEllipse([System.Drawing.Brushes]::White, [single]($x0 - 0.5), [single]-0.5, [single]($dd + 1), [single]($dd + 1))
        $red = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 239, 68, 68))
        $in = [single][math]::Max(1, $z * 0.06)
        $g.FillEllipse($red, [single]($x0 + $in), $in, [single]($dd - 2 * $in), [single]($dd - 2 * $in)); $red.Dispose()
    }
    $pen.Dispose(); $lg.Dispose(); $path.Dispose(); $g.Dispose()
    return $bmp
}
# 做一个多尺寸的 .ico（给桌面图标、开始菜单、窗口用）
function Write-AppIco([string]$path) {
    $sizes = @(16, 20, 24, 32, 40, 48, 64, 128, 256)
    $pngs = @(foreach ($z in $sizes) {
        $b = Draw-AppIcon $z $false
        $ms = [IO.MemoryStream]::new(); $b.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png); $b.Dispose()
        , $ms.ToArray()
    })
    $out = [IO.MemoryStream]::new(); $bw = [IO.BinaryWriter]::new($out)
    $bw.Write([uint16]0); $bw.Write([uint16]1); $bw.Write([uint16]$sizes.Count)
    $off = 6 + 16 * $sizes.Count
    for ($i = 0; $i -lt $sizes.Count; $i++) {
        $z = $sizes[$i]; $bz = if ($z -ge 256) { 0 } else { $z }
        $bw.Write([byte]$bz); $bw.Write([byte]$bz); $bw.Write([byte]0); $bw.Write([byte]0)
        $bw.Write([uint16]1); $bw.Write([uint16]32)
        $bw.Write([uint32]$pngs[$i].Length); $bw.Write([uint32]$off)
        $off += $pngs[$i].Length
    }
    foreach ($p in $pngs) { $bw.Write([byte[]]$p) }
    $bw.Flush()
    [IO.File]::WriteAllBytes($path, $out.ToArray())
}
$IcoVersion = 1
function Ensure-AppIco {
    foreach ($dir in @($AppDir, $DataDir)) {
        $p = Join-Path $dir 'DeskTodo.ico'
        try {
            if (-not (Test-Path -LiteralPath $p) -or [int]$script:S.icoVer -ne $IcoVersion) { Write-AppIco $p }
            $script:IcoPath = $p
            return
        } catch { Write-Log ('图标文件写不进 ' + $dir + '：' + $_.Exception.Message) }
    }
}
function Set-WindowIcon($w) {
    if (-not $script:IcoPath) { return }
    try { $w.Icon = [System.Windows.Media.Imaging.BitmapFrame]::Create([Uri]::new($script:IcoPath)) } catch {}
}
function New-TrayIconImage { $b = Draw-AppIcon 32 $false; return [System.Drawing.Icon]::FromHandle($b.GetHicon()) }
# 托盘图标按屏幕缩放画，清楚不糊；有超期的待办时带红点
$script:TrayAlert = $null
function Update-TrayIcon([bool]$alert) {
    if ($script:TrayAlert -eq $alert) { return }
    try {
        $sz = [System.Windows.Forms.SystemInformation]::SmallIconSize.Width
        if ($sz -lt 16) { $sz = 16 }
        $b = Draw-AppIcon $sz $alert
        $h = $b.GetHicon(); $b.Dispose()
        $old = $script:TrayHicon
        $script:Tray.Icon = [System.Drawing.Icon]::FromHandle($h)
        $script:TrayHicon = $h; $script:TrayAlert = $alert
        if ($old) { [void][DeskTodoNative]::DestroyIcon($old) }
    } catch { Write-Log ('托盘图标出错：' + $_.Exception.Message) }
}

$script:Tray = [System.Windows.Forms.NotifyIcon]::new()
Update-TrayIcon $false
$script:Tray.Text = '桌面待办（Ctrl+Alt+N）'
$script:Tray.Visible = $true

function Notify([string]$title, [string]$text) {
    if ($text.Length -gt 200) { $text = $text.Substring(0, 200) + '…' }
    $script:Tray.ShowBalloonTip(8000, $title, $text, [System.Windows.Forms.ToolTipIcon]::Info)
}

# ---------- 窗口显示 ----------
function Show-Main {
    $window.Show()
    if ($window.WindowState -eq 'Minimized') { $window.WindowState = 'Normal' }
    $window.Topmost = $true
    [void]$window.Activate()
    $window.Topmost = [bool]$script:S.topmost
    if ($script:S.collapsed) { Set-Collapsed $false }
    [void]$InputBox.Focus()
}

function Hide-Main {
    $window.Hide()
    Save-Settings
    if (-not $script:S.trayTipShown) {
        $script:S.trayTipShown = $true
        Save-Settings
        Notify '桌面待办缩到这里了' '按 Ctrl+Alt+N 或点右下角蓝色勾图标随时呼出。'
    }
}

function Toggle-Main {
    if ($window.IsVisible -and $window.IsActive) { Hide-Main } else { Show-Main }
}

function Get-ExpandedHeight {
    $h = [double]$script:S.height
    if ($h -lt 200) { $h = 540 }
    return $h
}

function Set-Collapsed([bool]$c) {
    $vis = if ($c) { [System.Windows.Visibility]::Collapsed } else { [System.Windows.Visibility]::Visible }
    if ($c) {
        # 收起前先记住当前展开时的大小
        if (-not $script:S.collapsed -and $window.ActualHeight -ge 200) {
            $script:S.height = $window.ActualHeight
            $script:S.width = $window.ActualWidth
        }
        $script:S.collapsed = $true
        $InputArea.Visibility = $vis; $ListArea.Visibility = $vis; $Footer.Visibility = $vis
        $window.MinHeight = 0
        $window.ResizeMode = [System.Windows.ResizeMode]::NoResize
        $window.SizeToContent = [System.Windows.SizeToContent]::Height
        $BtnCollapse.Content = $Glyph.Down
    } else {
        # 先固定高度再显示内容，避免窗口按内容自动变矮
        $window.SizeToContent = [System.Windows.SizeToContent]::Manual
        $window.Height = Get-ExpandedHeight
        $window.MinHeight = 180
        $window.ResizeMode = [System.Windows.ResizeMode]::CanResizeWithGrip
        $InputArea.Visibility = $vis; $ListArea.Visibility = $vis; $Footer.Visibility = $vis
        $script:S.collapsed = $false
        $BtnCollapse.Content = $Glyph.Up
    }
    Save-Settings
}

function Set-Pin([bool]$on) {
    $script:S.topmost = $on
    $window.Topmost = $on
    if ($on) { $BtnPin.Content = $Glyph.Pinned; $BtnPin.ToolTip = '已置顶，点击取消' }
    else     { $BtnPin.Content = $Glyph.Pin;    $BtnPin.ToolTip = '点击置顶' }
    Save-Settings
}

# ---------- 对话框 ----------
function New-DialogWindow([string]$title) {
    $d = [System.Windows.Window]::new()
    $d.Title = $title
    if ($window.IsVisible) { $d.Owner = $window; $d.WindowStartupLocation = 'CenterOwner' } else { $d.WindowStartupLocation = 'CenterScreen' }
    $d.SizeToContent = 'WidthAndHeight'
    $d.ResizeMode = 'NoResize'
    $d.Topmost = $true
    $d.ShowInTaskbar = $false
    $d.FontFamily = [System.Windows.Media.FontFamily]::new('Microsoft YaHei UI')
    $d.FontSize = 13
    Set-WindowIcon $d
    return $d
}

function New-DialogButtons {
    $bp = [System.Windows.Controls.StackPanel]::new()
    $bp.Orientation = 'Horizontal'
    $bp.HorizontalAlignment = 'Right'
    $bp.Margin = Th 0 12 0 0
    $ok = [System.Windows.Controls.Button]::new()
    $ok.Content = '确定'; $ok.Width = 72; $ok.Height = 28; $ok.IsDefault = $true; $ok.Margin = Th 0 0 8 0
    $cancel = [System.Windows.Controls.Button]::new()
    $cancel.Content = '取消'; $cancel.Width = 72; $cancel.Height = 28; $cancel.IsCancel = $true
    [void]$bp.Children.Add($ok); [void]$bp.Children.Add($cancel)
    return @($bp, $ok)
}

function Show-TextDialog([string]$title, [string]$init, [string]$note = '', [bool]$multi = $false) {
    $script:Dlg = New-DialogWindow $title
    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.Margin = Th 16 14 16 14
    if ($note) {
        $nt = [System.Windows.Controls.TextBlock]::new()
        $nt.Text = $note; $nt.TextWrapping = 'Wrap'; $nt.Width = 340
        $nt.Foreground = Brush '#6E6450'; $nt.Margin = Th 0 0 0 8
        [void]$sp.Children.Add($nt)
    }
    $script:DlgTb = [System.Windows.Controls.TextBox]::new()
    $script:DlgTb.Width = 340; $script:DlgTb.MinHeight = 64
    $script:DlgTb.TextWrapping = 'Wrap'; $script:DlgTb.Padding = Th 5 4 5 4
    $script:DlgTb.Text = $init
    if ($multi) {
        $script:DlgTb.Width = 460; $script:DlgTb.MinHeight = 160; $script:DlgTb.MaxHeight = 460
        $script:DlgTb.AcceptsReturn = $true; $script:DlgTb.VerticalScrollBarVisibility = 'Auto'
    }
    [void]$sp.Children.Add($script:DlgTb)
    $btns = New-DialogButtons
    [void]$sp.Children.Add($btns[0])
    $script:DlgResult = $null
    $btns[1].add_Click({ $script:DlgResult = $script:DlgTb.Text; $script:Dlg.DialogResult = $true })
    $script:Dlg.Content = $sp
    $script:Dlg.add_ContentRendered({ [void]$script:DlgTb.Focus(); $script:DlgTb.SelectAll() })
    [void]$script:Dlg.ShowDialog()
    return $script:DlgResult
}

function New-TimeCombo([int]$max, [int]$value) {
    $cb = [System.Windows.Controls.ComboBox]::new()
    $cb.Width = 64
    $cb.IsEditable = $true
    $cb.MaxDropDownHeight = 260
    for ($i = 0; $i -le $max; $i++) { [void]$cb.Items.Add(('{0:00}' -f $i)) }
    $cb.Text = '{0:00}' -f $value
    return $cb
}

$OnQuickTime = {
    param($s, $e)
    $tag = [string]$s.Tag
    if ($tag -eq '现在') { $tag = (Get-Date).ToString('HH:mm') }
    $hm = $tag.Split(':')
    $script:DlgHour.Text = $hm[0]
    $script:DlgMin.Text = $hm[1]
}

function Show-DatePickerDialog($init, [string]$title = '选截止时间') {
    if (-not $init) { $init = (Get-Date).Date.AddDays(1).AddHours(18) }
    $script:Dlg = New-DialogWindow $title
    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.Margin = Th 16 12 16 14
    $script:DlgCal = [System.Windows.Controls.Calendar]::new()
    $script:DlgCal.SelectedDate = $init.Date
    $script:DlgCal.DisplayDate = $init.Date
    [void]$sp.Children.Add($script:DlgCal)

    # 时、分可以下拉选，也可以直接输入
    $row = [System.Windows.Controls.StackPanel]::new()
    $row.Orientation = 'Horizontal'; $row.Margin = Th 0 8 0 0
    $lab = [System.Windows.Controls.TextBlock]::new()
    $lab.Text = '时间'; $lab.VerticalAlignment = 'Center'; $lab.Margin = Th 0 0 8 0
    $script:DlgHour = New-TimeCombo 23 $init.Hour
    $colon = [System.Windows.Controls.TextBlock]::new()
    $colon.Text = ':'; $colon.VerticalAlignment = 'Center'; $colon.Margin = Th 5 0 5 0
    $script:DlgMin = New-TimeCombo 59 $init.Minute
    [void]$row.Children.Add($lab); [void]$row.Children.Add($script:DlgHour)
    [void]$row.Children.Add($colon); [void]$row.Children.Add($script:DlgMin)
    [void]$sp.Children.Add($row)

    $quick = [System.Windows.Controls.WrapPanel]::new()
    $quick.Margin = Th 0 8 0 0
    foreach ($q in '09:00', '12:00', '14:00', '17:30', '18:00') {
        $t = [System.Windows.Controls.TextBlock]::new()
        $t.Text = $q; $t.Tag = $q
        $t.Foreground = Brush '#2F6FED'; $t.Cursor = $Hand; $t.Margin = Th 0 0 12 0
        $t.add_MouseLeftButtonUp($OnQuickTime)
        [void]$quick.Children.Add($t)
    }
    [void]$sp.Children.Add($quick)

    $btns = New-DialogButtons
    [void]$sp.Children.Add($btns[0])
    $script:DlgResult = $null
    $btns[1].add_Click({
        $day = $script:DlgCal.SelectedDate
        $h = 0; $mi = 0
        $okH = [int]::TryParse(([string]$script:DlgHour.Text).Trim(), [ref]$h) -and $h -ge 0 -and $h -le 23
        $okM = [int]::TryParse(([string]$script:DlgMin.Text).Trim(), [ref]$mi) -and $mi -ge 0 -and $mi -le 59
        if (-not $day) {
            [System.Windows.MessageBox]::Show($script:Dlg, '先在日历上点一下日期。', '桌面待办') | Out-Null
            return
        }
        if (-not ($okH -and $okM)) {
            [System.Windows.MessageBox]::Show($script:Dlg, '时间填得不对：小时填 0～23，分钟填 0～59。', '桌面待办') | Out-Null
            return
        }
        $script:DlgResult = ([datetime]$day).Date.AddHours($h).AddMinutes($mi)
        $script:Dlg.DialogResult = $true
    })
    $script:Dlg.Content = $sp
    [void]$script:Dlg.ShowDialog()
    return $script:DlgResult
}

# ---------- 操作 ----------
# ---------- 从一句话里认出提醒时间和项目 ----------
# 例：「明天下午3点 交材料 #年会」→ 提醒 明天 15:00，项目 年会筹办，题目「交材料」
$CnDigitMap = @{ '零' = 0; '〇' = 0; '一' = 1; '二' = 2; '两' = 2; '三' = 3; '四' = 4; '五' = 5; '六' = 6; '七' = 7; '八' = 8; '九' = 9; '十' = 10 }
function ConvertFrom-CnNum([string]$s) {
    if ($s -match '^\d+$') { return [int]$s }
    if ($s.Length -eq 1) { if ($CnDigitMap.ContainsKey($s)) { return [int]$CnDigitMap[$s] }; return -1 }
    # 十一、二十、二十三
    $i = $s.IndexOf('十')
    if ($i -lt 0) { return -1 }
    $a = if ($i -eq 0) { 1 } else { $CnDigitMap[$s.Substring(0, $i)] }
    $b = if ($i -eq $s.Length - 1) { 0 } else { $CnDigitMap[$s.Substring($i + 1)] }
    if ($null -eq $a -or $null -eq $b) { return -1 }
    return ($a * 10 + $b)
}
$WeekCn = @{ '一' = 1; '二' = 2; '三' = 3; '四' = 4; '五' = 5; '六' = 6; '日' = 7; '天' = 7; '1' = 1; '2' = 2; '3' = 3; '4' = 4; '5' = 5; '6' = 6; '7' = 7 }

function Find-ProjectByWord([string]$w) {
    $w = $w.Trim()
    if (-not $w) { return $null }
    $all = @(Get-ProjectsOrdered $false)
    $hit = @($all | Where-Object { $_.text -eq $w })
    if ($hit.Count -eq 0) { $hit = @($all | Where-Object { $_.text.StartsWith($w) }) }
    if ($hit.Count -eq 0) { $hit = @($all | Where-Object { $_.text.Contains($w) }) }
    if ($hit.Count -ge 1) { return $hit[0] }
    return $null
}

function Parse-TodoText([string]$text, [datetime]$now) {
    $r = @{ text = $text; when = $null; proj = $null; projWord = ''; notes = @(); past = '' }
    $s = ' ' + $text + ' '

    # 项目：#年会 或 ＃年会
    $m = [regex]::Match($s, '[#＃]([^\s#＃，,。；;]+)')
    if ($m.Success) {
        $r.projWord = $m.Groups[1].Value
        $p = Find-ProjectByWord $r.projWord
        if ($p) { $r.proj = $p; $s = $s.Remove($m.Index, $m.Length).Insert($m.Index, ' ') }
    }

    # 日期
    $date = $null; $defH = 18; $defM = 0; $dayPart = ''
    $num = '(\d{1,2}|[一二三四五六七八九十两]{1,3})'
    $pats = @(
        @{ re = '大后天'; f = { $now.Date.AddDays(3) } },
        @{ re = '后天'; f = { $now.Date.AddDays(2) } },
        @{ re = '明早|明天早上|明天上午'; f = { $script:_dp = '上午'; $now.Date.AddDays(1) } },
        @{ re = '明晚|明天晚上'; f = { $script:_dp = '晚上'; $now.Date.AddDays(1) } },
        @{ re = '明天|明日'; f = { $now.Date.AddDays(1) } },
        @{ re = '今晚|今天晚上'; f = { $script:_dp = '晚上'; $now.Date } },
        @{ re = '今天|今日'; f = { $now.Date } },
        @{ re = '(下下|下|这|本)?(?:周|星期|礼拜)([一二三四五六日天1-7])'; f = {
            $wd = $WeekCn[$mm.Groups[2].Value]; $pre = $mm.Groups[1].Value
            $dow = [int]$now.DayOfWeek; if ($dow -eq 0) { $dow = 7 }
            $mon = $now.Date.AddDays(1 - $dow)
            if ($pre -eq '下') { $mon.AddDays(7 + $wd - 1) }
            elseif ($pre -eq '下下') { $mon.AddDays(14 + $wd - 1) }
            elseif ($pre) { $mon.AddDays($wd - 1) }
            else { $d0 = $mon.AddDays($wd - 1); if ($d0 -lt $now.Date) { $d0 = $d0.AddDays(7) }; $d0 } } },
        @{ re = '(\d{1,2})月(\d{1,2})[日号]?'; f = {
            $mo = [int]$mm.Groups[1].Value; $da = [int]$mm.Groups[2].Value
            if ($mo -lt 1 -or $mo -gt 12 -or $da -lt 1 -or $da -gt [DateTime]::DaysInMonth($now.Year, $mo)) { $null }
            else { $d0 = [datetime]::new($now.Year, $mo, $da); if ($d0 -lt $now.Date.AddDays(-30)) { $d0 = $d0.AddYears(1) }; $d0 } } },
        @{ re = '(?<![\d/])(\d{1,2})/(\d{1,2})(?![\d/])'; f = {
            $mo = [int]$mm.Groups[1].Value; $da = [int]$mm.Groups[2].Value
            if ($mo -lt 1 -or $mo -gt 12 -or $da -lt 1 -or $da -gt [DateTime]::DaysInMonth($now.Year, $mo)) { $null }
            else { $d0 = [datetime]::new($now.Year, $mo, $da); if ($d0 -lt $now.Date.AddDays(-30)) { $d0 = $d0.AddYears(1) }; $d0 } } },
        @{ re = '(?<![\d月])(\d{1,2})[号日](?=\s|上午|下午|晚上|中午|早上|\d|[一二三四五六七八九十两]+点|$)'; f = {
            $da = [int]$mm.Groups[1].Value
            if ($da -lt 1 -or $da -gt 31) { $null }
            else {
                $d0 = $null
                for ($k = 0; $k -lt 3 -and -not $d0; $k++) {
                    $b = [datetime]::new($now.Year, $now.Month, 1).AddMonths($k)
                    if ($da -le [DateTime]::DaysInMonth($b.Year, $b.Month)) { $c = $b.AddDays($da - 1); if ($c -ge $now.Date) { $d0 = $c } }
                }
                $d0 } } },
        @{ re = '(\d{1,2}|[一二三四五六七八九十两]{1,3})天(?:后|以后|之后)'; f = { $n = ConvertFrom-CnNum $mm.Groups[1].Value; if ($n -ge 0) { $now.Date.AddDays($n) } else { $null } } }
    )
    $script:_dp = ''
    $past = [regex]::Match($s, '昨天|昨日|昨晚|前天|大前天|上周[一二三四五六日天]|上个?星期[一二三四五六日天]')
    if ($past.Success) { $r.past = $past.Value; return $r }
    foreach ($p in $pats) {
        $mm = [regex]::Match($s, $p.re)
        if (-not $mm.Success) { continue }
        $d = & $p.f
        if ($d) { $date = $d; $dayPart = $script:_dp; $s = $s.Remove($mm.Index, $mm.Length).Insert($mm.Index, ' '); break }
    }

    # 时间：下午3点、3点半、15:30、晚上八点一刻
    $hour = -1; $min = 0
    $tm = [regex]::Match($s, '(上午|早上|早晨|中午|下午|傍晚|晚上)?\s*(\d{1,2})[:：](\d{2})')
    if (-not $tm.Success) { $tm = [regex]::Match($s, '(上午|早上|早晨|中午|下午|傍晚|晚上)?\s*' + $num + '[点时](半|一刻|三刻|(\d{1,2})分?|' + '([一二三四五六七八九十]{1,3})分?)?') }
    if ($tm.Success) {
        $h = ConvertFrom-CnNum $tm.Groups[2].Value
        $mi = 0
        if ($tm.Groups[3].Success -and $tm.Groups[3].Value) {
            $g = $tm.Groups[3].Value
            if ($g -eq '半') { $mi = 30 } elseif ($g -eq '一刻') { $mi = 15 } elseif ($g -eq '三刻') { $mi = 45 }
            else { $mi = ConvertFrom-CnNum ($g.TrimEnd('分')) }
        }
        if ($h -ge 0 -and $h -le 24 -and $mi -ge 0 -and $mi -lt 60) {
            $part = $tm.Groups[1].Value
            if (-not $part) { $part = $dayPart }
            if ($part -match '下午|傍晚|晚上' -and $h -lt 12) { $h += 12 }
            elseif ($part -eq '中午' -and $h -lt 5) { $h += 12 }
            elseif (-not $part -and $h -ge 1 -and $h -le 6) { $h += 12 }   # 只说「3点」一般是下午
            if ($h -eq 24) { $h = 23; $mi = 59 }
            $hour = $h; $min = $mi
            $s = $s.Remove($tm.Index, $tm.Length).Insert($tm.Index, ' ')
        }
    }
    if ($hour -lt 0 -and $dayPart) { if ($dayPart -eq '上午') { $hour = 9 } else { $hour = 20 } }

    if ($date -or $hour -ge 0) {
        if (-not $date) {
            $date = $now.Date
            if ($date.AddHours($hour).AddMinutes($min) -le $now) { $date = $date.AddDays(1) }
        }
        if ($hour -lt 0) { $hour = $defH; $min = $defM }
        $r.when = $date.AddHours($hour).AddMinutes($min)
    }

    $t = ($s -replace '[\s　]+', ' ').Trim()
    $t = $t.Trim(' ', '，', ',', '、', '：', ':', '；', ';')
    $t = ($t -replace '^(要|得|需要)?在\s*', '')
    if ($r.when -or $r.proj) { if ($t) { $r.text = $t } }
    return $r
}

function Update-ParseHint {
    try { Update-ParseHintCore } catch { $ParseHint.Visibility = 'Collapsed'; Write-Log ('认时间出错：' + $_.Exception.Message) }
}
function Update-ParseHintCore {
    if ($script:S.mode -eq 'log' -or -not $InputBox.Text.Trim()) { $ParseHint.Visibility = 'Collapsed'; return }
    $pr = Parse-TodoText $InputBox.Text (Get-Date)
    if ($pr.past) {
        $ParseHint.Text = '「' + $pr.past + '」已经过去了，不设提醒' + $(if ($pr.proj) { '；项目「' + $pr.proj.text + '」' } else { '' })
        $ParseHint.Visibility = 'Visible'; return
    }
    $parts = @()
    if ($pr.when) { $parts += ('提醒 ' + (Format-When $pr.when)) }
    if ($pr.proj) { $parts += ('项目「' + $pr.proj.text + '」') }
    $miss = if ($pr.projWord -and -not $pr.proj) { '没有叫「' + $pr.projWord + '」的项目，这一段当普通文字' } else { '' }
    if ($parts.Count -eq 0 -and -not $miss) { $ParseHint.Visibility = 'Collapsed'; return }
    $ParseHint.Text = $(if ($parts.Count -gt 0) { '认出：' + ($parts -join ' · ') + '　记成「' + $pr.text + '」' } else { '' }) + $(if ($miss) { $(if ($parts.Count -gt 0) { '；' } else { '' }) + $miss } else { '' })
    $ParseHint.Visibility = 'Visible'
}

function Add-Todo {
    $text = $InputBox.Text.Trim()
    if (-not $text) { return }
    $pr = $null
    try { $pr = Parse-TodoText $text (Get-Date) } catch { Write-Log ('认时间出错：' + $_.Exception.Message) }
    if (-not $pr) { $pr = @{ text = $text; when = $null; proj = $null } }
    $text = $pr.text
    $t = New-Todo $text
    $t.project = if ($pr.proj) { $pr.proj.id } elseif ($script:Pending.proj) { $script:Pending.proj } else { 'cat-daily' }
    if ($t.project -eq 'cat-reimb') { $sa = Split-Amount $text; $t.text = $sa[0]; $t.amount = $sa[1]; $t.rstate = '待报' }
    $t.times = @(@($script:Pending.sel | ForEach-Object { ([datetime]$_.dt).ToString($Fmt) }) + @($(if ($pr.when) { $pr.when.ToString($Fmt) })) | Where-Object { $_ } | Sort-Object -Unique)
    Sync-Times $t
    $t.star = [bool]$script:Pending.star
    [void]$script:Items.Insert(0, $t)
    Save-Items
    Reset-Pending
    $InputBox.Clear()
    $ParseHint.Visibility = 'Collapsed'
    Render-Chips
    Render
    Check-Reminders
}

function Set-Done([string]$id, [bool]$done) {
    $it = Find-Todo $id
    if (-not $it) { return }
    $it.done = $done
    $it.doneAt = if ($done) { (Get-Date).ToString($Fmt) } else { '' }
    Save-Items
    Render
}

function Edit-Todo([string]$id) {
    $it = Find-Todo $id
    if (-not $it) { return }
    $r = Show-TextDialog '修改' $it.text
    if ($r -and $r.Trim()) { $it.text = $r.Trim(); Save-Items; Render }
}

# 删了不弹窗：底部出一条"已删除 · 撤销"，8 秒内能救回来；图片等撤销条消失后才真正删
$script:Undo = $null
$script:UndoTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:UndoTimer.Interval = [TimeSpan]::FromSeconds(8)
$script:UndoTimer.add_Tick({ Finish-Undo })
function Show-Undo([string]$text, $data) {
    Finish-Undo
    $script:Undo = $data
    $UndoText.Text = $text
    $UndoText.ToolTip = $text
    $UndoBar.Visibility = 'Visible'
    $script:UndoTimer.Stop(); $script:UndoTimer.Start()
}
function Finish-Undo {
    $script:UndoTimer.Stop()
    $u = $script:Undo; $script:Undo = $null
    $UndoBar.Visibility = 'Collapsed'
    if ($u -and $u.files) { try { Remove-NoteFiles $u.files } catch {} }
}
function Invoke-Undo {
    $script:UndoTimer.Stop()
    $u = $script:Undo; $script:Undo = $null
    $UndoBar.Visibility = 'Collapsed'
    if (-not $u) { return }
    switch ($u.kind) {
        'todo' { $i = [math]::Min([int]$u.index, $script:Items.Count); $script:Items.Insert($i, $u.item); Save-Items }
        'work' { [void]$script:Work.Add($u.item); Save-Work }
        'sub' {
            $t = Find-Todo $u.todo
            if ($t) {
                $l = [System.Collections.ArrayList]::new(); foreach ($y in @($t.subs)) { [void]$l.Add($y) }
                $l.Insert([math]::Min([int]$u.index, $l.Count), $u.item)
                $t.subs = @($l.ToArray()); Save-Items
            }
        }
    }
    Render
}
function Short-Text([string]$t, [int]$n = 16) { if ($t.Length -gt $n) { return $t.Substring(0, $n) + '…' }; return $t }

function Remove-Todo([string]$id) {
    $it = Find-Todo $id
    if (-not $it) { return }
    $idx = $script:Items.IndexOf($it)
    [void]$script:Items.Remove($it)
    Save-Items
    Render
    Show-Undo ('已删除「' + (Short-Text $it.text) + '」') @{ kind = 'todo'; item = $it; index = $idx; files = @(@($it.notes) + @($it.subs)) }
}

function Toggle-Star([string]$id) {
    $it = Find-Todo $id
    if (-not $it) { return }
    $it.star = -not $it.star
    Save-Items
    Render
}

function Set-Due([string]$id, [string]$op, [string]$arg) {
    $it = Find-Todo $id
    if (-not $it) { return }
    switch ($op) {
        'clear' { $it.times = @() }
        'del'   { $it.times = @(@($it.times) | Where-Object { $_ -ne $arg }) }
        'edit'  {
            $old = Parse-Dt $arg
            $d = Show-DatePickerDialog $old '修改这个提醒 / 截止时间'
            if (-not $d) { return }
            $it.times = @(@($it.times) | Where-Object { $_ -ne $arg }) + $d.ToString($Fmt)
        }
        'pick'  {
            $d = Show-DatePickerDialog (Get-NextTime $it (Get-Date))
            if (-not $d) { return }
            $it.times = @($it.times) + $d.ToString($Fmt)
        }
        'quick' {
            $q = (Get-QuickList)[[int]$arg]
            if (-not $q) { return }
            $it.times = @($it.times) + (Get-QuickTime $q).ToString($Fmt)
        }
    }
    Sync-Times $it
    Save-Items
    Render
    Check-Reminders
}


# ---------- 提醒 ----------
function Check-Reminders {
    $now = Get-Date
    $dueNow = New-Object System.Collections.Generic.List[string]
    $remind = New-Object System.Collections.Generic.List[string]
    $soon = New-Object System.Collections.Generic.List[string]
    $changed = $false
    foreach ($it in $script:Items) {
        if ($it.done) { continue }
        foreach ($ts in @($it.times)) {
            $d = Parse-Dt $ts
            if (-not $d -or $now -lt $d -or (@($it.fired) -contains $ts)) { continue }
            $it.fired = @($it.fired) + $ts
            $changed = $true
            if ($ts -eq $it.due) { $dueNow.Add($it.text); $it.warned = $true }
            else { $remind.Add(('{0}（{1}）' -f $it.text, $d.ToString('HH:mm'))) }
        }
        $dd = Parse-Dt $it.due
        if ($dd -and -not $it.warned -and $now -ge $dd.AddMinutes(-30) -and $now -lt $dd) {
            $it.warned = $true
            $changed = $true
            $soon.Add(('{0}（{1} 截止）' -f $it.text, $dd.ToString('HH:mm')))
        }
    }
    if ($dueNow.Count -gt 0) {
        if ($dueNow.Count -eq 1) { Notify '到截止时间了' $dueNow[0] }
        else { Notify ('{0} 项到截止时间了' -f $dueNow.Count) ($dueNow -join '；') }
        Show-Main
    } elseif ($remind.Count -gt 0) {
        Notify '提醒' ($remind -join '；')
        Show-Main
    } elseif ($soon.Count -gt 0) {
        Notify '30 分钟内到期' ($soon -join '；')
    }
    if ($changed) { Save-Items }
}

# ---------- 渲染 ----------
function Get-Active {
    $now = Get-Date
    $script:Items | Where-Object { -not $_.done } | Sort-Object -Property `
        @{ Expression = { -not $_.star } },
        @{ Expression = { $n = Get-NextTime $_ $now; if ($n) { $n } else { $d = Parse-Dt $_.due; if ($d) { $d } else { [datetime]::MaxValue } } } },
        @{ Expression = { $_.created }; Descending = $true }
}

function Get-Sig {
    $now = Get-Date
    $parts = foreach ($it in $script:Items) {
        if (-not $it.done) {
            $d = Parse-Dt $it.due
            $n = Get-NextTime $it $now
            $st = if (-not $d) { 0 } elseif ($d -lt $now) { 2 } elseif ($n -and $n.Date -eq $now.Date) { 1 } else { 3 }
            $ns = if ($n) { $n.ToString($Fmt) } else { '' }
            '{0}:{1}:{2}' -f $it.id, $st, $ns
        }
    }
    return $now.ToString('yyyyMMdd') + '|' + ($parts -join ',')
}

$OnCheck   = { param($s, $e) Set-Done ([string]$s.Tag) ([bool]$s.IsChecked) }
$OnDelete  = { param($s, $e) Remove-Todo ([string]$s.Tag) }
$OnEdit    = { param($s, $e) Edit-Todo ([string]$s.Tag) }
$OnStar    = { param($s, $e) Toggle-Star ([string]$s.Tag) }
$OnDue     = { param($s, $e) $p = ([string]$s.Tag).Split([char[]]'|', 3); Set-Due $p[0] $p[1] $p[2] }
$OnRestore = { param($s, $e) Set-Done ([string]$s.Tag) $false }
$OnCardDown = { param($s, $e) if ($e.ClickCount -eq 2) { $e.Handled = $true; Edit-Todo ([string]$s.Tag) } }

function New-MenuItem([string]$header, [string]$tag, [scriptblock]$handler) {
    $m = [System.Windows.Controls.MenuItem]::new()
    $m.Header = $header
    $m.Tag = $tag
    if ($handler) { $m.add_Click($handler) }
    return $m
}

function New-Card($it, [datetime]$now, [int]$no = 0) {
    $card = [System.Windows.Controls.Border]::new()
    $card.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $card.Padding = Th 8 7 4 7
    $card.Margin = Th 0 0 0 6
    $card.Background = TB 'Card'
    $card.BorderBrush = TB 'CardLine'
    $card.BorderThickness = Th 1 1 1 1
    $card.Tag = $it.id

    $due = Parse-Dt $it.due
    $next = Get-NextTime $it $now
    $overdue = (-not $it.done) -and $due -and ($due -lt $now)
    $dueToday = (-not $it.done) -and $next -and ($next.Date -eq $now.Date)
    if ($it.star -and -not $it.done) { $card.BorderBrush = Brush '#E25C4B'; $card.BorderThickness = Th 3 1 1 1 }
    if ($overdue) { $card.Background = TB 'Overdue' }

    $g = [System.Windows.Controls.Grid]::new()
    foreach ($w in 'auto', 'star', 'auto') {
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        if ($w -eq 'star') { $cd.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }
        else { $cd.Width = [System.Windows.GridLength]::Auto }
        [void]$g.ColumnDefinitions.Add($cd)
    }

    $cb = [System.Windows.Controls.CheckBox]::new()
    $cb.IsChecked = [bool]$it.done
    $cb.Tag = $it.id
    $cb.Margin = Th 0 3 8 0
    $cb.VerticalAlignment = 'Top'
    $cb.Cursor = $Hand
    $cb.ToolTip = if ($it.done) { '取消完成' } else { '做完了' }
    $cb.add_Click($OnCheck)
    [System.Windows.Controls.Grid]::SetColumn($cb, 0)
    [void]$g.Children.Add($cb)

    $sp = [System.Windows.Controls.StackPanel]::new()
    $tb = New-SelectableText $it.text 13.5
    [void](Enable-Inline $tb ('todo|' + $it.id))
    if ($it.done) { Set-DoneLook $tb }
    $ttl = New-NumTitle $no $tb
    $ttl.Margin = Th 0 0 108 0; $ttl.MinHeight = 24
    [void]$sp.Children.Add($ttl)

    $metaRow = [System.Windows.Controls.WrapPanel]::new()
    $metaRow.Margin = Th 0 4 0 0
    if ($it.done -or -not $script:HideProjTag) { [void]$metaRow.Children.Add((New-ProjTag $it)) }
    $meta = [System.Windows.Controls.TextBlock]::new()
    $meta.FontSize = 11
    $meta.VerticalAlignment = 'Center'
    $meta.Foreground = TB 'Muted'
    $meta.TextWrapping = 'Wrap'
    $meta.ToolTip = '双击修改日期和时间'
    $meta.add_MouseLeftButtonDown($OnMetaDbl)
    [void]$metaRow.Children.Add($meta)
    if ($it.done) {
        $da = Parse-Dt $it.doneAt
        $meta.Text = if ($da) { '完成于 ' + (Format-Show $da) } else { '已完成' }
    } else {
        $parts = @()
        if (Test-Reimb $it) { $mny = Format-Money $it.amount; $parts += ($(if ($mny) { '¥' + $mny } else { '未填金额' }) + ' · ' + $(if ($it.rstate) { $it.rstate } else { '待报' })) }
        $cr = Parse-Dt $it.created
        if ($cr) { $meta.ToolTip = '记于 ' + (Format-When $cr) + '；双击修改日期和时间' }
        $meta.TextWrapping = 'Wrap'
        $cnt = @($it.times).Count
        if ($overdue) {
            $parts += ('已超期 · 截止 ' + (Format-Show $due))
        } elseif ($cnt -eq 1) {
            if ($dueToday) { $parts += ('今天' + $(if (Test-ShowTime) { ' ' + $due.ToString('HH:mm') + ' ' } else { '' }) + '截止') } else { $parts += ('截止 ' + (Format-Show $due)) }
        } elseif ($cnt -gt 1) {
            $up = @(@($it.times) | ForEach-Object { Parse-Dt $_ } | Where-Object { $_ -ge $now } | ForEach-Object { Format-Show $_ } | Select-Object -Unique)
            if ($up.Count -le 1) { $parts += ('截止 ' + (Format-Show $due)) }
            else { $parts += ('提醒 ' + (@($up | Select-Object -First ($up.Count - 1)) -join '、') + '，截止 ' + $up[-1]) }
        }
        if ($overdue) { $meta.Foreground = Brush '#D0402F' }
        elseif ($dueToday) { $meta.Foreground = Brush '#D9822B' }
        $meta.Text = $parts -join '   '
    }
    [void]$sp.Children.Add($metaRow)

    $isOpen = Test-IsOpen $it.id
    $sum = Get-DetailSummary $it
    $lineRow = [System.Windows.Controls.WrapPanel]::new()
    $lineRow.Margin = Th 0 3 0 0
    [void]$sp.Children.Add($lineRow)
    if ($sum -or $isOpen) {
        $st = [System.Windows.Controls.TextBlock]::new()
        $arrow = if ($isOpen) { '▾ ' } else { '▸ ' }
        $st.Text = $arrow + $(if ($sum) { $sum } else { '具体事项' })
        $st.FontSize = 11
        $st.Margin = Th 0 3 0 0
        $st.Cursor = $Hand
        $st.Tag = $it.id
        $st.ToolTip = if ($isOpen) { '收起' } else { '展开看具体事项' }
        $st.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
        $st.add_MouseLeftButtonDown($OnToggleOpen)
        $st.Margin = Th 0 0 14 0
        [void]$lineRow.Children.Add($st)
    }
    # 进展：可以展开看每天记了什么，也可以收起来不看
    $progList = @(Get-TodoProgress $it)
    $popen = $script:ProgOpen.ContainsKey($it.id)
    if ($progList.Count -gt 0) {
        $pt = [System.Windows.Controls.TextBlock]::new()
        $pt.Text = $(if ($popen) { '▾ ' } else { '▸ ' }) + ('进展 {0} 条' -f $progList.Count)
        $pt.FontSize = 11; $pt.Cursor = $Hand; $pt.Tag = $it.id; $pt.Margin = Th 0 0 14 0
        $pt.ToolTip = if ($popen) { '收起进展' } else { '展开看每天的进展' }
        $pt.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
        $pt.add_MouseLeftButtonDown($OnProgToggle)
        [void]$lineRow.Children.Add($pt)
    }
    if (-not $it.done) {
        $ap = [System.Windows.Controls.TextBlock]::new()
        $ap.Text = '＋进展'
        $ap.FontSize = 11; $ap.Cursor = $Hand; $ap.Tag = $it.id; $ap.Margin = Th 0 0 14 0
        $ap.ToolTip = '这件事今天做了什么，记一笔；会进到今天的工作记录、周月视图和报告里'
        $ap.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
        $ap.add_MouseLeftButtonDown($OnAddProgress)
        [void]$lineRow.Children.Add($ap)
    }
    if ($progList.Count -gt 0) {
        if ($popen) {
            $pp = [System.Windows.Controls.StackPanel]::new()
            $pp.Margin = Th 0 4 0 0
            $k = 0
            foreach ($w in $progList) {
                $k++
                $dd = Parse-Day $w.date
                $when = if ($dd) { $dd.ToString('M月d日') } else { $w.date }
                if ((Test-ShowTime) -and $w.time) { $when += ' ' + $w.time }
                $prow = [System.Windows.Controls.DockPanel]::new()
                $prow.Margin = Th 4 1 0 0
                $pdt = [System.Windows.Controls.TextBlock]::new()
                $pdt.Text = ('{0}. {1}' -f $k, $when); $pdt.FontSize = 12; $pdt.Margin = Th 0 0 8 0; $pdt.Tag = $w.id
                $pdt.ToolTip = '双击修改这条进展的日期和时间'
                $pdt.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
                $pdt.add_MouseLeftButtonDown($OnWorkDateDbl)
                Set-LazyMenu $pdt { param($id) New-WorkLineMenu $id } $w.id
                [System.Windows.Controls.DockPanel]::SetDock($pdt, 'Left')
                [void]$prow.Children.Add($pdt)
                $pl = New-SelectableText (Get-WorkBody $w) 12
                $pl.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'SubBrush')
                $pl.ToolTip = '右键可以修改内容、日期和时间'
                [void](Enable-Inline $pl ('work|' + $w.id))
                Set-LazyMenu $pl { param($id) New-WorkLineMenu $id } $w.id $pl
                [void]$prow.Children.Add($pl)
                [void]$pp.Children.Add($prow)
            }
            [void]$sp.Children.Add($pp)
        }
    }
    [System.Windows.Controls.Grid]::SetColumn($sp, 1)
    [System.Windows.Controls.Grid]::SetColumnSpan($sp, 2)
    [void]$g.Children.Add($sp)

    $btnPanel = New-CardButtons $it.id $OnCopyTodo $OnDelete
    $ex = [System.Windows.Controls.Button]::new()
    $ex.Style = $IconBtnStyle
    $ex.Content = if ($isOpen) { $Glyph.Up } else { $Glyph.Down }
    $ex.Tag = $it.id
    $ex.ToolTip = if ($isOpen) { '收起' } else { '展开：写具体事项' }
    $ex.add_Click($OnToggleOpen)
    $btnPanel.Children.Insert(0, $ex)
    $btnPanel.Children.Insert(2, (New-EditBtn $it.id $OnEditBtnTodo 0))
    if ($isOpen) { $btnPanel.Opacity = 0.85 }
    [void]$g.Children.Add($btnPanel)
    if ($isOpen) {
        # 展开内容单独占一行，横跨整张卡片，不被右边的按钮挤窄
        foreach ($h in 'auto', 'auto') { [void]$g.RowDefinitions.Add(([System.Windows.Controls.RowDefinition]::new())) }
        $det = New-Details $it
        $det.Margin = Th 2 8 4 2
        [System.Windows.Controls.Grid]::SetRow($det, 1)
        [System.Windows.Controls.Grid]::SetColumn($det, 0)
        [System.Windows.Controls.Grid]::SetColumnSpan($det, 3)
        [void]$g.Children.Add($det)
    }
    $card.add_MouseEnter($OnCardEnter)
    $card.add_MouseLeave($OnCardLeave)

    $card.Child = $g
    $card.add_MouseLeftButtonDown($OnCardDown)
    Set-LazyMenu $card { param($id) $t0 = Find-Todo $id; if ($t0) { New-TodoMenu $t0 } } $it.id $tb
    return $card
}

function New-TodoMenu($it) {
    $menu = [System.Windows.Controls.ContextMenu]::new()
    if ($it.done) {
        [void]$menu.Items.Add((New-MenuItem '复制内容' $it.id $OnCopyTodo))
        [void]$menu.Items.Add((New-MenuItem '恢复为未完成' $it.id $OnRestore))
        [void]$menu.Items.Add((New-MenuItem '修改完成时间…' $it.id $OnEditDoneAt))
        [void]$menu.Items.Add((New-MenuItem '修改记录时间…' $it.id $OnEditCreated))
        [void]$menu.Items.Add((New-MenuItem '修改内容' $it.id $OnEdit))
        Add-ProjSubmenu $menu $it.id $it.project $OnTodoSetProj
    } else {
        [void]$menu.Items.Add((New-MenuItem '复制内容' $it.id $OnCopyTodo))
        [void]$menu.Items.Add((New-MenuItem '修改内容' $it.id $OnEdit))
        $openText = if (Test-IsOpen $it.id) { '收起具体事项' } else { '展开，写具体事项' }
        [void]$menu.Items.Add((New-MenuItem $openText $it.id { param($s, $e) Toggle-Open ([string]$s.Tag) }))
        [void]$menu.Items.Add((New-MenuItem '记进展…' $it.id $OnAddProgress))
        Add-ProjSubmenu $menu $it.id $it.project $OnTodoSetProj
        $smenu = New-MenuItem $(if ($it.start) { '开始日期：' + (Parse-Day $it.start).ToString('M月d日') } else { '开始日期' }) '' $null
        [void]$smenu.Items.Add((New-MenuItem '设置开始日期…' ($it.id + '|set') $OnTodoStart))
        if ($it.start) { [void]$smenu.Items.Add((New-MenuItem '清除开始日期' ($it.id + '|clear') $OnTodoStart)) }
        [void]$menu.Items.Add($smenu)
        if (Test-Reimb $it) {
            [void]$menu.Items.Add((New-MenuItem '报销金额…' $it.id $OnReimbAmount))
            $rm = New-MenuItem ('报销状态：' + $(if ($it.rstate) { $it.rstate } else { '待报' })) '' $null
            foreach ($rs in $ReimbStates) { [void]$rm.Items.Add((New-MenuItem $rs ($it.id + '|' + $rs) $OnReimbState)) }
            [void]$menu.Items.Add($rm)
        }
        $starText = if ($it.star) { '取消重要' } else { '标为重要（排到最上面）' }
        [void]$menu.Items.Add((New-MenuItem $starText $it.id $OnStar))
        $dueMenu = New-MenuItem '提醒 / 截止时间' '' $null
        $ql = Get-QuickList
        for ($qi = 0; $qi -lt $ql.Count; $qi++) {
            $label = '添加 ' + (Get-QuickName $ql[$qi]) + '（' + (Format-When (Get-QuickTime $ql[$qi])) + '）'
            [void]$dueMenu.Items.Add((New-MenuItem $label ($it.id + '|quick|' + $qi) $OnDue))
        }
        [void]$dueMenu.Items.Add((New-MenuItem '添加 选日期…' ($it.id + '|pick|') $OnDue))
        if (@($it.times).Count -gt 0) {
            [void]$dueMenu.Items.Add(([System.Windows.Controls.Separator]::new()))
            foreach ($ts in @($it.times)) {
                $dt = Parse-Dt $ts
                $tail = if ($ts -eq $it.due) { '（截止）' } else { '' }
                [void]$dueMenu.Items.Add((New-MenuItem ('修改 ' + (Format-When $dt) + $tail + '…') ($it.id + '|' + $ts) $OnEditDue))
                [void]$dueMenu.Items.Add((New-MenuItem ('删除 ' + (Format-When $dt) + $tail) ($it.id + '|del|' + $ts) $OnDue))
            }
            [void]$dueMenu.Items.Add((New-MenuItem '清除全部时间' ($it.id + '|clear|') $OnDue))
        }
        [void]$menu.Items.Add($dueMenu)
        [void]$menu.Items.Add((New-MenuItem '修改记录时间…' $it.id $OnEditCreated))
    }
    [void]$menu.Items.Add(([System.Windows.Controls.Separator]::new()))
    [void]$menu.Items.Add((New-MenuItem '删除' $it.id $OnDelete))
    return $menu
}

function New-Label([string]$text, $top) {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $text
    $t.FontSize = 12
    $t.Foreground = TB 'Muted'
    $t.Margin = Th 2 $top 0 6
    return $t
}

# 小窗上方的项目筛选：点一个项目就只看这个项目的待办
function Get-TodoProj($it) { if ($it.project) { return [string]$it.project }; return 'cat-daily' }
function Get-ProjFilter { $f = [string]$script:S.projFilter; if ($f -and (Find-Project $f)) { return $f }; return '' }
$OnProjFilter = {
    param($s, $e)
    $e.Handled = $true
    $f = [string]$s.Tag
    $script:S.projFilter = $f
    if ($f) { $script:Pending.proj = $f }
    Save-Settings
    Render-Chips
    Render
}
function New-FilterTab([string]$text, [string]$tag, [bool]$sel) {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $text; $t.FontSize = 12; $t.Margin = Th 0 0 12 3; $t.Cursor = $Hand; $t.Tag = $tag
    if ($sel) {
        $t.FontWeight = [System.Windows.FontWeights]::SemiBold
        $t.TextDecorations = [System.Windows.TextDecorations]::Underline
        $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
    } else { $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush') }
    $t.ToolTip = if ($tag) { '只看这个项目的待办；新记的待办也默认归到这个项目' } else { '看全部项目的待办' }
    $t.add_MouseLeftButtonUp($OnProjFilter)
    return $t
}
function Render-ProjBar($activeAll, $doneAll) {
    $ProjBar.Children.Clear()
    $pf = Get-ProjFilter
    $cnt = @{}
    foreach ($it in $activeAll) { $k = Get-TodoProj $it; $cnt[$k] = 1 + [int]$cnt[$k] }
    $has = @{}
    if ($script:S.showDone) { foreach ($it in $doneAll) { $has[(Get-TodoProj $it)] = $true } }
    $list = @(Get-ProjectsOrdered $true | Where-Object { $cnt.ContainsKey($_.id) -or $has.ContainsKey($_.id) -or $_.id -eq $pf })
    if ($list.Count -lt 2 -and -not $pf) { $ProjBarBox.Visibility = 'Collapsed'; return }
    $ProjBarBox.Visibility = 'Visible'
    $lb0 = [System.Windows.Controls.TextBlock]::new(); $lb0.Text = '只看：'; $lb0.FontSize = 12; $lb0.Margin = Th 0 0 4 3
    $lb0.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
    [void]$ProjBar.Children.Add($lb0)
    [void]$ProjBar.Children.Add((New-FilterTab ('全部 ' + $activeAll.Count) '' (-not $pf)))
    foreach ($p in $list) {
        $label = $p.text + $(if ($cnt.ContainsKey($p.id)) { ' ' + $cnt[$p.id] } else { '' })
        [void]$ProjBar.Children.Add((New-FilterTab $label $p.id ($p.id -eq $pf)))
    }
}

function Render-Todo {
    $List.Children.Clear()
    $now = Get-Date
    $words = Get-Words
    $searching = $words.Count -gt 0
    $activeAll = @(Get-Active)
    $doneAll0 = @($script:Items | Where-Object { $_.done })
    Render-ProjBar $activeAll $doneAll0
    $pf = if ($searching) { '' } else { Get-ProjFilter }
    $active = @($activeAll | Where-Object { (Test-TodoMatch $_ $words) -and (-not $pf -or (Get-TodoProj $_) -eq $pf) })
    $doneAll = @($doneAll0 | Where-Object { -not $pf -or (Get-TodoProj $_) -eq $pf })
    $visible = New-Object System.Collections.Generic.List[string]
    # 按项目分组、或者只看某个项目时，没做完的卡片上不再重复显示项目标签（右键里能换项目）
    $script:HideProjTag = [bool]($script:S.groupProj -or $pf)

    if ($active.Count -eq 0) {
        $empty = New-Label $(if ($searching) { '待办里没搜到' } elseif ($pf) { '「' + (Get-ProjName $pf) + '」里没有没做完的待办' } else { '暂无待办' }) 18
        $empty.HorizontalAlignment = 'Center'
        [void]$List.Children.Add($empty)
    }
    if ($script:S.groupProj -and $active.Count -gt 0) {
        $ord = @{}; $k = 0; foreach ($pp in (Get-ProjectsOrdered $true)) { $ord[$pp.id] = $k; $k++ }
        $grouped = @($active | Sort-Object @{ Expression = { if ($ord.ContainsKey($_.project)) { $ord[$_.project] } else { 9999 } } })
        $lastP = $null; $no = 0
        foreach ($it in $grouped) {
            if ($it.project -ne $lastP) {
                $lastP = $it.project; $no = 0
                $n = @($grouped | Where-Object { $_.project -eq $lastP }).Count
                $gl = New-Label ((Get-ProjName $lastP) + '（' + $n + '）') 8
                $gpj = Find-Project $lastP
                if ($gpj -and $gpj.kind -eq 'project') { $gl.Tag = $lastP; $gl.ToolTip = '双击修改项目名称'; $gl.add_MouseLeftButtonDown($OnProjLabelDbl) }
                [void]$List.Children.Add($gl)
            }
            $no++
            [void]$List.Children.Add((New-Card $it $now $no)); $visible.Add($it.id)
        }
    } else {
        $no = 0
        foreach ($it in $active) { $no++; [void]$List.Children.Add((New-Card $it $now $no)); $visible.Add($it.id) }
    }

    $doneShow = @()
    if ($searching) {
        $doneShow = @($doneAll | Where-Object { Test-TodoMatch $_ $words } | Sort-Object { $_.doneAt } -Descending)
        if ($doneShow.Count -gt 0) { [void]$List.Children.Add((New-Label ('已完成里搜到 {0} 条' -f $doneShow.Count) 10)) }
    } elseif ($script:S.showDone -and $doneAll.Count -gt 0) {
        $doneShow = @($doneAll | Sort-Object { $_.doneAt } -Descending)
    }
    if ($searching) {
        $no = 0
        foreach ($it in $doneShow) { $no++; $c = New-Card $it $now $no; $c.Opacity = 0.8; [void]$List.Children.Add($c); $visible.Add($it.id) }
    } elseif ($doneShow.Count -gt 0) {
        # 已完成按 今天 / 昨天 / 本周 / 更早 分段，"更早"默认收起
        $today = $now.Date; $wk = Get-WeekStart $today
        $segs = [ordered]@{ '今天完成' = @(); '昨天完成' = @(); '本周早些时候' = @(); '更早' = @() }
        foreach ($it in $doneShow) {
            $d = Parse-Dt $it.doneAt
            $k = if (-not $d) { '更早' } elseif ($d.Date -eq $today) { '今天完成' } elseif ($d.Date -eq $today.AddDays(-1)) { '昨天完成' } elseif ($d.Date -ge $wk) { '本周早些时候' } else { '更早' }
            $segs[$k] += $it
        }
        foreach ($k in @($segs.Keys)) {
            $arr = @($segs[$k]); if ($arr.Count -eq 0) { continue }
            $older = $k -eq '更早'
            $lbl = New-Label ($k + '（' + $arr.Count + '）' + $(if ($older) { $(if ($script:ShowOlderDone) { '  ▾ 收起' } else { '  ▸ 展开' }) } else { '' })) 10
            if ($older) {
                $lbl.Cursor = $Hand
                $lbl.add_MouseLeftButtonUp({ $script:ShowOlderDone = -not $script:ShowOlderDone; Render })
                [void]$List.Children.Add($lbl)
                if (-not $script:ShowOlderDone) { continue }
                $arr = @($arr | Select-Object -First 100)
            } else { [void]$List.Children.Add($lbl) }
            $no = 0
            foreach ($it in $arr) { $no++; $c = New-Card $it $now $no; $c.Opacity = 0.8; [void]$List.Children.Add($c); $visible.Add($it.id) }
        }
    }
    $script:VisibleIds = $visible.ToArray()

    $od = @($activeAll | Where-Object { $d = Parse-Dt $_.due; $d -and $d -lt $now }).Count
    $txt = '{0} 项' -f $activeAll.Count
    if ($od -gt 0) { $txt += ' · {0} 项超期' -f $od }
    $script:Tray.Text = '桌面待办：' + $txt
    Update-TrayIcon ($od -gt 0)
    $Count.ToolTip = ('还有 {0} 项没做完' -f $activeAll.Count) + $(if ($od -gt 0) { '，其中 {0} 项已经过了截止时间（卡片底色发红的那几条）' -f $od } else { '' })
    if ($searching) { $txt = '搜到 {0} 项' -f ($active.Count + $doneShow.Count) }
    $Count.Text = $txt
    $Count.Foreground = if ($od -gt 0 -and -not $searching) { Brush '#D0402F' } else { TB 'Sub' }

    $BtnShowDone.Text = if ($script:S.showDone) { '隐藏已完成' } else { '显示已完成（{0}）' -f $doneAll.Count }
    $BtnGroup.Text = if ($script:S.groupProj) { '按时间' } else { '按项目' }
    Update-ExpandLink
    $script:RenderSig = Get-Sig
}

# ---------- 输入框下的快捷标签 ----------
function Reset-Pending {
    $keep = if ($script:Pending -and $script:Pending.proj) { $script:Pending.proj } else { '' }
    $script:Pending = @{ sel = New-Object System.Collections.ArrayList; star = $false; proj = $keep }
}
Reset-Pending

function Find-PendingKey([string]$k) {
    foreach ($x in $script:Pending.sel) { if ($x.key -eq $k) { return $x } }
    return $null
}

$OnChip = {
    param($s, $e)
    $k = [string]$s.Tag
    if ($k -eq '重要') {
        $script:Pending.star = -not $script:Pending.star
    } elseif ($k -eq '选日期') {
        $d = Show-DatePickerDialog $null
        if ($d -and -not (Find-PendingKey ('c:' + $d.Ticks))) { [void]$script:Pending.sel.Add(@{ key = 'c:' + $d.Ticks; dt = $d }) }
    } elseif ($k -eq '__addquick') {
        Edit-Quick -1
    } elseif ($k.StartsWith('c:')) {
        $x = Find-PendingKey $k
        if ($x) { $script:Pending.sel.Remove($x) }
    } elseif ($k.StartsWith('q:')) {
        $x = Find-PendingKey $k
        if ($x) { $script:Pending.sel.Remove($x) }
        else {
            $q = (Get-QuickList)[[int]$k.Substring(2)]
            if ($q) { [void]$script:Pending.sel.Add(@{ key = $k; dt = (Get-QuickTime $q) }) }
        }
    }
    Render-Chips
    [void]$InputBox.Focus()
}

function Show-QuickDialog($q) {
    $title = if ($q) { '修改快捷时间' } else { '新增快捷时间' }
    $script:Dlg = New-DialogWindow $title
    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.Margin = Th 16 14 16 14
    $row = [System.Windows.Controls.StackPanel]::new()
    $row.Orientation = 'Horizontal'
    $script:QDay = [System.Windows.Controls.ComboBox]::new()
    $script:QDay.Width = 90; $script:QDay.Margin = Th 0 0 10 0
    $script:QDay.MaxDropDownHeight = 320
    $want = if ($q) { [string]$q.day } else { 'd1' }
    $script:QDayKeys = @($DayChoices.Keys)
    foreach ($k in $script:QDayKeys) { [void]$script:QDay.Items.Add($DayChoices[$k]) }
    $script:QDay.SelectedIndex = [math]::Max(0, [array]::IndexOf($script:QDayKeys, $want))
    $h = 18; $m = 0
    if ($q) { $hm = ([string]$q.time).Split(':'); $h = [int]$hm[0]; $m = [int]$hm[1] }
    $script:DlgHour = New-TimeCombo 23 $h
    $colon = [System.Windows.Controls.TextBlock]::new()
    $colon.Text = ':'; $colon.VerticalAlignment = 'Center'; $colon.Margin = Th 5 0 5 0
    $script:DlgMin = New-TimeCombo 59 $m
    [void]$row.Children.Add($script:QDay); [void]$row.Children.Add($script:DlgHour)
    [void]$row.Children.Add($colon); [void]$row.Children.Add($script:DlgMin)
    [void]$sp.Children.Add($row)

    $nl = [System.Windows.Controls.TextBlock]::new()
    $nl.Text = '标签上显示的名字（可以不填，不填就显示"明天 09:00"这样）'
    $nl.Margin = Th 0 12 0 4; $nl.Foreground = Brush '#6E6450'; $nl.TextWrapping = 'Wrap'; $nl.Width = 280
    [void]$sp.Children.Add($nl)
    $script:QName = [System.Windows.Controls.TextBox]::new()
    $script:QName.Padding = Th 4 3 4 3
    if ($q) { $script:QName.Text = [string]$q.name }
    [void]$sp.Children.Add($script:QName)

    $btns = New-DialogButtons
    [void]$sp.Children.Add($btns[0])
    $script:DlgResult = $null
    $btns[1].add_Click({
        $hh = 0; $mm = 0
        $okH = [int]::TryParse(([string]$script:DlgHour.Text).Trim(), [ref]$hh) -and $hh -ge 0 -and $hh -le 23
        $okM = [int]::TryParse(([string]$script:DlgMin.Text).Trim(), [ref]$mm) -and $mm -ge 0 -and $mm -le 59
        if (-not ($okH -and $okM) -or $script:QDay.SelectedIndex -lt 0) {
            [System.Windows.MessageBox]::Show($script:Dlg, '时间填得不对：小时填 0～23，分钟填 0～59。', '桌面待办') | Out-Null
            return
        }
        $script:DlgResult = @{
            name = $script:QName.Text.Trim()
            day = [string]$script:QDayKeys[$script:QDay.SelectedIndex]
            time = ('{0:00}:{1:00}' -f $hh, $mm)
        }
        $script:Dlg.DialogResult = $true
    })
    $script:Dlg.Content = $sp
    [void]$script:Dlg.ShowDialog()
    return $script:DlgResult
}

# 快捷列表变了以后，已选中的快捷项序号可能对不上，清掉只保留自己选的日期
function Reset-PendingQuick {
    $keep = @($script:Pending.sel | Where-Object { ([string]$_.key).StartsWith('c:') })
    $script:Pending.sel = New-Object System.Collections.ArrayList
    foreach ($x in $keep) { [void]$script:Pending.sel.Add($x) }
}

function Edit-Quick([int]$idx) {
    $list = New-Object System.Collections.ArrayList
    foreach ($q in (Get-QuickList)) { [void]$list.Add(@{ name = [string]$q.name; day = [string]$q.day; time = [string]$q.time }) }
    $old = if ($idx -ge 0 -and $idx -lt $list.Count) { $list[$idx] } else { $null }
    $r = Show-QuickDialog $old
    if (-not $r) { return }
    if ($old) { $list[$idx] = $r } else { [void]$list.Add($r) }
    $script:S.quick = @($list.ToArray())
    Save-Settings
    Reset-PendingQuick
    Render-Chips
}

function Remove-Quick([int]$idx) {
    $list = @(Get-QuickList)
    $script:S.quick = @(for ($i = 0; $i -lt $list.Count; $i++) { if ($i -ne $idx) { $list[$i] } })
    Save-Settings
    Reset-PendingQuick
    Render-Chips
}

$OnQuickEdit    = { param($s, $e) Edit-Quick ([int]$s.Tag) }
$OnQuickRemove  = { param($s, $e) Remove-Quick ([int]$s.Tag) }
$OnQuickDefault = { param($s, $e) $script:S.quick = @($DefaultQuick); Save-Settings; Reset-PendingQuick; Render-Chips }

function Render-TodoChips {
    $Chips.Children.Clear()
    if (-not $script:Pending.proj -or -not (Find-Project $script:Pending.proj)) { $script:Pending.proj = if ($script:S.lastProj -and (Find-Project $script:S.lastProj)) { [string]$script:S.lastProj } else { 'cat-daily' } }
    $pc = New-Chip ('项目：' + (Get-ProjName $script:Pending.proj) + ' ▾') '__proj' ($script:Pending.proj -ne 'cat-daily') '这条待办归到哪个项目；报销类可以在末尾写金额，比如"打车费 56"' $OnTodoProjChip
    $pjc = Find-Project $script:Pending.proj
    if ($pjc -and $pjc.kind -eq 'project') {
        $pcm = [System.Windows.Controls.ContextMenu]::new()
        [void]$pcm.Items.Add((New-MenuItem ('修改「' + $pjc.text + '」的名称…') $pjc.id $OnRenameProj))
        $pc.ContextMenu = $pcm
        $pc.ToolTip = '点一下选项目，右键可以改项目名称'
    }
    [void]$Chips.Children.Add($pc)
    $ql = Get-QuickList
    for ($i = 0; $i -lt $ql.Count; $i++) {
        $q = $ql[$i]
        $k = 'q:' + $i
        $sel = [bool](Find-PendingKey $k)
        $tip = (Format-When (Get-QuickTime $q)) + '。可以多选；右键可修改或移除这个快捷'
        $chip = New-Chip (Get-QuickName $q) $k $sel $tip $OnChip
        $cm = [System.Windows.Controls.ContextMenu]::new()
        [void]$cm.Items.Add((New-MenuItem '修改…' ([string]$i) $OnQuickEdit))
        [void]$cm.Items.Add((New-MenuItem '移除这个快捷' ([string]$i) $OnQuickRemove))
        [void]$cm.Items.Add(([System.Windows.Controls.Separator]::new()))
        [void]$cm.Items.Add((New-MenuItem '恢复默认快捷' '' $OnQuickDefault))
        $chip.ContextMenu = $cm
        [void]$Chips.Children.Add($chip)
    }
    foreach ($x in @($script:Pending.sel | Where-Object { ([string]$_.key).StartsWith('c:') } | Sort-Object { $_.dt })) {
        [void]$Chips.Children.Add((New-Chip ((Format-When ([datetime]$x.dt)) + ' ×') $x.key $true '点一下去掉这个时间' $OnChip))
    }
    [void]$Chips.Children.Add((New-Chip '选日期' '选日期' $false '自己选日期和时间，可以选多次' $OnChip))
    $star = New-Chip '★ 重要' '重要' $false '标为重要，排到最上面' $OnChip
    if ($script:Pending.star) { $star.Background = Brush '#E25C4B'; $star.Child.Foreground = Brush '#FFFFFF' }
    [void]$Chips.Children.Add($star)
    [void]$Chips.Children.Add((New-Chip '+' '__addquick' $false '新增一个快捷时间' $OnChip))
}

# ---------- 复制单条 / 项目管理 ----------
$script:FlashTimer = [System.Windows.Threading.DispatcherTimer]::new()
$script:FlashTimer.Interval = [TimeSpan]::FromMilliseconds(1200)
$script:FlashTimer.add_Tick({
    $script:FlashTimer.Stop()
    if ($script:FlashBtn) { $script:FlashBtn.Content = $Glyph.Copy; $script:FlashBtn.ToolTip = '复制这条'; $script:FlashBtn = $null }
})

function Copy-Text([string]$text, $sender) {
    try { [System.Windows.Clipboard]::SetText($text) }
    catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。'; return }
    if ($sender -is [System.Windows.Controls.Button]) {
        if ($script:FlashBtn) { $script:FlashBtn.Content = $Glyph.Copy }
        $sender.Content = $Glyph.Done
        $sender.ToolTip = '已复制'
        $script:FlashBtn = $sender
        $script:FlashTimer.Stop(); $script:FlashTimer.Start()
    }
}

$OnCopyTodo = { param($s, $e) $it = Find-Todo ([string]$s.Tag); if ($it) { Copy-TodoAll $it $s } }
$OnCopyWork = { param($s, $e) $w = Find-Work ([string]$s.Tag); if ($w) { Copy-Text $w.text $s } }

$OnCardEnter = { param($s, $e) try { $s.Child.Children[2].Opacity = 0.85 } catch {} }
$OnCardLeave = { param($s, $e) try { if (-not (Test-IsOpen ([string]$s.Tag))) { $s.Child.Children[2].Opacity = 0 } } catch {} }

function New-CardButtons([string]$id, [scriptblock]$onCopy, [scriptblock]$onDelete) {
    $panel = [System.Windows.Controls.StackPanel]::new()
    $panel.Orientation = 'Horizontal'
    $panel.VerticalAlignment = 'Top'
    $panel.Opacity = 0
    $cp = [System.Windows.Controls.Button]::new()
    $cp.Style = $IconBtnStyle; $cp.Content = $Glyph.Copy; $cp.Tag = $id; $cp.ToolTip = '复制整条：题目、具体事项和进展'
    $cp.add_Click($onCopy)
    $del = [System.Windows.Controls.Button]::new()
    $del.Style = $IconBtnStyle; $del.Content = $Glyph.Del; $del.Tag = $id; $del.ToolTip = '删除'
    $del.add_Click($onDelete)
    [void]$panel.Children.Add($cp); [void]$panel.Children.Add($del)
    [System.Windows.Controls.Grid]::SetColumn($panel, 2)
    return $panel
}

function Remove-Project([string]$p) {
    $script:S.projects = @(@($script:S.projects) | Where-Object { $_ -and ([string]$_) -ne $p })
    if ($script:LogPending.project -eq $p) { $script:LogPending.project = '' }
    Save-Settings
    Render-Chips
}

function Rename-Project([string]$old) {
    $r = Show-TextDialog '重命名项目' $old '改名后，已经记过的、属于这个项目的记录也会一起改'
    if (-not $r -or -not $r.Trim() -or $r.Trim() -eq $old) { return }
    $new = $r.Trim()
    foreach ($w in $script:Work) { if ($w.project -eq $old) { $w.project = $new } }
    Save-Work
    $script:S.projects = @(@($script:S.projects) | ForEach-Object { if (([string]$_) -eq $old) { $new } else { $_ } } | Select-Object -Unique)
    if ($script:LogPending.project -eq $old) { $script:LogPending.project = $new }
    Save-Settings
    Render-Chips
    Render
}

$OnProjRemove = { param($s, $e) Remove-Project ([string]$s.Tag) }
$OnProjRename = { param($s, $e) Rename-Project ([string]$s.Tag) }

# ---------- 待办展开：具体事项和备注 ----------
$script:Open = @{}
$script:FocusSubAdd = $null

function Get-SubText($it) {
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add($it.text)
    $i = 1
    foreach ($x in @($it.subs)) {
        $mark = if ($x.done) { '（已完成）' } else { '' }
        $body = if ($x.kind -eq 'image') { '[图片]' } else { $x.text }
        $lines.Add(('{0}. {1}{2}' -f $i, $body, $mark)); $i++
    }
    $pg = @(Get-TodoProgress $it)
    if ($pg.Count -gt 0) {
        $lines.Add('进展：')
        foreach ($w in $pg) { $lines.Add(('{0}：{1}' -f (Parse-Day $w.date).ToString('M月d日'), $w.text)) }
    }
    return ($lines -join "`r`n")
}

function Toggle-Open([string]$id) {
    if (Test-IsOpen $id) { $script:Open.Remove($id); $script:SearchOpen.Remove($id) } else { $script:Open[$id] = $true; $script:FocusSubAdd = $id }
    Save-Items
    Render
}

function Add-Sub([string]$id, [string]$text) {
    $it = Find-Todo $id
    if (-not $it -or -not $text.Trim()) { return }
    $it.subs = @($it.subs) + [pscustomobject]@{ id = [guid]::NewGuid().ToString('N'); kind = 'text'; text = $text.Trim(); file = ''; done = $false }
    Save-Items
    $script:FocusSubAdd = $id
    Render
}

function Find-Sub($it, [string]$sid) {
    foreach ($x in @($it.subs)) { if ($x.id -eq $sid) { return $x } }
    return $null
}

$OnToggleOpen = { param($s, $e) $e.Handled = $true; Toggle-Open ([string]$s.Tag) }
$OnSubCheck = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|'); $it = Find-Todo $p[0]; if (-not $it) { return }
    $x = Find-Sub $it $p[1]
    if ($x) {
        $x.done = [bool]$s.IsChecked
        Save-Items
        if ($x.done) { Add-AutoProgress $it $x } else { Remove-AutoProgress $it $x }
        Render
    }
}
$OnSubDelete = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|'); $it = Find-Todo $p[0]; if (-not $it) { return }
    $arr = @($it.subs); $x = $null; $idx = 0
    for ($i = 0; $i -lt $arr.Count; $i++) { if ($arr[$i].id -eq $p[1]) { $x = $arr[$i]; $idx = $i } }
    if (-not $x) { return }
    $it.subs = @($arr | Where-Object { $_.id -ne $p[1] }); Save-Items; Render
    Show-Undo ('已删除具体事项「' + (Short-Text $(if ($x.kind -eq 'image') { '图片' } else { $x.text })) + '」') @{ kind = 'sub'; todo = $it.id; item = $x; index = $idx; files = @($x) }
}
$OnSubEdit = {
    param($s, $e)
    if ($e.ClickCount -ne 2) { return }
    $e.Handled = $true
    $p = ([string]$s.Tag).Split('|'); $it = Find-Todo $p[0]; if (-not $it) { return }
    $x = Find-Sub $it $p[1]; if (-not $x -or $x.kind -eq 'image') { return }
    $r = Show-TextDialog '修改具体事项' $x.text '' $true
    if ($r -and $r.Trim()) { $old = $x.text; $x.text = $r.Trim(); Save-Items; Sync-AutoText $it $x $old; Render }
}
$OnSubAddKey = {
    param($s, $e)
    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    if ($e.Key -eq [System.Windows.Input.Key]::Return) { $e.Handled = $true; Add-Sub ([string]$s.Tag) $s.Text }
    elseif ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::V -and (Test-ClipHasOnlyImage)) {
        $imgs = Get-ClipImages
        if ($imgs.Count -gt 0) {
            $e.Handled = $true
            $id = [string]$s.Tag
            if ($s.Text.Trim()) { Add-Sub $id $s.Text }
            [void](Add-SubImages $id $imgs)
        }
    }
    elseif ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::V -and -not $s.Text.Trim()) {
        $txt = ''
        try { $txt = [System.Windows.Forms.Clipboard]::GetText() } catch {}
        $lines = @(($txt -split "\r?\n") | Where-Object { $_.Trim() })
        if ($lines.Count -gt 1) {
            $e.Handled = $true
            $it0 = Find-Todo ([string]$s.Tag)
            if ($it0) {
                foreach ($line in $lines) { $it0.subs = @($it0.subs) + [pscustomobject]@{ id = [guid]::NewGuid().ToString('N'); kind = 'text'; text = $line.Trim(); file = ''; done = $false } }
                Save-Items; $script:FocusSubAdd = $it0.id; Render
            }
        }
    }
}
$OnSubAddText = { param($s, $e) try { $s.Parent.Children[1].Visibility = if ($s.Text) { 'Collapsed' } else { 'Visible' } } catch {} }
$OnSubAddLoaded = {
    param($s, $e)
    if ($script:FocusSubAdd -eq [string]$s.Tag) { $script:FocusSubAdd = $null; [void]$s.Focus() }
}
$OnSwallow = { param($s, $e) $e.Handled = $true }
$OnSubRowEnter = { param($s, $e) try { $s.Children[1].Visibility = 'Visible' } catch {} }
$OnSubRowLeave = { param($s, $e) try { $s.Children[1].Visibility = 'Collapsed' } catch {} }

function New-HintBox([string]$tag, [string]$hint, [bool]$multi) {
    $g = [System.Windows.Controls.Grid]::new()
    $tb = [System.Windows.Controls.TextBox]::new()
    $tb.Tag = $tag
    $tb.Padding = Th 5 3 5 3
    $tb.TextWrapping = 'Wrap'
    $tb.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'InputBgBrush')
    $tb.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'TextBrush')
    $tb.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'InputLineBrush')
    $tb.SetResourceReference([System.Windows.Controls.TextBox]::CaretBrushProperty, 'TextBrush')
    if ($multi) { $tb.AcceptsReturn = $true; $tb.MinHeight = 46; $tb.MaxHeight = 180; $tb.VerticalScrollBarVisibility = 'Auto' }
    $h = [System.Windows.Controls.TextBlock]::new()
    $h.Text = $hint
    $h.FontSize = 12
    $h.Margin = Th 8 4 0 0
    $h.IsHitTestVisible = $false
    $h.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'HintBrush')
    [void]$g.Children.Add($tb); [void]$g.Children.Add($h)
    return @($g, $tb, $h)
}

function New-Details($it) {
    $panel = [System.Windows.Controls.StackPanel]::new()
    $panel.Margin = Th 0 8 6 2
    $panel.add_MouseLeftButtonDown($OnSwallow)
    $panel.Tag = $it.id
    $panel.AllowDrop = $true
    $panel.Background = [System.Windows.Media.Brushes]::Transparent
    $panel.add_DragOver($OnDetailsDragOver)
    $panel.add_Drop($OnDetailsDrop)

    $subNo = 0
    foreach ($x in @($it.subs)) {
        $subNo++
        $row = [System.Windows.Controls.Grid]::new()
        $row.Margin = Th 0 0 0 4
        $row.Background = [System.Windows.Media.Brushes]::Transparent
        $line = [System.Windows.Controls.DockPanel]::new()
        $cb = [System.Windows.Controls.CheckBox]::new()
        $cb.IsChecked = [bool]$x.done
        $cb.Tag = $it.id + '|' + $x.id
        $cb.Margin = Th 0 2 6 0; $cb.VerticalAlignment = 'Top'; $cb.Cursor = $Hand
        $cb.add_Click($OnSubCheck)
        [System.Windows.Controls.DockPanel]::SetDock($cb, 'Left')
        [void]$line.Children.Add($cb)
        $num = New-NumLabel $subNo
        [System.Windows.Controls.DockPanel]::SetDock($num, 'Left')
        [void]$line.Children.Add($num)
        if ($x.kind -eq 'image') {
            $th = New-Thumb $x.file ($it.id + '|' + $x.id) $OnSubImgOpen 180
            if ($x.done) { $th.Opacity = 0.5 }
            [void]$line.Children.Add($th)
        } else {
            $t = New-SelectableText $x.text 0
            [void](Enable-Inline $t ('sub|' + $it.id + '|' + $x.id))
            if ($x.done) { Set-DoneLook $t }
            [void]$line.Children.Add($t)
        }
        [void]$row.Children.Add($line)

        $tools = [System.Windows.Controls.Border]::new()
        $tools.HorizontalAlignment = 'Right'
        $tools.VerticalAlignment = 'Top'
        $tools.CornerRadius = [System.Windows.CornerRadius]::new(4)
        $tools.Padding = Th 2 0 0 0
        $tools.Visibility = 'Collapsed'
        $tools.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'CardBrush')
        $tp = [System.Windows.Controls.StackPanel]::new()
        $tp.Orientation = 'Horizontal'
        $cpy = [System.Windows.Controls.Button]::new()
        $cpy.Style = $IconBtnStyle; $cpy.Content = $Glyph.Copy; $cpy.Width = 22; $cpy.Height = 20
        $cpy.Tag = $it.id + '|' + $x.id; $cpy.ToolTip = '复制这一项'
        $cpy.add_Click($OnSubCopy)
        [void]$tp.Children.Add($cpy)
        if ($x.kind -ne 'image') { [void]$tp.Children.Add((New-EditBtn ($it.id + '|' + $x.id) $OnSubEditBtn 20)) }
        $del = [System.Windows.Controls.Button]::new()
        $del.Style = $IconBtnStyle; $del.Content = $Glyph.Close; $del.Width = 22; $del.Height = 20
        $del.Tag = $it.id + '|' + $x.id; $del.ToolTip = '删除这一项'
        $del.add_Click($OnSubDelete)
        [void]$tp.Children.Add($del)
        $tools.Child = $tp
        [void]$row.Children.Add($tools)
        $row.add_MouseEnter($OnSubRowEnter)
        $row.add_MouseLeave($OnSubRowLeave)
        [void]$panel.Children.Add($row)
    }

    $add = New-HintBox $it.id '加一条具体事项，回车（截图可以直接 Ctrl+V）' $false
    $add[1].add_PreviewKeyDown($OnSubAddKey)
    $add[1].add_TextChanged($OnSubAddText)
    $add[1].add_Loaded($OnSubAddLoaded)
    $addRow = [System.Windows.Controls.DockPanel]::new()
    $addRow.Margin = Th 0 2 0 6
    $pick = [System.Windows.Controls.Button]::new()
    $pick.Style = $IconBtnStyle; $pick.Content = [string][char]0xE8B9; $pick.Margin = Th 4 0 0 0
    $pick.Tag = $it.id; $pick.ToolTip = '从电脑里选图片，每张存成一项'
    $pick.add_Click($OnSubPickBtn)
    [System.Windows.Controls.DockPanel]::SetDock($pick, 'Right')
    [void]$addRow.Children.Add($pick)
    [void]$addRow.Children.Add($add[0])
    [void]$panel.Children.Add($addRow)

    return $panel
}

function Get-DetailSummary($it) {
    $n = @($it.subs).Count
    if ($n -eq 0) { return '' }
    return ('具体事项 {0}/{1}' -f @(@($it.subs) | Where-Object { $_.done }).Count, $n)
}

# 卡片上的项目标签：点一下换项目，已完成的也能换（归档用）
function New-ProjTag($it) {
    $b = [System.Windows.Controls.Border]::new()
    $b.CornerRadius = [System.Windows.CornerRadius]::new(3)
    $b.Padding = Th 5 0 4 1
    $b.Margin = Th 0 0 8 2
    $b.VerticalAlignment = 'Top'
    $b.Cursor = $Hand
    $b.Tag = $it.id
    $b.ToolTip = '点一下换项目（已完成的也能换），这件事的进展跟着一起换'
    $b.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'ChipBrush')
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = (Get-ProjName (Get-TodoProj $it)) + ' ▾'
    $t.FontSize = 11
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'ChipTextBrush')
    $b.Child = $t
    $b.add_MouseLeftButtonDown($OnSwallow)
    $b.add_MouseLeftButtonUp($OnProjTagClick)
    [System.Windows.Controls.DockPanel]::SetDock($b, 'Left')
    return $b
}
$OnProjTagClick = {
    param($s, $e)
    $e.Handled = $true
    $it = Find-Todo ([string]$s.Tag); if (-not $it) { return }
    Open-MenuAt (New-ProjectMenu $OnTodoSetProj (Get-TodoProj $it) ($it.id + '|')) $s
}

# ---------- 备注：多条文字 / 图片 ----------
$ImgDir = Join-Path $DataDir 'images'
$ImgExt = @('.png', '.jpg', '.jpeg', '.bmp', '.gif')
$script:FocusNote = $null

function New-Note([string]$kind, [string]$text, [string]$file) {
    [pscustomobject]@{ id = [guid]::NewGuid().ToString('N'); kind = $kind; text = $text; file = $file; at = (Get-Date).ToString($Fmt) }
}

function Get-ImgDir { return (Join-Path $script:ShareDir 'images') }
function Get-ImgPath([string]$file) { return (Join-Path (Get-ImgDir) $file) }

function Add-NoteText([string]$id, [string]$text) {
    $it = Find-Todo $id
    if (-not $it) { return }
    $text = $text.Trim()
    if (-not $text) { return }
    $it.notes = @($it.notes) + (New-Note 'text' $text '')
    Save-Items
    $script:FocusNote = $id
    Render
}

function Add-NoteImage([string]$id, $drawingImage) {
    $it = Find-Todo $id
    if (-not $it -or -not $drawingImage) { return }
    $name = Save-ClipImage $drawingImage
    $it.notes = @($it.notes) + (New-Note 'image' '' $name)
    Save-Items
    $script:FocusNote = $id
    Render
}

function Add-NoteImageFile([string]$id, [string]$path) {
    if (-not (Test-Path -LiteralPath $path)) { return }
    if ($ImgExt -notcontains ([IO.Path]::GetExtension($path).ToLower())) { return }
    $img = Load-DrawingImage $path
    if ($img) { Add-NoteImage $id $img; $img.Dispose() }
}

# 读进内存再用，不锁住原文件
function Load-DrawingImage([string]$path) {
    try {
        $bytes = [IO.File]::ReadAllBytes($path)
        $ms = New-Object IO.MemoryStream(, $bytes)
        return [System.Drawing.Image]::FromStream($ms)
    } catch { Write-Log ('读取图片失败：' + $_.Exception.Message); return $null }
}

function Paste-NoteImage([string]$id) {
    $done = $false
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $img = [System.Windows.Forms.Clipboard]::GetImage()
            if ($img) { Add-NoteImage $id $img; $img.Dispose(); $done = $true }
        } elseif ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
            foreach ($f in [System.Windows.Forms.Clipboard]::GetFileDropList()) {
                if ($ImgExt -contains ([IO.Path]::GetExtension($f).ToLower())) { Add-NoteImageFile $id $f; $done = $true }
            }
        }
    } catch { Write-Log ('粘贴图片失败：' + $_.Exception.Message) }
    return $done
}

function Pick-NoteImage([string]$id) {
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Filter = '图片|*.png;*.jpg;*.jpeg;*.bmp;*.gif'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    foreach ($f in $dlg.FileNames) { Add-NoteImageFile $id $f }
}

function Find-Note($it, [string]$nid) {
    foreach ($n in @($it.notes)) { if ($n.id -eq $nid) { return $n } }
    return $null
}

function Remove-NoteFiles($notes) {
    foreach ($n in @($notes)) {
        if ($n.kind -eq 'image' -and $n.file) { Remove-Item -LiteralPath (Get-ImgPath $n.file) -Force -ErrorAction SilentlyContinue }
    }
}

function Copy-Note($n, $sender) {
    if ($n.kind -eq 'image') {
        $img = Load-DrawingImage (Get-ImgPath $n.file)
        if (-not $img) { Notify '复制失败' '图片文件找不到了。'; return }
        try { [System.Windows.Forms.Clipboard]::SetImage($img) } catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。'; return }
        if ($sender -is [System.Windows.Controls.Button]) {
            if ($script:FlashBtn) { $script:FlashBtn.Content = $Glyph.Copy }
            $sender.Content = $Glyph.Done; $script:FlashBtn = $sender
            $script:FlashTimer.Stop(); $script:FlashTimer.Start()
        }
    } else {
        Copy-Text $n.text $sender
    }
}

function Split-NoteTag([string]$tag) { $p = $tag.Split('|'); $it = Find-Todo $p[0]; if (-not $it) { return $null }; return @($it, (Find-Note $it $p[1])) }

$OnNoteCopy = { param($s, $e) $r = Split-NoteTag ([string]$s.Tag); if ($r -and $r[1]) { Copy-Note $r[1] $s } }
$OnNoteDelete = {
    param($s, $e)
    $r = Split-NoteTag ([string]$s.Tag); if (-not $r -or -not $r[1]) { return }
    $ans = [System.Windows.MessageBox]::Show($window, '删除这条备注？', '桌面待办', 'YesNo', 'Question')
    if ($ans -ne 'Yes') { return }
    Remove-NoteFiles @($r[1])
    $r[0].notes = @(@($r[0].notes) | Where-Object { $_.id -ne $r[1].id })
    Save-Items; Render
}
$OnNoteEdit = {
    param($s, $e)
    $r = Split-NoteTag ([string]$s.Tag); if (-not $r -or -not $r[1] -or $r[1].kind -ne 'text') { return }
    $t = Show-TextDialog '修改备注' $r[1].text '' $true
    if ($null -ne $t -and $t.Trim()) { $r[1].text = $t.Trim(); Save-Items; Render }
}
$OnNoteUp = {
    param($s, $e)
    $r = Split-NoteTag ([string]$s.Tag); if (-not $r -or -not $r[1]) { return }
    $arr = New-Object System.Collections.ArrayList; foreach ($n in @($r[0].notes)) { [void]$arr.Add($n) }
    $i = $arr.IndexOf($r[1]); if ($i -le 0) { return }
    $arr.RemoveAt($i); $arr.Insert($i - 1, $r[1]); $r[0].notes = @($arr.ToArray()); Save-Items; Render
}
$OnImgOpen = {
    param($s, $e)
    $e.Handled = $true
    $r = Split-NoteTag ([string]$s.Tag); if (-not $r -or -not $r[1]) { return }
    $f = Get-ImgPath $r[1].file
    if (Test-Path -LiteralPath $f) { Start-Process $f }
}
$OnNoteKey = {
    param($s, $e)
    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    if ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::Return) {
        $e.Handled = $true; Add-NoteText ([string]$s.Tag) $s.Text
    } elseif ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::V) {
        $hasText = $false
        try { $hasText = [System.Windows.Forms.Clipboard]::ContainsText() } catch {}
        if (-not $hasText) {
            $pending = $s.Text
            if (Paste-NoteImage ([string]$s.Tag)) {
                $e.Handled = $true
                if ($pending.Trim()) { Add-NoteText ([string]$s.Tag) $pending }
            }
        }
    }
}
$OnNoteLoaded = { param($s, $e) if ($script:FocusNote -eq [string]$s.Tag) { $script:FocusNote = $null; [void]$s.Focus() } }
$OnNoteAddBtn = { param($s, $e) $tb = $s.Parent.Tag; Add-NoteText ([string]$tb.Tag) $tb.Text }
$OnNotePasteBtn = {
    param($s, $e)
    $id = [string]$s.Tag
    if (-not (Paste-NoteImage $id)) {
        $txt = ''
        try { $txt = [System.Windows.Forms.Clipboard]::GetText() } catch {}
        if ($txt.Trim()) { Add-NoteText $id $txt } else { Notify '剪贴板里没有图片或文字' '先截图或复制一段文字，再点粘贴。' }
    }
}
$OnNotePickBtn = { param($s, $e) Pick-NoteImage ([string]$s.Tag) }
$OnDetailsDragOver = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { $e.Effects = 'Copy' } else { $e.Effects = 'None' }
    $e.Handled = $true
}
$OnDetailsDrop = {
    param($s, $e)
    if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
    $imgs = New-Object System.Collections.ArrayList
    foreach ($f in $e.Data.GetData([System.Windows.DataFormats]::FileDrop)) {
        if ($ImgExt -contains ([IO.Path]::GetExtension($f).ToLower())) { $img = Load-DrawingImage $f; if ($img) { [void]$imgs.Add($img) } }
    }
    $e.Handled = $true
    [void](Add-SubImages ([string]$s.Tag) $imgs)
}

function New-SmallBtn([string]$glyph, [string]$tag, [string]$tip, [scriptblock]$h) {
    $b = [System.Windows.Controls.Button]::new()
    $b.Style = $IconBtnStyle; $b.Content = $glyph; $b.Width = 22; $b.Height = 20
    $b.Tag = $tag; $b.ToolTip = $tip; $b.add_Click($h)
    return $b
}

function New-NoteView($it, $n, [bool]$first, [int]$no) {
    $tag = $it.id + '|' + $n.id
    $box = [System.Windows.Controls.Border]::new()
    $box.Margin = Th 0 0 0 6
    $box.Padding = Th 6 4 2 5
    $box.CornerRadius = [System.Windows.CornerRadius]::new(4)
    $box.BorderThickness = Th 2 0 0 0
    $box.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'AccentBrush')
    $box.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'BgBrush')
    $sp = [System.Windows.Controls.StackPanel]::new()

    $head = [System.Windows.Controls.DockPanel]::new()
    $btns = [System.Windows.Controls.StackPanel]::new()
    $btns.Orientation = 'Horizontal'; $btns.Opacity = 0.7
    [System.Windows.Controls.DockPanel]::SetDock($btns, 'Right')
    if (-not $first) { [void]$btns.Children.Add((New-SmallBtn $Glyph.Up $tag '上移' $OnNoteUp)) }
    [void]$btns.Children.Add((New-SmallBtn $Glyph.Copy $tag '复制这条' $OnNoteCopy))
    if ($n.kind -eq 'text') { [void]$btns.Children.Add((New-SmallBtn $Glyph.Edit $tag '修改' $OnNoteEdit)) }
    [void]$btns.Children.Add((New-SmallBtn $Glyph.Close $tag '删除' $OnNoteDelete))
    [void]$head.Children.Add($btns)
    $tm = [System.Windows.Controls.TextBlock]::new()
    $at = Parse-Dt $n.at
    $tm.Text = ('{0}.  ' -f $no) + $(if ($at) { Format-When $at } else { '' })
    $tm.FontSize = 11; $tm.VerticalAlignment = 'Center'
    $tm.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
    [void]$head.Children.Add($tm)
    [void]$sp.Children.Add($head)

    if ($n.kind -eq 'image') {
        [void]$sp.Children.Add((New-Thumb $n.file $tag $OnImgOpen 220))
    } else {
        # 只读文本框：可以用鼠标选中其中一段复制
        $tb = [System.Windows.Controls.TextBox]::new()
        $tb.Text = $n.text
        $tb.IsReadOnly = $true
        $tb.TextWrapping = 'Wrap'
        $tb.BorderThickness = Th 0 0 0 0
        $tb.Background = [System.Windows.Media.Brushes]::Transparent
        $tb.Padding = Th 0 2 4 0
        $tb.MaxHeight = 240
        $tb.VerticalScrollBarVisibility = 'Auto'
        $tb.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'TextBrush')
        [void]$sp.Children.Add($tb)
    }
    $box.Child = $sp
    return $box
}

function New-NotesSection($it) {
    $wrap = [System.Windows.Controls.StackPanel]::new()
    $lab = [System.Windows.Controls.TextBlock]::new()
    $lab.Text = '备注'
    $lab.FontSize = 12; $lab.Margin = Th 0 4 0 4
    $lab.FontWeight = [System.Windows.FontWeights]::SemiBold
    $lab.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
    [void]$wrap.Children.Add($lab)

    $list = @($it.notes)
    for ($i = 0; $i -lt $list.Count; $i++) { [void]$wrap.Children.Add((New-NoteView $it $list[$i] ($i -eq 0) ($i + 1))) }

    $in = New-HintBox $it.id '记一条：领导原话、联系人、材料位置… 可以直接 Ctrl+V 粘贴截图；Ctrl+回车保存' $true
    $in[1].add_PreviewKeyDown($OnNoteKey)
    $in[1].add_TextChanged($OnSubAddText)
    $in[1].add_Loaded($OnNoteLoaded)
    $in[2].TextWrapping = 'Wrap'; $in[2].Margin = Th 8 4 8 0
    [void]$wrap.Children.Add($in[0])

    $bar = [System.Windows.Controls.WrapPanel]::new()
    $bar.Margin = Th 0 5 0 0
    $bar.Tag = $in[1]
    foreach ($b in @(@('保存这条', $OnNoteAddBtn, '把上面输入的文字存成一条备注（Ctrl+回车）'),
                    @('粘贴截图/文字', $OnNotePasteBtn, '把剪贴板里的截图或文字直接存成一条'),
                    @('选图片…', $OnNotePickBtn, '从电脑里选图片，也可以直接把图片文件拖到这里'))) {
        $t = [System.Windows.Controls.TextBlock]::new()
        $t.Text = $b[0]; $t.FontSize = 12; $t.Margin = Th 0 0 14 0; $t.Cursor = $Hand
        $t.Tag = $it.id; $t.ToolTip = $b[2]
        $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
        $h = $b[1]
        $t.add_MouseLeftButtonUp($h)
        [void]$bar.Children.Add($t)
    }
    [void]$wrap.Children.Add($bar)
    return $wrap
}

# ---------- 具体事项里放图片 ----------
function Save-ClipImage($drawingImage) {
    if (-not (Test-Path (Get-ImgDir))) { New-Item -ItemType Directory -Force -Path (Get-ImgDir) | Out-Null }
    $name = [guid]::NewGuid().ToString('N') + '.png'
    $bmp = New-Object System.Drawing.Bitmap $drawingImage
    $bmp.Save((Get-ImgPath $name), [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    return $name
}

# 从剪贴板取图片：返回 System.Drawing.Image 列表（截图或复制的图片文件）
function Get-ClipImages {
    $list = New-Object System.Collections.ArrayList
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsImage()) {
            $img = [System.Windows.Forms.Clipboard]::GetImage()
            if ($img) { [void]$list.Add($img) }
        } elseif ([System.Windows.Forms.Clipboard]::ContainsFileDropList()) {
            foreach ($f in [System.Windows.Forms.Clipboard]::GetFileDropList()) {
                if ($ImgExt -contains ([IO.Path]::GetExtension($f).ToLower())) {
                    $img = Load-DrawingImage $f
                    if ($img) { [void]$list.Add($img) }
                }
            }
        }
    } catch { Write-Log ('读取剪贴板图片失败：' + $_.Exception.Message) }
    return , $list
}

function Test-ClipHasOnlyImage {
    try {
        if ([System.Windows.Forms.Clipboard]::ContainsText()) { return $false }
        return ([System.Windows.Forms.Clipboard]::ContainsImage() -or [System.Windows.Forms.Clipboard]::ContainsFileDropList())
    } catch { return $false }
}

function Add-SubImages([string]$id, $images) {
    $it = Find-Todo $id
    if (-not $it -or $images.Count -eq 0) { return $false }
    foreach ($img in $images) {
        $name = Save-ClipImage $img
        $img.Dispose()
        $it.subs = @($it.subs) + [pscustomobject]@{ id = [guid]::NewGuid().ToString('N'); kind = 'image'; text = ''; file = $name; done = $false }
    }
    Save-Items
    $script:FocusSubAdd = $id
    Render
    return $true
}

function Copy-Sub($x, $sender) {
    if ($x.kind -eq 'image') { Copy-Note ([pscustomobject]@{ kind = 'image'; file = $x.file }) $sender }
    else { Copy-Text $x.text $sender }
}

$OnSubCopy = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|'); $it = Find-Todo $p[0]; if (-not $it) { return }
    $x = Find-Sub $it $p[1]; if ($x) { Copy-Sub $x $s }
}
$OnSubImgOpen = {
    param($s, $e)
    $e.Handled = $true
    $p = ([string]$s.Tag).Split('|'); $it = Find-Todo $p[0]; if (-not $it) { return }
    $x = Find-Sub $it $p[1]
    if ($x -and $x.file) { $f = Get-ImgPath $x.file; if (Test-Path -LiteralPath $f) { Start-Process $f } }
}
$OnSubPasteBtn = {
    param($s, $e)
    $id = [string]$s.Tag
    $imgs = Get-ClipImages
    if ($imgs.Count -gt 0) { [void](Add-SubImages $id $imgs); return }
    $txt = ''
    try { $txt = [System.Windows.Forms.Clipboard]::GetText() } catch {}
    if ($txt.Trim()) { foreach ($line in ($txt -split "\r?\n")) { if ($line.Trim()) { Add-Sub $id $line } } }
    else { Notify '剪贴板里没有图片或文字' '先截图或复制一段文字，再点粘贴。' }
}
$OnSubPickBtn = {
    param($s, $e)
    $id = [string]$s.Tag
    $dlg = [System.Windows.Forms.OpenFileDialog]::new()
    $dlg.Filter = '图片|*.png;*.jpg;*.jpeg;*.bmp;*.gif'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
    $imgs = New-Object System.Collections.ArrayList
    foreach ($f in $dlg.FileNames) { $img = Load-DrawingImage $f; if ($img) { [void]$imgs.Add($img) } }
    [void](Add-SubImages $id $imgs)
}

function New-Thumb([string]$file, [string]$tag, [scriptblock]$onClick, [int]$maxH) {
    $f = Get-ImgPath $file
    if (-not (Test-Path -LiteralPath $f)) {
        $miss = [System.Windows.Controls.TextBlock]::new()
        $miss.Text = '（图片还没同步过来，或者已经删了）'; $miss.Foreground = TB 'Muted'
        return $miss
    }
    try {
        $bi = [System.Windows.Media.Imaging.BitmapImage]::new()
        $bi.BeginInit()
        $bi.CacheOption = [System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad
        $bi.DecodePixelWidth = 600
        $bi.UriSource = [Uri]::new($f)
        $bi.EndInit()
        $img = [System.Windows.Controls.Image]::new()
        $img.Source = $bi
        $img.Stretch = 'Uniform'
        $img.StretchDirection = 'DownOnly'
        $img.MaxHeight = $maxH
        $img.HorizontalAlignment = 'Left'
        $img.Margin = Th 0 3 4 0
        $img.Cursor = $Hand
        $img.Tag = $tag
        $img.ToolTip = '点击用看图软件打开'
        $img.add_MouseLeftButtonUp($onClick)
        return $img
    } catch {
        Write-Log ('显示图片失败：' + $_.Exception.Message)
        $err = [System.Windows.Controls.TextBlock]::new()
        $err.Text = '（图片打不开）'; $err.Foreground = TB 'Muted'
        return $err
    }
}

function New-NumLabel([int]$n) {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = '{0}.' -f $n
    $t.MinWidth = 18
    $t.Margin = Th 0 0 4 0
    $t.VerticalAlignment = 'Top'
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
    return $t
}

function New-LinkBar([string]$id, $items) {
    $bar = [System.Windows.Controls.WrapPanel]::new()
    $bar.Margin = Th 0 4 0 0
    foreach ($b in $items) {
        $t = [System.Windows.Controls.TextBlock]::new()
        $t.Text = $b[0]; $t.FontSize = 12; $t.Margin = Th 0 0 14 0; $t.Cursor = $Hand
        $t.Tag = $id; $t.ToolTip = $b[2]
        $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
        $h = $b[1]
        $t.add_MouseLeftButtonUp($h)
        [void]$bar.Children.Add($t)
    }
    return $bar
}

# ---------- 可选中文字 / 整条复制（含图片） ----------
function New-SelectableText([string]$text, [double]$size) {
    $tb = [System.Windows.Controls.TextBox]::new()
    $tb.Text = $text
    $tb.IsReadOnly = $true
    $tb.IsReadOnlyCaretVisible = $false
    $tb.TextWrapping = 'Wrap'
    $tb.BorderThickness = Th 0 0 0 0
    $tb.Background = [System.Windows.Media.Brushes]::Transparent
    $tb.Padding = Th 0 0 0 0
    $tb.Margin = Th -2 0 0 0
    if ($size -gt 0) { $tb.FontSize = $size }
    $tb.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'TextBrush')
    return $tb
}

function Set-DoneLook($tb) {
    $tb.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'MutedBrush')
    try { $tb.TextDecorations = [System.Windows.TextDecorations]::Strikethrough } catch {}
}

function Flash-Copied($sender) {
    if ($sender -is [System.Windows.Controls.Button]) {
        if ($script:FlashBtn) { $script:FlashBtn.Content = $Glyph.Copy }
        $sender.Content = $Glyph.Done
        $sender.ToolTip = '已复制'
        $script:FlashBtn = $sender
        $script:FlashTimer.Stop(); $script:FlashTimer.Start()
    }
}

# 右键菜单第一项：复制鼠标选中的那一段；没选中就复制整段
$OnCopySel = {
    param($s, $e)
    $tb = $null
    try { $tb = $s.Parent.PlacementTarget } catch {}
    if ($tb -is [System.Windows.Controls.TextBox]) {
        if ($tb.SelectionLength -gt 0) { $tb.Copy() } else { Copy-Text $tb.Text $null }
    }
}

# 先挂一个空菜单，真正点右键（或双击说明行）时才把内容建出来
function Set-LazyMenu($el, [scriptblock]$builder, $arg, $selBox = $null) {
    $cm = [System.Windows.Controls.ContextMenu]::new()
    $cm.Tag = @{ b = $builder; a = $arg; tb = $selBox; done = $false }
    $ph = [System.Windows.Controls.MenuItem]::new(); $ph.Header = '…'; $ph.IsEnabled = $false
    [void]$cm.Items.Add($ph)
    $el.ContextMenu = $cm
    $el.add_ContextMenuOpening($OnLazyMenu)
    if ($selBox -and $selBox -ne $el) { $selBox.ContextMenu = $cm; $selBox.add_ContextMenuOpening($OnLazyMenu) }
}
function Fill-LazyMenu($cm) {
    if (-not $cm -or -not ($cm.Tag -is [hashtable]) -or $cm.Tag.done) { return }
    $info = $cm.Tag
    $info.done = $true
    $built = & $info.b $info.a
    if (-not $built) { return }
    $items = @($built.Items); $built.Items.Clear()
    $cm.Items.Clear()
    foreach ($x in $items) { [void]$cm.Items.Add($x) }
    if ($info.tb) {
        $mi = [System.Windows.Controls.MenuItem]::new()
        $mi.Header = '复制选中的文字'; $mi.InputGestureText = 'Ctrl+C'; $mi.add_Click($OnCopySel)
        $cm.Items.Insert(0, $mi)
        $cm.Items.Insert(1, [System.Windows.Controls.Separator]::new())
    }
}
$OnLazyMenu = { param($s, $e) Fill-LazyMenu $s.ContextMenu }

function Add-SelMenu($menu, $tb) {
    $mi = [System.Windows.Controls.MenuItem]::new()
    $mi.Header = '复制选中的文字'
    $mi.InputGestureText = 'Ctrl+C'
    $mi.add_Click($OnCopySel)
    $menu.Items.Insert(0, $mi)
    $menu.Items.Insert(1, ([System.Windows.Controls.Separator]::new()))
    $tb.ContextMenu = $menu
}

function ConvertTo-CfHtml([string]$fragment) {
    $pre = "<html><head><meta charset=""utf-8""></head><body><!--StartFragment-->"
    $post = "<!--EndFragment--></body></html>"
    $hdrTpl = "Version:0.9`r`nStartHTML:{0:D10}`r`nEndHTML:{1:D10}`r`nStartFragment:{2:D10}`r`nEndFragment:{3:D10}`r`n"
    $enc = New-Object System.Text.UTF8Encoding($false)
    $hdrLen = $enc.GetByteCount(($hdrTpl -f 0, 0, 0, 0))
    $startHtml = $hdrLen
    $startFrag = $startHtml + $enc.GetByteCount($pre)
    $endFrag = $startFrag + $enc.GetByteCount($fragment)
    $endHtml = $endFrag + $enc.GetByteCount($post)
    return (($hdrTpl -f $startHtml, $endHtml, $startFrag, $endFrag) + $pre + $fragment + $post)
}

function HtmlLine([string]$t) { return ([System.Net.WebUtility]::HtmlEncode($t) -replace "\r?\n", '<br>') }

function Get-ImgHtml([string]$file) {
    $f = Get-ImgPath $file
    if (-not (Test-Path -LiteralPath $f)) { return '[图片]' }
    return ('<img src="{0}" style="max-width:560px">' -f ([Uri]::new($f)).AbsoluteUri)
}

# 整条待办：纯文字 + 带图片的网页格式一起放进剪贴板（Word、WPS 里粘贴会带上图片）
function Copy-TodoAll($it, $sender) {
    $text = Get-SubText $it
    $h = New-Object System.Text.StringBuilder
    [void]$h.Append('<p><b>' + (HtmlLine $it.text) + '</b></p>')
    $i = 1
    foreach ($x in @($it.subs)) {
        $mark = if ($x.done) { '（已完成）' } else { '' }
        $body = if ($x.kind -eq 'image') { '<br>' + (Get-ImgHtml $x.file) } else { HtmlLine $x.text }
        [void]$h.Append(('<p>{0}. {1}{2}</p>' -f $i, $body, $mark)); $i++
    }
    $pg = @(Get-TodoProgress $it)
    if ($pg.Count -gt 0) {
        [void]$h.Append('<p>进展：</p>')
        foreach ($w in $pg) { [void]$h.Append(('<p>{0}：{1}</p>' -f (Parse-Day $w.date).ToString('M月d日'), (HtmlLine $w.text))) }
    }
    try {
        $data = [System.Windows.Forms.DataObject]::new()
        $data.SetData([System.Windows.Forms.DataFormats]::UnicodeText, $text)
        $data.SetData([System.Windows.Forms.DataFormats]::Text, $text)
        $bytes = (New-Object System.Text.UTF8Encoding($false)).GetBytes((ConvertTo-CfHtml $h.ToString()))
        $data.SetData([System.Windows.Forms.DataFormats]::Html, (New-Object IO.MemoryStream(, $bytes)))
        [System.Windows.Forms.Clipboard]::SetDataObject($data, $true)
    } catch {
        Write-Log ('整条复制失败，改为只复制文字：' + $_.Exception.Message)
        Copy-Text $text $sender
        return
    }
    Flash-Copied $sender
}

$OnEditBtnTodo = { param($s, $e) Edit-Todo ([string]$s.Tag) }
$OnEditBtnWork = { param($s, $e) Edit-Work ([string]$s.Tag) }
$OnSubEditBtn = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|'); $it = Find-Todo $p[0]; if (-not $it) { return }
    $x = Find-Sub $it $p[1]; if (-not $x -or $x.kind -eq 'image') { return }
    $r = Show-TextDialog '修改具体事项' $x.text '' $true
    if ($r -and $r.Trim()) { $old = $x.text; $x.text = $r.Trim(); Save-Items; Sync-AutoText $it $x $old; Render }
}

function New-EditBtn([string]$tag, [scriptblock]$h, [int]$size) {
    $b = [System.Windows.Controls.Button]::new()
    $b.Style = $IconBtnStyle; $b.Content = $Glyph.Edit; $b.Tag = $tag; $b.ToolTip = '修改'
    if ($size -gt 0) { $b.Width = $size; $b.Height = $size }
    $b.add_Click($h)
    return $b
}

# ---------- 搜索、全部展开/收起 ----------
$script:Query = ''
$script:SearchOpen = @{}
$script:VisibleIds = @()

function Get-Words { return @(([string]$script:Query).Trim() -split '\s+' | Where-Object { $_ }) }

function Test-Words([string]$hay, $words) {
    foreach ($w in $words) { if ($hay.IndexOf($w, [StringComparison]::OrdinalIgnoreCase) -lt 0) { return $false } }
    return $true
}

function Get-SubsHay($it) { return ((@($it.subs) | Where-Object { $_.kind -ne 'image' } | ForEach-Object { $_.text }) -join "`n") }

function Test-TodoMatch($it, $words) {
    if ($words.Count -eq 0) { return $true }
    return (Test-Words ($it.text + "`n" + (Get-SubsHay $it) + "`n" + (Get-ProjName $it.project)) $words)
}

function Test-WorkMatch($w, $words) {
    if ($words.Count -eq 0) { return $true }
    $tt = if ($w.src) { $t = Find-Todo $w.src; if ($t) { $t.text } else { '' } } else { '' }
    return (Test-Words ($w.text + "`n" + (Get-ProjName $w.proj) + "`n" + $tt + "`n" + $w.date) $words)
}

function Test-IsOpen([string]$id) { return ($script:Open.ContainsKey($id) -or $script:SearchOpen.ContainsKey($id)) }

function Show-Search {
    $SearchBar.Visibility = 'Visible'
    [void]$SearchBox.Focus()
    $SearchBox.SelectAll()
}

# 搜到的字在具体事项里的，自动展开那条（只在搜索词变化时算一次，之后可以手动收起）
function Update-SearchOpen {
    $script:SearchOpen = @{}
    $words = Get-Words
    if ($words.Count -eq 0) { return }
    foreach ($it in @($script:Items)) {
        if (@($it.subs).Count -gt 0 -and (Test-Words (Get-SubsHay $it) @($words | Select-Object -First 1))) { $script:SearchOpen[$it.id] = $true }
    }
}

function Close-Search {
    $SearchBox.Text = ''
    $SearchBar.Visibility = 'Collapsed'
    [void]$InputBox.Focus()
}

function Toggle-ExpandAll {
    $closed = @($script:VisibleIds | Where-Object { -not (Test-IsOpen $_) })
    if ($closed.Count -gt 0) {
        foreach ($id in $script:VisibleIds) { $script:Open[$id] = $true; $script:ProgOpen[$id] = $true }
    } else {
        $script:Open = @{}
        $script:SearchOpen = @{}
        $script:ProgOpen = @{}
    }
    Render
}

$script:ProgOpen = @{}
$OnProgToggle = {
    param($s, $e)
    $e.Handled = $true
    $id = [string]$s.Tag
    if ($script:ProgOpen.ContainsKey($id)) { $script:ProgOpen.Remove($id) } else { $script:ProgOpen[$id] = $true }
    Render
}

function Update-ExpandLink {
    $closed = @($script:VisibleIds | Where-Object { -not (Test-IsOpen $_) })
    $BtnExpandAll.Text = if ($script:VisibleIds.Count -gt 0 -and $closed.Count -eq 0) { '全部收起' } else { '全部展开' }
    $BtnExpandAll.Visibility = if ($script:VisibleIds.Count -gt 0) { 'Visible' } else { 'Collapsed' }
}

# ---------- 项目与类别 ----------
# 每条待办（事项）和每条工作记录（进展）都归到一个项目或固定类别下
$script:Projects = New-Object System.Collections.ArrayList
$BuiltinCats = @(
    @{ id = 'cat-meeting'; name = '会议' },
    @{ id = 'cat-trip'; name = '出差' },
    @{ id = 'cat-study'; name = '学习' },
    @{ id = 'cat-daily'; name = '日常' },
    @{ id = 'cat-reimb'; name = '报销' }
)
$ReimbStates = @('待报', '已提交', '已到账')

function ConvertTo-Project($o) {
    $id = [string]$o.id
    if (-not $id) { $id = [guid]::NewGuid().ToString('N') }
    $kind = if ([string]$o.kind -eq 'category') { 'category' } else { 'project' }
    $st = if ([string]$o.status -eq 'closed') { 'closed' } else { 'active' }
    $cr = [string]$o.created; if (-not $cr) { $cr = (Get-Date).ToString('yyyy-MM-dd') }
    [pscustomobject]@{ id = $id; text = [string]$o.text; kind = $kind; status = $st; start = [string]$o.start; created = $cr; upd = [long]$o.upd }
}
# 项目名存在 text 字段里（存储层按 text 判断有效记录）

function Get-NameId([string]$name) {
    $md5 = [System.Security.Cryptography.MD5]::Create()
    $h = $md5.ComputeHash([Text.Encoding]::UTF8.GetBytes($name.Trim().ToLower()))
    return ('p-' + (($h[0..5] | ForEach-Object { $_.ToString('x2') }) -join ''))
}

function Find-Project([string]$id) {
    foreach ($p in $script:Projects) { if ($p.id -eq $id) { return $p } }
    return $null
}

function Find-ProjectByName([string]$name) {
    $n = $name.Trim()
    foreach ($p in $script:Projects) { if ($p.text -eq $n) { return $p } }
    return $null
}

# 按名字找项目，没有就新建（同名项目在两台电脑上生成的编号相同，同步时不会重复）
function Ensure-Project([string]$name) {
    $n = $name.Trim()
    if (-not $n) { return (Find-Project 'cat-daily') }
    $p = Find-ProjectByName $n
    if ($p) { return $p }
    $id = Get-NameId $n
    $p = Find-Project $id
    if ($p) { return $p }
    $p = ConvertTo-Project ([pscustomobject]@{ id = $id; text = $n; kind = 'project'; status = 'active'; start = (Get-Date).ToString('yyyy-MM-dd') })
    [void]$script:Projects.Add($p)
    return $p
}

function Ensure-BuiltinCats {
    foreach ($c in $BuiltinCats) {
        if (-not (Find-Project $c.id)) {
            [void]$script:Projects.Add((ConvertTo-Project ([pscustomobject]@{ id = $c.id; text = $c.name; kind = 'category'; status = 'active'; created = '2026-01-01' })))
        }
    }
}

function Get-ProjName([string]$id) {
    $p = Find-Project $id
    if ($p) { return $p.text }
    return '日常'
}

function Test-Reimb($it) { return ($it.project -eq 'cat-reimb') }

# 排序：进行中的项目（按名字）→ 已结束的项目 → 会议、出差、学习、日常、报销
function Get-ProjectsOrdered([bool]$withClosed = $true) {
    $act = @($script:Projects | Where-Object { $_.kind -eq 'project' -and $_.status -eq 'active' } | Sort-Object text)
    $cls = @($script:Projects | Where-Object { $_.kind -eq 'project' -and $_.status -eq 'closed' } | Sort-Object text)
    $cats = @(foreach ($c in $BuiltinCats) { Find-Project $c.id })
    if ($withClosed) { return @($act + $cls + $cats | Where-Object { $_ }) }
    return @($act + $cats | Where-Object { $_ })
}

function Save-Projects { Request-Save $script:ProjStore }

# 改项目名字：项目下的事项和进展都跟着走，各处显示一起变
function Rename-ProjectById([string]$id) {
    $p = Find-Project $id
    if (-not $p) { return }
    $owner = if ($script:CalWin -and $script:CalWin.IsActive) { $script:CalWin } else { $window }
    if ($p.kind -eq 'category') { [System.Windows.MessageBox]::Show($owner, '「' + $p.text + '」是固定类别，不能改名。', '桌面待办') | Out-Null; return }
    $r = Show-TextDialog '修改项目名称' $p.text '改名后，这个项目下的待办、进展和报告都跟着用新名字'
    if (-not $r -or -not $r.Trim() -or $r.Trim() -eq $p.text) { return }
    $other = Find-ProjectByName $r.Trim()
    if ($other -and $other.id -ne $p.id) { [System.Windows.MessageBox]::Show($owner, '已经有一个叫「' + $r.Trim() + '」的项目了，换个名字。', '桌面待办') | Out-Null; return }
    $p.text = $r.Trim()
    Save-Projects
    Render-Chips
    Render
}
$OnRenameProj = { param($s, $e) $e.Handled = $true; Rename-ProjectById ([string]$s.Tag) }

function Add-RenameLink($panel, [string]$projId) {
    $pj = Find-Project $projId
    if (-not $pj -or $pj.kind -ne 'project') { return }
    $l = New-LinkText ('修改项目名称（现在叫「' + $pj.text + '」）') $projId $OnRenameProj
    $l.FontSize = 12; $l.Margin = Th 0 4 0 0
    [void]$panel.Children.Add($l)
}

function Get-ProjOrderIndex([string]$id) {
    $list = Get-ProjectsOrdered $true
    for ($i = 0; $i -lt $list.Count; $i++) { if ($list[$i].id -eq $id) { return $i } }
    return 9999
}

function New-Project-Dialog {
    $r = Show-TextDialog '新建项目' '' '输入项目名称，比如"A项目""年度研讨会"'
    if (-not $r -or -not $r.Trim()) { return $null }
    $p = Ensure-Project $r.Trim()
    if ($p.status -eq 'closed') { $p.status = 'active' }
    Save-Projects
    return $p
}

# 旧数据迁移：旧的 #标签 变成项目；没有项目的待办和工作记录归到"日常"
function Invoke-Migration {
    Ensure-BuiltinCats
    if ($script:S.migratedV2) { Save-Projects; return }
    foreach ($n in @($script:S.projects)) { if ($n) { [void](Ensure-Project ([string]$n)) } }
    foreach ($w in $script:Work) {
        if (-not $w.proj) {
            if ($w.project) { $w.proj = (Ensure-Project $w.project).id }
        }
    }
    foreach ($it in $script:Items) {
        if (-not $it.project) {
            $from = @($script:Work | Where-Object { $_.src -eq $it.id -and $_.proj -and $_.proj -ne 'cat-daily' } | Select-Object -First 1)
            $it.project = if ($from.Count -gt 0) { $from[0].proj } else { 'cat-daily' }
        }
    }
    foreach ($w in $script:Work) {
        if ($w.src) {
            $t = Find-Todo $w.src
            if ($t) {
                if (-not $w.proj -or $w.proj -eq 'cat-daily') { $w.proj = $t.project }
                if (-not $t.start -or $w.date -lt $t.start) { $t.start = $w.date }
            }
        }
        if (-not $w.proj) { $w.proj = 'cat-daily' }
    }
    Save-Projects; Save-Items; Save-Work
    $script:S.migratedV2 = $true; Save-Settings
}

# 从待办文字末尾识别报销金额，比如"打车费 56""高铁票 553.5元"
function Split-Amount([string]$text) {
    $m = [regex]::Match($text, '^(.*?)[\s　]*(?:¥|￥)?\s*(\d+(?:\.\d{1,2})?)\s*元?\s*$')
    if ($m.Success -and $m.Groups[1].Value.Trim()) { return @($m.Groups[1].Value.Trim(), $m.Groups[2].Value) }
    return @($text, '')
}

function Format-Money([string]$a) {
    $v = 0.0
    if ([double]::TryParse($a, [ref]$v)) { return ('{0:N2}' -f $v) }
    return ''
}

# ---------- 进展（工作记录）与事项的关联 ----------
function Get-TodoProgress($it) { return @($script:Work | Where-Object { $_.src -eq $it.id } | Sort-Object date, time, seq) }

function Add-Progress($it, [string]$text, [string]$date) {
    if (-not $text.Trim()) { return }
    if (-not $date) { $date = (Get-Date).ToString('yyyy-MM-dd') }
    $tm = if ($date -eq (Get-Date).ToString('yyyy-MM-dd')) { (Get-Date).ToString('HH:mm') } else { '' }
    [void]$script:Work.Add((New-WorkEntry $text.Trim() $date $tm $it.project $it.id))
    if (-not $it.start -or $date -lt $it.start) { $it.start = $date; Save-Items }
    Save-Work
}

# 打勾具体事项：当天自动记一条"完成：xxx"；取消打勾：把这条自动记录撤掉
function Add-AutoProgress($it, $x) {
    $body = if ($x.kind -eq 'image') { '[图片]' } else { $x.text }
    $e = New-WorkEntry ('完成：' + $body) (Get-Date).ToString('yyyy-MM-dd') (Get-Date).ToString('HH:mm') $it.project $it.id
    $e.auto = $true; $e.subId = $x.id
    [void]$script:Work.Add($e)
    if (-not $it.start) { $it.start = (Get-Date).ToString('yyyy-MM-dd') }
    Save-Work
}

# 具体事项改了文字，自动记的那条"完成：xxx"跟着改（自己改过那条记录的就不动）
function Sync-AutoText($it, $x, [string]$old) {
    $hit = @($script:Work | Where-Object { $_.auto -and $_.subId -eq $x.id -and $_.src -eq $it.id -and $_.text -eq ('完成：' + $old) })
    foreach ($w in $hit) { $w.text = '完成：' + $x.text }
    if ($hit.Count -gt 0) { Save-Work }
}

# 显示用：自动记的"完成：xxx"显示成"✓ xxx"
function Get-WorkBody($w) {
    $b = [string]$w.text
    if ($w.auto -and $b.StartsWith('完成：')) { return ('✓ ' + $b.Substring(3)) }
    return $b
}
# 周/月视图里一条工作记录怎么显示：挂在某件待办下的，前面带上待办的名字；和名字一样的不重复
function Get-WorkShowText($w, $src) {
    $b = Get-WorkBody $w
    if (-not $src) { return $b }
    $t = [string]$src.text
    if ($b.Trim() -eq $t.Trim() -or $b.StartsWith($t)) { return $b }
    return ($t + '：' + $b)
}

# 以前勾掉、当时没自动记进展的具体事项，补记一条（只补一次）
function Invoke-SubBackfill {
    if ($script:S.subBackfill) { return }
    $n = 0
    foreach ($it in @($script:Items)) {
        foreach ($x in @($it.subs)) {
            if (-not $x.done) { continue }
            if (@($script:Work | Where-Object { $_.auto -and $_.subId -eq $x.id }).Count -gt 0) { continue }
            $d = Parse-Dt $it.doneAt
            $tm = if ($d) { $d.ToString('HH:mm') } else { '' }
            if (-not $d) { $last = @(Get-TodoProgress $it | Select-Object -Last 1); if ($last.Count -gt 0) { $d = Parse-Day $last[0].date } }
            if (-not $d) { $d = Parse-Dt $it.created }
            if (-not $d) { $d = Get-Date }
            $body = if ($x.kind -eq 'image') { '[图片]' } else { $x.text }
            $e = New-WorkEntry ('完成：' + $body) $d.ToString('yyyy-MM-dd') $tm (Get-TodoProj $it) $it.id
            $e.id = 'auto-' + $it.id + '-' + $x.id   # 两台电脑各补一遍也是同一条，不会重复
            $e.auto = $true; $e.subId = $x.id
            if (-not (Find-Work $e.id)) { [void]$script:Work.Add($e); $n++ }
        }
    }
    if ($n -gt 0) { Save-Work }
    $script:S.subBackfill = $true; Save-Settings
}

function Remove-AutoProgress($it, $x) {
    $hit = @($script:Work | Where-Object { $_.auto -and $_.subId -eq $x.id -and $_.src -eq $it.id })
    foreach ($w in $hit) { [void]$script:Work.Remove($w) }
    if ($hit.Count -gt 0) { Save-Work }
}

$OnAddProgress = {
    param($s, $e)
    $e.Handled = $true
    $it = Find-Todo ([string]$s.Tag); if (-not $it) { return }
    $r = Show-TextDialog ('记进展：' + $it.text) '' '今天在这件事上做了什么？会记到今天的工作记录里，归到这件待办和它的项目下。' $true
    if ($r -and $r.Trim()) { Add-Progress $it $r ''; Render; Notify '已记下今天的进展' $r.Trim() }
}

# ---------- 选项目的菜单（输入区、右键菜单共用） ----------
function New-ProjectMenu([scriptblock]$onPick, [string]$current, [string]$tagPrefix) {
    $cm = [System.Windows.Controls.ContextMenu]::new()
    $acts = @(Get-ProjectsOrdered $false)
    $hasProj = $false
    foreach ($p in $acts) {
        if ($p.kind -eq 'category' -and $hasProj) { [void]$cm.Items.Add(([System.Windows.Controls.Separator]::new())); $hasProj = $false }
        if ($p.kind -eq 'project') { $hasProj = $true }
        $label = $p.text
        if ($p.id -eq $current) { $label = '✓ ' + $label }
        [void]$cm.Items.Add((New-MenuItem $label ($tagPrefix + $p.id) $onPick))
    }
    [void]$cm.Items.Add(([System.Windows.Controls.Separator]::new()))
    [void]$cm.Items.Add((New-MenuItem '新建项目…' ($tagPrefix + '__new') $onPick))
    [void]$cm.Items.Add((New-MenuItem '管理项目…' '' { Show-Calendar 'project' }))
    return $cm
}

function Open-MenuAt($menu, $target) {
    $menu.PlacementTarget = $target
    $menu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
    $menu.IsOpen = $true
}

function Resolve-PickedProject([string]$prj) {
    if ($prj -eq '__new') { $p = New-Project-Dialog; if ($p) { return $p.id }; return $null }
    return $prj
}

# 待办输入区：选项目
$OnTodoProjPick = {
    param($s, $e)
    $prj = Resolve-PickedProject ([string]$s.Tag)
    if (-not $prj) { return }
    $script:Pending.proj = $prj
    $script:S.lastProj = $prj; Save-Settings
    Render-Chips; [void]$InputBox.Focus()
}
$OnTodoProjChip = { param($s, $e) Open-MenuAt (New-ProjectMenu $OnTodoProjPick ([string]$script:Pending.proj) '') $s }

# 工作记录输入区：选项目、选关联事项
$OnLogProjPick = {
    param($s, $e)
    $prj = Resolve-PickedProject ([string]$s.Tag)
    if (-not $prj) { return }
    $script:LogPending.proj = $prj
    if ($script:LogPending.todo) { $t = Find-Todo $script:LogPending.todo; if ($t -and $t.project -ne $prj) { $script:LogPending.todo = '' } }
    $script:S.lastLogProj = $prj; Save-Settings
    Render-Chips; [void]$InputBox.Focus()
}
$OnLogProjChip = { param($s, $e) Open-MenuAt (New-ProjectMenu $OnLogProjPick ([string]$script:LogPending.proj) '') $s }
$OnLogTodoPick = {
    param($s, $e)
    $id = [string]$s.Tag
    if ($id -eq '__none') { $script:LogPending.todo = '' }
    else { $script:LogPending.todo = $id; $t = Find-Todo $id; if ($t) { $script:LogPending.proj = $t.project } }
    Render-Chips; [void]$InputBox.Focus()
}
$OnLogTodoChip = {
    param($s, $e)
    $cm = [System.Windows.Controls.ContextMenu]::new()
    [void]$cm.Items.Add((New-MenuItem '不属于哪件待办' '__none' $OnLogTodoPick))
    [void]$cm.Items.Add(([System.Windows.Controls.Separator]::new()))
    $prj = [string]$script:LogPending.proj
    $list = @($script:Items | Where-Object { -not $_.done -and (-not $prj -or $_.project -eq $prj) } | Sort-Object created -Descending | Select-Object -First 25)
    if ($list.Count -eq 0) { $mi = New-MenuItem '（这个项目下没有没做完的待办）' '' $null; $mi.IsEnabled = $false; [void]$cm.Items.Add($mi) }
    foreach ($t in $list) {
        $label = $t.text; if ($label.Length -gt 30) { $label = $label.Substring(0, 30) + '…' }
        if ($t.id -eq $script:LogPending.todo) { $label = '✓ ' + $label }
        [void]$cm.Items.Add((New-MenuItem $label $t.id $OnLogTodoPick))
    }
    Open-MenuAt $cm $s
}

# 右键菜单：把一条待办 / 工作记录换到别的项目
$OnTodoSetProj = {
    param($s, $e)
    $p = ([string]$s.Tag).Split([char[]]'|', 2)
    $it = Find-Todo $p[0]; if (-not $it) { return }
    $prj = Resolve-PickedProject $p[1]; if (-not $prj) { return }
    $it.project = $prj
    foreach ($w in @($script:Work | Where-Object { $_.src -eq $it.id })) { $w.proj = $prj }
    if ($prj -eq 'cat-reimb' -and -not $it.rstate) { $it.rstate = '待报' }
    Save-Items; Save-Work; Render
}
$OnWorkSetProj = {
    param($s, $e)
    $p = ([string]$s.Tag).Split([char[]]'|', 2)
    $w = Find-Work $p[0]; if (-not $w) { return }
    $prj = Resolve-PickedProject $p[1]; if (-not $prj) { return }
    $w.proj = $prj
    if ($w.src) { $t = Find-Todo $w.src; if ($t -and $t.project -ne $prj) { $w.src = '' } }
    Save-Work; Render
}
$OnWorkSetTodo = {
    param($s, $e)
    $p = ([string]$s.Tag).Split([char[]]'|', 2)
    $w = Find-Work $p[0]; if (-not $w) { return }
    if ($p[1] -eq '__none') { $w.src = '' } else { $w.src = $p[1]; $t = Find-Todo $p[1]; if ($t) { $w.proj = $t.project } }
    Save-Work; Render
}

function Add-ProjSubmenu($menu, [string]$id, [string]$current, [scriptblock]$handler) {
    $pm = New-MenuItem ('项目：' + (Get-ProjName $current)) '' $null
    foreach ($p in @(Get-ProjectsOrdered $false)) {
        $label = if ($p.id -eq $current) { '✓ ' + $p.text } else { $p.text }
        [void]$pm.Items.Add((New-MenuItem $label ($id + '|' + $p.id) $handler))
    }
    [void]$pm.Items.Add(([System.Windows.Controls.Separator]::new()))
    [void]$pm.Items.Add((New-MenuItem '新建项目…' ($id + '|__new') $handler))
    $cur = Find-Project $current
    if ($cur -and $cur.kind -eq 'project') { [void]$pm.Items.Add((New-MenuItem ('修改「' + $cur.text + '」的名称…') $cur.id $OnRenameProj)) }
    [void]$menu.Items.Add($pm)
}

# 开始日期、报销金额和状态
$OnTodoStart = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|')
    $it = Find-Todo $p[0]; if (-not $it) { return }
    if ($p[1] -eq 'clear') { $it.start = '' }
    else {
        $init = Parse-Day $it.start; if (-not $init) { $init = (Get-Date).Date }
        $d = Show-DatePickerDialog $init.AddHours(9) '这件事从哪天开始'
        if (-not $d) { return }
        $it.start = $d.ToString('yyyy-MM-dd')
    }
    Save-Items; Render
}
$OnReimbAmount = {
    param($s, $e)
    $it = Find-Todo ([string]$s.Tag); if (-not $it) { return }
    $r = Show-TextDialog '报销金额' ([string]$it.amount) '填数字，单位元，比如 56 或 553.5'
    if ($null -eq $r) { return }
    $v = 0.0
    if ($r.Trim() -and -not [double]::TryParse($r.Trim(), [ref]$v)) { [System.Windows.MessageBox]::Show('金额要填数字。', '桌面待办') | Out-Null; return }
    $it.amount = $r.Trim()
    Save-Items; Render
}
$OnReimbState = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|')
    $it = Find-Todo $p[0]; if (-not $it) { return }
    $it.rstate = $p[1]
    if ($p[1] -eq '已到账' -and -not $it.done) { $it.done = $true; $it.doneAt = (Get-Date).ToString($Fmt) }
    Save-Items; Render
}

# ---------- 通用修改：时间、日期、进展、具体事项 ----------
# 修改一条待办的记录时间 / 完成时间
function Edit-TodoStamp([string]$id, [string]$field) {
    $it = Find-Todo $id
    if (-not $it) { return }
    $cur = Parse-Dt ([string]$it.$field)
    if (-not $cur) { $cur = Get-Date }
    $title = if ($field -eq 'doneAt') { '修改完成时间' } else { '修改记录时间' }
    $d = Show-DatePickerDialog $cur $title
    if (-not $d) { return }
    $it.$field = $d.ToString($Fmt)
    Save-Items
    Render
}
$OnEditCreated = { param($s, $e) Edit-TodoStamp ([string]$s.Tag) 'created' }
$OnEditDoneAt  = { param($s, $e) Edit-TodoStamp ([string]$s.Tag) 'doneAt' }

# 修改已有的一个提醒 / 截止时间
$OnEditDue = { param($s, $e) $q = ([string]$s.Tag).Split([char[]]'|', 2); Set-Due $q[0] 'edit' $q[1] }

# 进展（工作记录）的右键菜单：改内容、改日期时间、删除；小窗、工作台都用这一套
$OnWorkMenuEdit = { param($s, $e) Edit-Work ([string]$s.Tag) }
$OnWorkMenuDate = { param($s, $e) Set-WorkDate ([string]$s.Tag) }
$OnWorkMenuDel  = { param($s, $e) Remove-Work ([string]$s.Tag) }
function New-WorkLineMenu([string]$wid) {
    $m = [System.Windows.Controls.ContextMenu]::new()
    [void]$m.Items.Add((New-MenuItem '修改内容…' $wid $OnWorkMenuEdit))
    [void]$m.Items.Add((New-MenuItem '修改日期和时间…' $wid $OnWorkMenuDate))
    [void]$m.Items.Add(([System.Windows.Controls.Separator]::new()))
    [void]$m.Items.Add((New-MenuItem '删除这条进展' $wid $OnWorkMenuDel))
    return $m
}

# 具体事项的右键菜单（用在今日工作里展开的具体事项上）
$OnSubMenuEdit = {
    param($s, $e)
    $q = ([string]$s.Tag).Split('|'); $it = Find-Todo $q[0]; if (-not $it) { return }
    $x = Find-Sub $it $q[1]; if (-not $x -or $x.kind -eq 'image') { return }
    $r = Show-TextDialog '修改具体事项' $x.text '' $true
    if ($r -and $r.Trim()) { $old = $x.text; $x.text = $r.Trim(); Save-Items; Sync-AutoText $it $x $old; Render }
}
$OnSubMenuToggle = {
    param($s, $e)
    $q = ([string]$s.Tag).Split('|'); $it = Find-Todo $q[0]; if (-not $it) { return }
    $x = Find-Sub $it $q[1]; if (-not $x) { return }
    $x.done = -not $x.done
    Save-Items
    if ($x.done) { Add-AutoProgress $it $x } else { Remove-AutoProgress $it $x }
    Render
}
function New-SubLineMenu($it, $x) {
    $tag = $it.id + '|' + $x.id
    $m = [System.Windows.Controls.ContextMenu]::new()
    if ($x.kind -ne 'image') { [void]$m.Items.Add((New-MenuItem '修改这一项…' $tag $OnSubMenuEdit)) }
    [void]$m.Items.Add((New-MenuItem $(if ($x.done) { '改回未完成' } else { '标为已完成（会在今天自动记一条进展）' }) $tag $OnSubMenuToggle))
    return $m
}

# ---------- 双击直接修改（所见即改） ----------
# 用法：显示文字的只读框调用 Enable-Inline，Tag 写明改的是哪条：
#   todo|待办id   sub|待办id|事项id   work|记录id   proj|项目id
function Apply-InlineValue([string]$tag, [string]$text) {
    $t = $text.Trim()
    if (-not $t) { return $false }
    $q = $tag.Split('|')
    switch ($q[0]) {
        'todo' { $it = Find-Todo $q[1]; if ($it -and $it.text -ne $t) { $it.text = $t; Save-Items; return $true } }
        'sub'  { $it = Find-Todo $q[1]; if ($it) { $x = Find-Sub $it $q[2]; if ($x -and $x.kind -ne 'image' -and $x.text -ne $t) { $old = $x.text; $x.text = $t; Save-Items; Sync-AutoText $it $x $old; return $true } } }
        'work' { $w = Find-Work $q[1]; if ($w -and $w.text -ne $t) { $w.text = $t; Save-Work; return $true } }
        'proj' {
            $p = Find-Project $q[1]
            if ($p -and $p.kind -eq 'project' -and $p.text -ne $t) {
                $other = Find-ProjectByName $t
                if ($other -and $other.id -ne $p.id) { [System.Windows.MessageBox]::Show('已经有一个叫「' + $t + '」的项目了，换个名字。', '桌面待办') | Out-Null; return $false }
                $p.text = $t; Save-Projects; return $true
            }
        }
    }
    return $false
}

function Get-InlineRaw([string]$tag) {
    $q = $tag.Split('|')
    switch ($q[0]) {
        'todo' { $it = Find-Todo $q[1]; if ($it) { return $it.text } }
        'sub'  { $it = Find-Todo $q[1]; if ($it) { $x = Find-Sub $it $q[2]; if ($x -and $x.kind -ne 'image') { return $x.text } } }
        'work' { $w = Find-Work $q[1]; if ($w) { return $w.text } }
        'proj' { $p = Find-Project $q[1]; if ($p) { return $p.text } }
    }
    return $null
}

function Start-InlineBox($tb) {
    $raw = Get-InlineRaw ([string]$tb.Tag)
    if ($null -eq $raw) { return }
    $tb.Resources['inlineCancel'] = $false
    $tb.Text = $raw
    $tb.IsReadOnly = $false
    $tb.AcceptsReturn = $true
    $tb.IsReadOnlyCaretVisible = $true
    $tb.BorderThickness = Th 1 1 1 1
    $tb.Padding = Th 3 2 3 2
    $tb.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'InputBgBrush')
    $tb.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'AccentBrush')
    $tb.SetResourceReference([System.Windows.Controls.TextBox]::CaretBrushProperty, 'TextBrush')
    [void]$tb.Focus()
    $tb.SelectAll()
}

$OnInlineDbl = {
    param($s, $e)
    if (-not $s.IsReadOnly) { return }
    $e.Handled = $true
    Start-InlineBox $s
}
$OnInlineKey = {
    param($s, $e)
    if ($s.IsReadOnly) { return }
    if ($e.Key -eq [System.Windows.Input.Key]::Return -and -not ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift)) {
        $e.Handled = $true; [System.Windows.Input.Keyboard]::ClearFocus()
    } elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) {
        $e.Handled = $true; $s.Resources['inlineCancel'] = $true; [System.Windows.Input.Keyboard]::ClearFocus()
    }
}
$OnInlineLost = {
    param($s, $e)
    if ($e.NewFocus -is [System.Windows.Controls.ContextMenu] -or $e.NewFocus -is [System.Windows.Controls.MenuItem]) { return }
    if ($s.IsReadOnly) { return }
    $s.IsReadOnly = $true
    $cancel = [bool]$s.Resources['inlineCancel']
    $changed = $false
    if (-not $cancel) { $changed = Apply-InlineValue ([string]$s.Tag) $s.Text }
    Render
    if ($changed -and $script:CalWin) { Render-Cal }
}

function Enable-Inline($tb, [string]$tag) {
    $tb.Tag = $tag
    $tip = '双击直接修改，回车保存，Esc 取消'
    $tb.ToolTip = if ($tb.ToolTip) { [string]$tb.ToolTip + '；' + $tip } else { $tip }
    $tb.add_PreviewMouseDoubleClick($OnInlineDbl)
    $tb.add_PreviewKeyDown($OnInlineKey)
    $tb.add_LostKeyboardFocus($OnInlineLost)
    return $tb
}

# 显示用的 TextBlock 双击后换成输入框（周/月视图、报告里用）
$OnSwapKey = {
    param($s, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return -and -not ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Shift)) { $e.Handled = $true; [System.Windows.Input.Keyboard]::ClearFocus() }
    elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) { $e.Handled = $true; $s.Resources['inlineCancel'] = $true; [System.Windows.Input.Keyboard]::ClearFocus() }
}
$OnSwapLost = {
    param($s, $e)
    if ($e.NewFocus -is [System.Windows.Controls.ContextMenu] -or $e.NewFocus -is [System.Windows.Controls.MenuItem]) { return }
    if ($s.Resources['inlineDone']) { return }
    $s.Resources['inlineDone'] = $true
    $script:CalInline = $false
    if (-not [bool]$s.Resources['inlineCancel']) { [void](Apply-InlineValue ([string]$s.Tag) $s.Text) }
    Render
}

function Start-SwapEdit($textBlock, [string]$tag) {
    $raw = Get-InlineRaw $tag
    if ($null -eq $raw) { return }
    $parent = $textBlock.Parent
    if (-not $parent) { return }
    $box = New-EditBox $raw $false
    $box.AcceptsReturn = $true
    $box.Tag = $tag
    $box.FontSize = $textBlock.FontSize
    $box.Resources['inlineCancel'] = $false
    $box.add_PreviewKeyDown($OnSwapKey)
    $box.add_LostKeyboardFocus($OnSwapLost)
    $box.add_MouseLeftButtonUp({ param($s, $e) $e.Handled = $true })
    $box.add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true })
    if ($parent -is [System.Windows.Controls.Panel]) {
        $i = $parent.Children.IndexOf($textBlock)
        $parent.Children.RemoveAt($i)
        $parent.Children.Insert($i, $box)
    } elseif ($parent -is [System.Windows.Controls.Border]) {
        $parent.Child = $box
    } else { return }
    $script:CalInline = $true
    [void]$box.Focus()
    $box.SelectAll()
}

# 双击时打开所在卡片的右键菜单（用在"#项目 截止…"这类说明行上）
function Open-AncestorMenu($el) {
    $cur = $el
    while ($cur) {
        if ($cur -is [System.Windows.FrameworkElement] -and $cur.ContextMenu -and -not ($cur -is [System.Windows.Controls.TextBox])) {
            Fill-LazyMenu $cur.ContextMenu
            $cur.ContextMenu.PlacementTarget = $cur
            $cur.ContextMenu.Placement = [System.Windows.Controls.Primitives.PlacementMode]::MousePoint
            $cur.ContextMenu.IsOpen = $true
            return
        }
        $cur = [System.Windows.Media.VisualTreeHelper]::GetParent($cur)
    }
}
$OnMetaDbl = {
    param($s, $e)
    if ($e.ClickCount -ne 2) { return }
    $e.Handled = $true
    Open-AncestorMenu $s
}

$OnProjLabelDbl = {
    param($s, $e)
    if ($e.ClickCount -ne 2) { return }
    $e.Handled = $true
    Rename-ProjectById ([string]$s.Tag)
}
$OnWorkDateDbl = {
    param($s, $e)
    if ($e.ClickCount -ne 2) { return }
    $e.Handled = $true
    Set-WorkDate ([string]$s.Tag)
}

# ---------- 今日工作 ----------
$script:Work = New-Object System.Collections.ArrayList
$script:ViewDate = (Get-Date).Date
$script:LogPending = @{ time = ''; proj = ''; todo = '' }

function Parse-Day([string]$s) {
    $d = [datetime]::MinValue
    if ([datetime]::TryParseExact($s, 'yyyy-MM-dd', $Inv, [Globalization.DateTimeStyles]::None, [ref]$d)) { return $d }
    return $null
}

function Format-Day([datetime]$d) {
    $s = $d.ToString('M月d日') + ' ' + $WeekNames[[int]$d.DayOfWeek]
    $diff = ($d.Date - (Get-Date).Date).Days
    if ($diff -eq 0) { $s = '今天 · ' + $s } elseif ($diff -eq -1) { $s = '昨天 · ' + $s } elseif ($diff -eq 1) { $s = '明天 · ' + $s }
    return $s
}

# 这一天完成的待办：直接显示在工作记录里（不再复制一份）；以前导入过同名记录的不重复显示
function Get-DoneRows([string]$ds, $entries) {
    $same = @{}
    foreach ($w in $entries) { if ($w.src) { $same[$w.src + '|' + ([string]$w.text).Trim()] = $true } }
    return @($script:Items | Where-Object { $_.done -and ([string]$_.doneAt).StartsWith($ds) -and -not $same.ContainsKey($_.id + '|' + ([string]$_.text).Trim()) })
}
function Get-DoneTime($it) { $s0 = [string]$it.doneAt; if ($s0.Length -ge 16) { return $s0.Substring(11, 5) }; return '' }
# 工作记录 + 完成的待办，按时间排好
function Get-DayRows([string]$ds) {
    $entries = @($script:Work | Where-Object { $_.date -eq $ds })
    $rows = @()
    foreach ($w in $entries) { $rows += [pscustomobject]@{ tm = [string]$w.time; seq = [string]$w.seq; w = $w; t = $null } }
    foreach ($it in (Get-DoneRows $ds $entries)) { $rows += [pscustomobject]@{ tm = (Get-DoneTime $it); seq = ''; w = $null; t = $it } }
    return @($rows | Sort-Object tm, seq)
}

function New-DoneLogCard($it, [int]$no = 0) {
    $card = [System.Windows.Controls.Border]::new()
    $card.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $card.Padding = Th 8 7 8 7
    $card.Margin = Th 0 0 0 6
    $card.Background = TB 'Card'
    $card.BorderBrush = TB 'CardLine'
    $card.BorderThickness = Th 1 1 1 1
    $card.Opacity = 0.85
    $card.Tag = $it.id
    $dp = [System.Windows.Controls.DockPanel]::new()
    $tmv = Get-DoneTime $it
    if ($tmv -and (Test-ShowTime)) {
        $tm = [System.Windows.Controls.TextBlock]::new()
        $tm.Text = $tmv; $tm.Margin = Th 0 0 10 0; $tm.VerticalAlignment = 'Top'
        $tm.FontWeight = [System.Windows.FontWeights]::SemiBold
        $tm.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
        [System.Windows.Controls.DockPanel]::SetDock($tm, 'Left')
        [void]$dp.Children.Add($tm)
    }
    $sp = [System.Windows.Controls.StackPanel]::new()
    $row = [System.Windows.Controls.DockPanel]::new()
    $ck = [System.Windows.Controls.TextBlock]::new(); $ck.Text = '✓ '; $ck.FontSize = 13.5
    $ck.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
    [System.Windows.Controls.DockPanel]::SetDock($ck, 'Left')
    $tb = New-SelectableText $it.text 13.5
    [void](Enable-Inline $tb ('todo|' + $it.id))
    $tb.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'SubBrush')
    [void]$row.Children.Add($ck); [void]$row.Children.Add($tb)
    [void]$sp.Children.Add((New-NumTitle $no $row))
    $meta = [System.Windows.Controls.TextBlock]::new()
    $meta.FontSize = 11; $meta.Margin = Th 0 3 0 0; $meta.TextWrapping = 'Wrap'
    $meta.Text = '完成的待办 · ' + (Get-ProjName (Get-TodoProj $it)) + $(if ($it.copied) { ' · 已复制到表格' } else { '' })
    $meta.ToolTip = '这天打勾完成的待办，会一起复制到表格。双击这一行可以改完成时间、换项目、取消完成'
    $meta.add_MouseLeftButtonDown($OnMetaDbl)
    $meta.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
    [void]$sp.Children.Add($meta)
    $subs = @($it.subs)
    if ($subs.Count -gt 0) {
        $key = 'd:' + $it.id
        $open = $script:LogOpen.ContainsKey($key)
        $tg = [System.Windows.Controls.TextBlock]::new()
        $tg.Text = $(if ($open) { '▾ ' } else { '▸ ' }) + (Get-DetailSummary $it)
        $tg.FontSize = 11; $tg.Margin = Th 0 3 0 0; $tg.Cursor = $Hand; $tg.Tag = $key
        $tg.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
        $tg.add_MouseLeftButtonDown($OnLogToggle)
        [void]$sp.Children.Add($tg)
        if ($open) {
            $lines = Get-SubLines $it
            for ($i = 0; $i -lt $subs.Count; $i++) {
                $sl = New-SelectableText $lines[$i] 12
                if ($subs[$i].kind -ne 'image') { [void](Enable-Inline $sl ('sub|' + $it.id + '|' + $subs[$i].id)) }
                $sl.Margin = Th 12 1 0 0
                $sl.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'SubBrush')
                [void]$sp.Children.Add($sl)
            }
        }
    }
    [void]$dp.Children.Add($sp)
    $card.Child = $dp
    Set-LazyMenu $card { param($id) $t0 = Find-Todo $id; if ($t0) { New-TodoMenu $t0 } } $it.id $tb
    return $card
}

function ConvertTo-WorkEntry($o) {
    $id = [string]$o.id
    if (-not $id) { $id = [guid]::NewGuid().ToString('N') }
    [pscustomobject]@{
        id = $id; date = [string]$o.date; time = [string]$o.time; text = [string]$o.text
        project = [string]$o.project; src = [string]$o.src; copied = [bool]$o.copied; seq = [string]$o.seq
        upd = [long]$o.upd; proj = [string]$o.proj; auto = [bool]$o.auto; subId = [string]$o.subId
    }
}

function New-WorkEntry([string]$text, [string]$date, [string]$time, [string]$proj, [string]$src) {
    if (-not $proj) { $proj = 'cat-daily' }
    [pscustomobject]@{
        id = [guid]::NewGuid().ToString('N'); date = $date; time = $time; text = $text
        project = ''; src = $src; copied = $false; seq = [string](Get-Date).Ticks; upd = [long]0
        proj = $proj; auto = $false; subId = ''
    }
}

function Load-Work { [void](Load-Store $script:WorkStore) }

function Save-Work { Request-Save $script:WorkStore }

function Find-Work([string]$id) {
    foreach ($w in $script:Work) { if ($w.id -eq $id) { return $w } }
    return $null
}

function Remember-Project([string]$p) {
    if (-not $p) { return }
    $list = New-Object System.Collections.Generic.List[string]
    $list.Add($p)
    foreach ($x in @($script:S.projects)) {
        if ($x -and ([string]$x) -ne $p -and $list.Count -lt 8) { $list.Add([string]$x) }
    }
    $script:S.projects = $list.ToArray()
    Save-Settings
}

function Show-TimeDialog([string]$init) {
    $now = Get-Date
    $h = $now.Hour; $mi = $now.Minute
    if ($init -match '^(\d{1,2}):(\d{2})$') { $h = [int]$Matches[1]; $mi = [int]$Matches[2] }
    $script:Dlg = New-DialogWindow '选时间'
    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.Margin = Th 16 14 16 14
    $row = [System.Windows.Controls.StackPanel]::new()
    $row.Orientation = 'Horizontal'
    $lab = [System.Windows.Controls.TextBlock]::new()
    $lab.Text = '时间'; $lab.VerticalAlignment = 'Center'; $lab.Margin = Th 0 0 8 0
    $script:DlgHour = New-TimeCombo 23 $h
    $colon = [System.Windows.Controls.TextBlock]::new()
    $colon.Text = ':'; $colon.VerticalAlignment = 'Center'; $colon.Margin = Th 5 0 5 0
    $script:DlgMin = New-TimeCombo 59 $mi
    [void]$row.Children.Add($lab); [void]$row.Children.Add($script:DlgHour)
    [void]$row.Children.Add($colon); [void]$row.Children.Add($script:DlgMin)
    [void]$sp.Children.Add($row)

    $quick = [System.Windows.Controls.WrapPanel]::new()
    $quick.Margin = Th 0 8 0 0
    foreach ($q in '现在', '09:00', '10:00', '14:00', '16:00') {
        $t = [System.Windows.Controls.TextBlock]::new()
        $t.Text = $q; $t.Tag = $q
        $t.Foreground = Brush '#2F6FED'; $t.Cursor = $Hand; $t.Margin = Th 0 0 12 0
        $t.add_MouseLeftButtonUp($OnQuickTime)
        [void]$quick.Children.Add($t)
    }
    [void]$sp.Children.Add($quick)

    $btns = New-DialogButtons
    [void]$sp.Children.Add($btns[0])
    $script:DlgResult = $null
    $btns[1].add_Click({
        $hh = 0; $mm = 0
        $okH = [int]::TryParse(([string]$script:DlgHour.Text).Trim(), [ref]$hh) -and $hh -ge 0 -and $hh -le 23
        $okM = [int]::TryParse(([string]$script:DlgMin.Text).Trim(), [ref]$mm) -and $mm -ge 0 -and $mm -le 59
        if (-not ($okH -and $okM)) {
            [System.Windows.MessageBox]::Show($script:Dlg, '时间填得不对：小时填 0～23，分钟填 0～59。', '桌面待办') | Out-Null
            return
        }
        $script:DlgResult = '{0:00}:{1:00}' -f $hh, $mm
        $script:Dlg.DialogResult = $true
    })
    $script:Dlg.Content = $sp
    [void]$script:Dlg.ShowDialog()
    return $script:DlgResult
}

function Add-Log {
    $text = $InputBox.Text.Trim()
    if (-not $text) { return }
    $proj = [string]$script:LogPending.proj
    $m = [regex]::Match($text, '(?:^|\s)[#＃](\S+)')
    if ($m.Success) {
        $proj = (Ensure-Project $m.Groups[1].Value).id
        $text = $text.Remove($m.Index, $m.Length).Trim()
        Save-Projects
    }
    if (-not $text) { return }
    $time = [string]$script:LogPending.time
    if (-not $time) { $time = (Get-Date).ToString('HH:mm') }
    $src = [string]$script:LogPending.todo
    $vd = $script:ViewDate.ToString('yyyy-MM-dd')
    if ($src) { $t = Find-Todo $src; if ($t) { $proj = $t.project; if (-not $t.start -or $vd -lt $t.start) { $t.start = $vd; Save-Items } } else { $src = '' } }
    $entry = New-WorkEntry $text ($script:ViewDate.ToString('yyyy-MM-dd')) $time $proj $src
    [void]$script:Work.Add($entry)
    Save-Work
    $script:LogPending.time = ''
    $InputBox.Clear()
    Render-Chips
    Render
}

function Edit-Work([string]$id) {
    $w = Find-Work $id
    if (-not $w) { return }
    $r = Show-TextDialog '修改' $w.text
    if ($r -and $r.Trim()) { $w.text = $r.Trim(); Save-Work; Render }
}

function Remove-Work([string]$id) {
    $w = Find-Work $id
    if (-not $w) { return }
    [void]$script:Work.Remove($w)
    Save-Work
    Render
    Show-Undo ('已删除记录「' + (Short-Text $w.text) + '」') @{ kind = 'work'; item = $w }
}

function Set-WorkTime([string]$id) {
    $w = Find-Work $id
    if (-not $w) { return }
    $r = Show-TimeDialog $w.time
    if ($r) { $w.time = $r; Save-Work; Render }
}

function Set-WorkProject([string]$id, [string]$p) {
    $w = Find-Work $id
    if (-not $w) { return }
    if ($p -eq '__clear') { $p = '' }
    elseif ($p -eq '__new') {
        $r = Show-TextDialog '设置项目' $w.project '输入项目名称'
        if (-not $r -or -not $r.Trim()) { return }
        $p = $r.Trim()
    }
    $w.project = $p
    Save-Work
    if ($p) { Remember-Project $p }
    Render
}

function Get-ImportCandidates([string]$ds) {
    $have = @{}
    foreach ($w in $script:Work) { if ($w.src) { $have[$w.src] = $true } }
    $script:Items | Where-Object { $_.done -and ([string]$_.doneAt).StartsWith($ds) -and -not $have.ContainsKey($_.id) }
}

function Import-DoneTodos {
    $ds = $script:ViewDate.ToString('yyyy-MM-dd')
    $c = @(Get-ImportCandidates $ds)
    if ($c.Count -eq 0) { return }
    foreach ($it in $c) {
        $tm = ''
        if (([string]$it.doneAt).Length -ge 16) { $tm = ([string]$it.doneAt).Substring(11, 5) }
        [void]$script:Work.Add((New-WorkEntry $it.text $ds $tm $it.project $it.id))
    }
    Save-Work
    Render
    Notify '已导入' ('把当天完成的 {0} 条待办加进了工作记录，时间按打勾的时间算，可以右键改。' -f $c.Count)
}

function Get-Cell($w, [string]$col) {
    $d = Parse-Day $w.date
    switch ($col.Trim()) {
        '日期'       { return $w.date }
        '日期(月日)' { if ($d) { return $d.ToString('M月d日') }; return $w.date }
        '日期（月日）' { if ($d) { return $d.ToString('M月d日') }; return $w.date }
        '星期'       { if ($d) { return $WeekNames[[int]$d.DayOfWeek] }; return '' }
        '时间'       { return $w.time }
        '工作内容'   { return (Get-WorkFullText $w) }
        '内容'       { return (Get-WorkFullText $w) }
        '项目'       { return (Get-ProjName $w.proj) }
        { $_ -eq '事项' -or $_ -eq '待办' } { $t = if ($w.src) { Find-Todo $w.src } else { $null }; if ($t) { return $t.text }; return '' }
    }
    return ''
}

function Get-WorkFullText($w) {
    if (-not $w.src) { return $w.text }
    $subs = Get-SubLines (Find-Todo $w.src)
    if ($subs.Count -eq 0) { return $w.text }
    return ($w.text + '（' + ($subs -join '；') + '）')
}

function Copy-LogToSheet {
    $ds = $script:ViewDate.ToString('yyyy-MM-dd')
    # 工作记录和这天完成的待办一起复制；完成的待办临时拼成一条记录的样子
    $entries = @(Get-DayRows $ds | ForEach-Object {
        if ($_.w) { $_.w }
        else { [pscustomobject]@{ date = $ds; time = $_.tm; text = $_.t.text; proj = (Get-TodoProj $_.t); src = $_.t.id; copied = [bool]$_.t.copied; todo = $_.t } }
    })
    if ($entries.Count -eq 0) { Notify '没有可复制的' '这一天还没有工作记录。'; return }
    $targets = @($entries | Where-Object { -not $_.copied })
    $again = $false
    if ($targets.Count -eq 0) { $targets = $entries; $again = $true }
    $cols = ([string]$script:S.logCols) -split '[,，]'
    $rows = foreach ($w in $targets) {
        $cells = foreach ($c in $cols) { ((Get-Cell $w $c) -replace "[\t\r\n]+", ' ') }
        $cells -join "`t"
    }
    try { [System.Windows.Clipboard]::SetText((@($rows) -join "`r`n")) }
    catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。'; return }
    foreach ($w in $targets) { if ($w.PSObject.Properties['todo']) { $w.todo.copied = $true } else { $w.copied = $true } }
    Save-Work
    Save-Items
    Render
    $title = if ($again) { '已重新复制全部 {0} 条' -f $targets.Count } else { '已复制 {0} 条新记录' -f $targets.Count }
    Notify $title '打开腾讯表格，点一下新一行第一列的格子，按 Ctrl+V，内容会自动分到各列。'
}

function Edit-Columns {
    $note = "按腾讯表格从左到右的列顺序填写，用逗号隔开。`n可以用：日期、日期(月日)、星期、时间、工作内容、项目、待办`n不需要自动填的列就空着，比如：日期,时间,工作内容,,项目（第 4 列粘贴后是空的）"
    $r = Show-TextDialog '表格列设置' ([string]$script:S.logCols) $note
    if ($r -and $r.Trim()) {
        $script:S.logCols = $r.Trim()
        Save-Settings
        Notify '列设置已保存' ('从左到右：' + $script:S.logCols)
    }
}

$OnLogEdit     = { param($s, $e) Edit-Work ([string]$s.Tag) }
$OnLogDelete   = { param($s, $e) Remove-Work ([string]$s.Tag) }
$OnLogTime     = { param($s, $e) Set-WorkTime ([string]$s.Tag) }
$OnLogDate     = { param($s, $e) Set-WorkDate ([string]$s.Tag) }

function Set-WorkDate([string]$id) {
    $w = Find-Work $id
    if (-not $w) { return }
    $init = Parse-Day $w.date
    if (-not $init) { $init = (Get-Date).Date }
    if ($w.time -match '^(\d{1,2}):(\d{2})$') { $init = $init.AddHours([int]$Matches[1]).AddMinutes([int]$Matches[2]) }
    $d = Show-DatePickerDialog $init '改到哪一天'
    if (-not $d) { return }
    $old = $w.date
    $w.date = $d.ToString('yyyy-MM-dd')
    $w.time = $d.ToString('HH:mm')
    Save-Work
    Render
    if ($w.date -ne $old) { Notify '日期已修改' ('「' + $w.text + '」改到了 ' + $d.ToString('M月d日 HH:mm') + '，各处已经同步。') }
}
$OnLogProj     = { param($s, $e) $p = ([string]$s.Tag).Split([char[]]'|', 2); Set-WorkProject $p[0] $p[1] }
$OnLogCopied   = { param($s, $e) $w = Find-Work ([string]$s.Tag); if ($w) { $w.copied = -not $w.copied; Save-Work; Render } }
$script:LogOpen = @{}
$OnLogToggle = {
    param($s, $e)
    $e.Handled = $true
    $id = [string]$s.Tag
    if ($script:LogOpen.ContainsKey($id)) { $script:LogOpen.Remove($id) } else { $script:LogOpen[$id] = $true }
    Render
}
function Get-LogExpandable($entries) { return @($entries | Where-Object { $_.src -and ($t = Find-Todo $_.src) -and @($t.subs).Count -gt 0 } | ForEach-Object { $_.id }) }
function Update-LogExpandLink($entries) {
    $ids = Get-LogExpandable $entries
    $script:LogVisible = $ids
    $closed = @($ids | Where-Object { -not $script:LogOpen.ContainsKey($_) })
    $BtnLogExpand.Visibility = if ($ids.Count -gt 0) { 'Visible' } else { 'Collapsed' }
    $BtnLogExpand.Text = if ($ids.Count -gt 0 -and $closed.Count -eq 0) { '全部收起' } else { '全部展开' }
}
$OnLogCardDown = { param($s, $e) if ($e.ClickCount -eq 2) { $e.Handled = $true; Edit-Work ([string]$s.Tag) } }

function New-LogCard($w, [int]$no = 0) {
    $card = [System.Windows.Controls.Border]::new()
    $card.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $card.Padding = Th 8 7 4 7
    $card.Margin = Th 0 0 0 6
    $card.Background = TB 'Card'
    $card.BorderBrush = TB 'CardLine'
    $card.BorderThickness = Th 1 1 1 1
    $card.Tag = $w.id

    $grid = [System.Windows.Controls.Grid]::new()
    foreach ($wd in 'auto', 'star', 'auto') {
        $cd = [System.Windows.Controls.ColumnDefinition]::new()
        if ($wd -eq 'star') { $cd.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star) }
        else { $cd.Width = [System.Windows.GridLength]::Auto }
        [void]$grid.ColumnDefinitions.Add($cd)
    }

    $tm = [System.Windows.Controls.TextBlock]::new()
    $tm.Text = if ($w.time) { $w.time } else { '--:--' }
    $tm.Foreground = TB 'Accent'
    $tm.FontWeight = [System.Windows.FontWeights]::SemiBold
    $tm.Margin = Th 0 0 10 0
    $tm.VerticalAlignment = 'Top'
    $tm.Cursor = $Hand
    $tm.Tag = $w.id
    $tm.ToolTip = '点击改时间'
    $tm.add_MouseLeftButtonUp($OnLogTime)
    $tm.ToolTip = '点一下修改时间；右键这条记录可以改日期'
    if (-not (Test-ShowTime)) { $tm.Visibility = 'Collapsed' }
    [System.Windows.Controls.Grid]::SetColumn($tm, 0)
    [void]$grid.Children.Add($tm)

    $sp = [System.Windows.Controls.StackPanel]::new()
    $tb = New-SelectableText $w.text 13.5
    [void](Enable-Inline $tb ('work|' + $w.id))
    $ttl = New-NumTitle $no $tb
    $ttl.Margin = Th 0 0 82 0; $ttl.MinHeight = 24
    [void]$sp.Children.Add($ttl)
    $srcT = if ($w.src) { Find-Todo $w.src } else { $null }
    if ($srcT -and @($srcT.subs).Count -gt 0) {
        $lopen = $script:LogOpen.ContainsKey($w.id)
        $dn = @(@($srcT.subs) | Where-Object { $_.done }).Count
        $tg = [System.Windows.Controls.TextBlock]::new()
        $tg.Text = $(if ($lopen) { '▾ ' } else { '▸ ' }) + ('这件事的具体事项 {0}/{1}' -f $dn, @($srcT.subs).Count)
        $tg.FontSize = 11; $tg.Margin = Th 0 3 0 0; $tg.Cursor = $Hand; $tg.Tag = $w.id
        $tg.ToolTip = if ($lopen) { '收起' } else { '展开看具体事项' }
        $tg.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
        $tg.add_MouseLeftButtonDown($OnLogToggle)
        [void]$sp.Children.Add($tg)
        if ($lopen) {
            $subArr = @($srcT.subs); $lines0 = Get-SubLines $srcT
            for ($si = 0; $si -lt $subArr.Count; $si++) {
                $sx = $subArr[$si]
                $srow = [System.Windows.Controls.DockPanel]::new()
                $srow.Margin = Th 12 1 0 0
                $sn = [System.Windows.Controls.TextBlock]::new()
                $sn.Text = ('{0}.' -f ($si + 1)) + $(if ($sx.done) { '✓' } else { '' }); $sn.FontSize = 12; $sn.MinWidth = 20; $sn.Margin = Th 0 0 4 0
                $sn.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
                [System.Windows.Controls.DockPanel]::SetDock($sn, 'Left')
                [void]$srow.Children.Add($sn)
                if ($sx.kind -eq 'image') {
                    $sl = New-SelectableText '[图片]' 12
                } else {
                    $sl = New-SelectableText $sx.text 12
                    [void](Enable-Inline $sl ('sub|' + $srcT.id + '|' + $sx.id))
                }
                $sl.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'SubBrush')
                if ($sx.done) { Set-DoneLook $sl }
                $sl.ToolTip = [string]$sl.ToolTip + $(if ($sl.ToolTip) { '；' } else { '' }) + '右键可以修改或打勾'
                Set-LazyMenu $sl { param($a) $t0 = Find-Todo $a[0]; if ($t0) { $x0 = Find-Sub $t0 $a[1]; if ($x0) { New-SubLineMenu $t0 $x0 } } } @($srcT.id, $sx.id) $sl
                [void]$srow.Children.Add($sl)
                [void]$sp.Children.Add($srow)
            }
        }
    }
    $parts = @()
    if ($script:LogShowDate) { $dd = Parse-Day $w.date; if ($dd) { $parts += ($dd.ToString('M月d日') + ' ' + $WeekNames[[int]$dd.DayOfWeek]) } }
    if ($w.proj -and $w.proj -ne 'cat-daily') { $parts += ('#' + (Get-ProjName $w.proj)) }
    if ($w.src) { $st0 = Find-Todo $w.src; if ($st0) { $parts += ('待办：' + $(if ($st0.text.Length -gt 14) { $st0.text.Substring(0, 14) + '…' } else { $st0.text })) } }
    if ($w.auto) { $parts += '打勾自动记' }
    if ($w.copied) { $parts += '已复制到表格' }
    if ($parts.Count -gt 0) {
        $meta = [System.Windows.Controls.TextBlock]::new()
        $meta.FontSize = 11
        $meta.Margin = Th 0 3 0 0
        $meta.Foreground = TB 'Muted'
        $meta.ToolTip = '双击修改项目、属于哪件待办、日期和时间'
        $meta.add_MouseLeftButtonDown($OnMetaDbl)
        $meta.Text = $parts -join '   '
        [void]$sp.Children.Add($meta)
    }
    [System.Windows.Controls.Grid]::SetColumn($sp, 1)
    [System.Windows.Controls.Grid]::SetColumnSpan($sp, 2)
    [void]$grid.Children.Add($sp)

    $lb = New-CardButtons $w.id $OnCopyWork $OnLogDelete
    $lb.Children.Insert(1, (New-EditBtn $w.id $OnEditBtnWork 0))
    [void]$grid.Children.Add($lb)
    $card.add_MouseEnter($OnCardEnter)
    $card.add_MouseLeave($OnCardLeave)

    $card.Child = $grid
    $card.add_MouseLeftButtonDown($OnLogCardDown)
    Set-LazyMenu $card { param($id) $w0 = Find-Work $id; if ($w0) { New-LogMenu $w0 } } $w.id $tb
    return $card
}

function New-LogMenu($w) {
    $menu = [System.Windows.Controls.ContextMenu]::new()
    [void]$menu.Items.Add((New-MenuItem '复制内容' $w.id $OnCopyWork))
    [void]$menu.Items.Add((New-MenuItem '修改内容' $w.id $OnLogEdit))
    [void]$menu.Items.Add((New-MenuItem '修改时间' $w.id $OnLogTime))
    [void]$menu.Items.Add((New-MenuItem '改到其他日期…' $w.id $OnLogDate))
    Add-ProjSubmenu $menu $w.id $w.proj $OnWorkSetProj
    $tm2 = New-MenuItem '属于哪件待办' '' $null
    [void]$tm2.Items.Add((New-MenuItem '不属于哪件待办' ($w.id + '|__none') $OnWorkSetTodo))
    foreach ($t in @($script:Items | Where-Object { $_.project -eq $w.proj -and (-not $_.done -or $_.id -eq $w.src) } | Sort-Object created -Descending | Select-Object -First 20)) {
        $lb2 = if ($t.text.Length -gt 26) { $t.text.Substring(0, 26) + '…' } else { $t.text }
        if ($t.id -eq $w.src) { $lb2 = '✓ ' + $lb2 }
        [void]$tm2.Items.Add((New-MenuItem $lb2 ($w.id + '|' + $t.id) $OnWorkSetTodo))
    }
    [void]$menu.Items.Add($tm2)
    $cpText = if ($w.copied) { '标为未复制（下次会再复制）' } else { '标为已复制' }
    [void]$menu.Items.Add((New-MenuItem $cpText $w.id $OnLogCopied))
    [void]$menu.Items.Add(([System.Windows.Controls.Separator]::new()))
    [void]$menu.Items.Add((New-MenuItem '删除' $w.id $OnLogDelete))
    return $menu
}

function Render-Log {
    $List.Children.Clear()
    $ds = $script:ViewDate.ToString('yyyy-MM-dd')
    $DayLabel.Text = Format-Day $script:ViewDate
    $words = Get-Words
    $script:LogShowDate = $words.Count -gt 0
    if ($words.Count -gt 0) {
        # 搜索时不分日期，所有记录里找，新的在前
        $found = @($script:Work | Where-Object { Test-WorkMatch $_ $words } | Sort-Object @{ Expression = { $_.date }; Descending = $true }, @{ Expression = { $_.time }; Descending = $true })
        if ($found.Count -eq 0) {
            $empty = New-Label '工作记录里没搜到' 18
            $empty.HorizontalAlignment = 'Center'
            [void]$List.Children.Add($empty)
        }
        $no = 0
        foreach ($w in $found) { $no++; [void]$List.Children.Add((New-LogCard $w $no)) }
        Update-LogExpandLink $found
        $Count.Text = '搜到 {0} 条' -f $found.Count
        $Count.Foreground = TB 'Sub'; $Count.ToolTip = $null
        $script:RenderSig = Get-Sig
        return
    }
    $rows = Get-DayRows $ds
    $entries = @($rows | Where-Object { $_.w } | ForEach-Object { $_.w })
    if ($rows.Count -eq 0) {
        $empty = New-Label '这一天还没有记录' 18
        $empty.HorizontalAlignment = 'Center'
        [void]$List.Children.Add($empty)
    }
    $no = 0
    foreach ($r0 in $rows) {
        $no++
        if ($r0.w) { [void]$List.Children.Add((New-LogCard $r0.w $no)) } else { [void]$List.Children.Add((New-DoneLogCard $r0.t $no)) }
    }
    Update-LogExpandLink $entries
    $doneN = @($rows | Where-Object { $_.t }).Count
    $Count.Text = '{0} 条' -f $rows.Count
    $Count.ToolTip = if ($doneN -gt 0) { '其中 {0} 条是这天完成的待办' -f $doneN } else { $null }
    $Count.Foreground = TB 'Sub'
    $BtnImport.Visibility = 'Collapsed'

    $newCount = @($rows | Where-Object { ($_.w -and -not $_.w.copied) -or ($_.t -and -not $_.t.copied) }).Count
    if ($rows.Count -eq 0) { $BtnCopyLog.Text = '复制到表格' }
    elseif ($newCount -gt 0) { $BtnCopyLog.Text = '复制到表格（{0} 条新）' -f $newCount }
    else { $BtnCopyLog.Text = '重新复制全部（{0} 条）' -f $rows.Count }
    $script:RenderSig = Get-Sig
}

$OnLogChip = {
    param($s, $e)
    $k = [string]$s.Tag
    if ($k -eq '__time') {
        $r = Show-TimeDialog ([string]$script:LogPending.time)
        if ($r) { $script:LogPending.time = $r }
    } elseif ($k -eq '__newproj') {
        $r = Show-TextDialog '新项目' '' '输入项目名称，记下的这条会归到这个项目'
        if ($r -and $r.Trim()) { $script:LogPending.project = $r.Trim(); Remember-Project $r.Trim() }
    } elseif ($k.StartsWith('p:')) {
        $p = $k.Substring(2)
        if ($script:LogPending.project -eq $p) { $script:LogPending.project = '' } else { $script:LogPending.project = $p }
    }
    Render-Chips
    [void]$InputBox.Focus()
}

function New-Chip([string]$label, [string]$tag, [bool]$sel, [string]$tip, [scriptblock]$handler) {
    $b = [System.Windows.Controls.Border]::new()
    $b.CornerRadius = [System.Windows.CornerRadius]::new(10)
    $b.Padding = Th 9 2 9 3
    $b.Margin = Th 0 0 6 4
    $b.Cursor = $Hand
    $b.Tag = $tag
    if ($tip) { $b.ToolTip = $tip }
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $label
    $t.FontSize = 12
    if ($sel) { $b.Background = TB 'Accent'; $t.Foreground = Brush '#FFFFFF' }
    else { $b.Background = TB 'Chip'; $t.Foreground = TB 'ChipText' }
    $b.Child = $t
    $b.add_MouseLeftButtonUp($handler)
    return $b
}

function Render-LogChips {
    $Chips.Children.Clear()
    if (-not $script:LogPending.proj -or -not (Find-Project $script:LogPending.proj)) { $script:LogPending.proj = if ($script:S.lastLogProj -and (Find-Project $script:S.lastLogProj)) { [string]$script:S.lastLogProj } else { 'cat-daily' } }
    $lpc = New-Chip ('项目：' + (Get-ProjName $script:LogPending.proj) + ' ▾') '__proj' ($script:LogPending.proj -ne 'cat-daily') '这条记录归到哪个项目，右键可以改项目名称' $OnLogProjChip
    $lpj = Find-Project $script:LogPending.proj
    if ($lpj -and $lpj.kind -eq 'project') {
        $lcm = [System.Windows.Controls.ContextMenu]::new()
        [void]$lcm.Items.Add((New-MenuItem ('修改「' + $lpj.text + '」的名称…') $lpj.id $OnRenameProj))
        $lpc.ContextMenu = $lcm
    }
    [void]$Chips.Children.Add($lpc)
    $tt = '属于待办 ▾'
    if ($script:LogPending.todo) { $t = Find-Todo $script:LogPending.todo; if ($t) { $tt = '待办：' + $(if ($t.text.Length -gt 10) { $t.text.Substring(0, 10) + '…' } else { $t.text }) + ' ▾' } }
    [void]$Chips.Children.Add((New-Chip $tt '__todo' ([bool]$script:LogPending.todo) '这条记录属于哪件待办，就算那件待办这天的进展' $OnLogTodoChip))
    $tl = if ($script:LogPending.time) { '时间 ' + $script:LogPending.time } else { '时间：现在' }
    [void]$Chips.Children.Add((New-Chip $tl '__time' ([bool]$script:LogPending.time) '默认按现在的时间记，补记之前的事可以点这里改' $OnLogChip))
}

function Set-Mode([string]$m) {
    if ($m -ne 'log') { $m = 'todo' }
    $script:S.mode = $m
    $isLog = ($m -eq 'log')

    $bold = [System.Windows.FontWeights]::SemiBold; $norm = [System.Windows.FontWeights]::Normal
    if ($isLog) {
        $TabTodo.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush'); $TabTodo.FontWeight = $norm
        $TabLog.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush'); $TabLog.FontWeight = $bold
        $Hint.Text = '做了什么，回车保存；打字时下面可以选项目、属于哪件待办'
        $DayBar.Visibility = 'Visible'; $FootLog.Visibility = 'Visible'; $FootTodo.Visibility = 'Collapsed'
        $ProjBarBox.Visibility = 'Collapsed'
        $script:ViewDate = (Get-Date).Date
    } else {
        $TabTodo.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush'); $TabTodo.FontWeight = $bold
        $TabLog.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush'); $TabLog.FontWeight = $norm
        $Hint.Text = '记一条，回车保存；可以写「明天下午3点」「#项目名」'
        $DayBar.Visibility = 'Collapsed'; $FootLog.Visibility = 'Collapsed'; $FootTodo.Visibility = 'Visible'
    }
    Render-Chips
    Render
    Save-Settings
}

function Render {
    if ($script:S.mode -eq 'log') { Render-Log } else { Render-Todo }
    $BtnTime.Opacity = if (Test-ShowTime) { 1 } else { 0.35 }
    $BtnTime.ToolTip = if (Test-ShowTime) { '现在显示几点几分，点一下隐藏' } else { '现在只显示日期，点一下显示几点几分' }
    if ($script:CalWin) { Render-Cal }
}
function Render-Chips { if ($script:S.mode -eq 'log') { Render-LogChips } else { Render-TodoChips }; Update-ChipsVis }
# 平时只留输入框，点进去打字时才出来那排项目、提醒时间
function Update-ChipsVis {
    $busy = $InputBox.IsKeyboardFocusWithin -or $InputBox.Text.Length -gt 0
    if ($script:S.mode -ne 'log' -and $script:Pending -and ($script:Pending.sel.Count -gt 0 -or $script:Pending.star)) { $busy = $true }
    $Chips.Visibility = if ($busy) { 'Visible' } else { 'Collapsed' }
}
function Add-Entry    { if ($script:S.mode -eq 'log') { Add-Log } else { Add-Todo } }

# ---------- 外观：配色和透明度 ----------
$Themes = [ordered]@{
    '便签黄' = @{ Text='#2B2B2B'; Sub='#8A7F68'; Muted='#A39A87'; Accent='#2F6FED'; Bg='#FFFBF2'; Header='#F3EAD3'; Line='#E3DAC6'; Hover='#E8DEC6'; InputLine='#D9CFB8'; InputBg='#FFFFFF'; Hint='#B0A690'; Card='#FFFFFF'; CardLine='#ECE4D3'; Chip='#EFE6D1'; ChipText='#6E6450'; Overdue='#FFF2EF' }
    '简洁白' = @{ Text='#1F2328'; Sub='#6B7280'; Muted='#9CA3AF'; Accent='#2563EB'; Bg='#FFFFFF'; Header='#F3F4F6'; Line='#E2E4E8'; Hover='#E5E7EB'; InputLine='#D1D5DB'; InputBg='#FFFFFF'; Hint='#9CA3AF'; Card='#F9FAFB'; CardLine='#E5E7EB'; Chip='#EEF0F3'; ChipText='#4B5563'; Overdue='#FEF2F2' }
    '薄荷绿' = @{ Text='#1F2D25'; Sub='#5F7A69'; Muted='#8FA597'; Accent='#15803D'; Bg='#F3FAF6'; Header='#DDEFE4'; Line='#C9E2D3'; Hover='#CFE6D8'; InputLine='#BFD9C9'; InputBg='#FFFFFF'; Hint='#94AE9E'; Card='#FFFFFF'; CardLine='#D6EADF'; Chip='#E1F0E7'; ChipText='#3F5E4B'; Overdue='#FDEEEC' }
    '天空蓝' = @{ Text='#1B2635'; Sub='#5C6F87'; Muted='#93A3B8'; Accent='#1D6FD8'; Bg='#F4F8FD'; Header='#DFEAF8'; Line='#CADBF0'; Hover='#D2E1F4'; InputLine='#C3D5EC'; InputBg='#FFFFFF'; Hint='#98A8BD'; Card='#FFFFFF'; CardLine='#DAE5F3'; Chip='#E3EDF9'; ChipText='#3E5470'; Overdue='#FDEFEC' }
    '樱花粉' = @{ Text='#2E2226'; Sub='#8C6670'; Muted='#B39AA0'; Accent='#D6336C'; Bg='#FFF6F7'; Header='#FBE3E7'; Line='#F1CDD4'; Hover='#F6D5DB'; InputLine='#EDC6CE'; InputBg='#FFFFFF'; Hint='#BFA2A9'; Card='#FFFFFF'; CardLine='#F3DCE1'; Chip='#FBE7EB'; ChipText='#7A4E59'; Overdue='#FFEDEA' }
    '深色'   = @{ Text='#E6E8EB'; Sub='#A0A7B2'; Muted='#737B87'; Accent='#6EA8FE'; Bg='#1F2228'; Header='#2A2E36'; Line='#3A3F49'; Hover='#383D47'; InputLine='#454B56'; InputBg='#2A2E36'; Hint='#737B87'; Card='#282C33'; CardLine='#353A43'; Chip='#333842'; ChipText='#C3C8D0'; Overdue='#3A2626' }
}

function TB([string]$k) { return $window.Resources[$k + 'Brush'] }

function Mix-Rgb([int[]]$a, [int[]]$b, [double]$t) {
    return @(
        [int][math]::Round($a[0] + ($b[0] - $a[0]) * $t),
        [int][math]::Round($a[1] + ($b[1] - $a[1]) * $t),
        [int][math]::Round($a[2] + ($b[2] - $a[2]) * $t))
}
function To-Hex([int[]]$c) { return ('#{0:X2}{1:X2}{2:X2}' -f $c[0], $c[1], $c[2]) }
function Get-Lum([int[]]$c) { return (0.299 * $c[0] + 0.587 * $c[1] + 0.114 * $c[2]) }

function New-PaletteFromColor([string]$hex) {
    $col = [System.Drawing.ColorTranslator]::FromHtml($hex)
    $base = @([int]$col.R, [int]$col.G, [int]$col.B)
    $white = @(255, 255, 255); $black = @(0, 0, 0)
    if ((Get-Lum $base) -ge 100) {
        $sub = Mix-Rgb $base $black 0.55
        $acc = Mix-Rgb $base $black 0.25
        $i = 0
        while ((Get-Lum $acc) -gt 120 -and $i -lt 10) { $acc = Mix-Rgb $acc $black 0.15; $i++ }
        return @{
            Text = '#24262B'; Sub = To-Hex $sub; Muted = To-Hex (Mix-Rgb $sub $white 0.4); Accent = To-Hex $acc
            Bg = To-Hex (Mix-Rgb $base $white 0.88); Header = To-Hex (Mix-Rgb $base $white 0.6)
            Line = To-Hex (Mix-Rgb $base $white 0.5); Hover = To-Hex (Mix-Rgb $base $white 0.45)
            InputLine = To-Hex (Mix-Rgb $base $white 0.55); InputBg = '#FFFFFF'; Hint = To-Hex (Mix-Rgb $sub $white 0.45)
            Card = '#FFFFFF'; CardLine = To-Hex (Mix-Rgb $base $white 0.75)
            Chip = To-Hex (Mix-Rgb $base $white 0.75); ChipText = To-Hex (Mix-Rgb $base $black 0.5); Overdue = '#FFEFEC'
        }
    }
    return @{
        Text = '#E8EAED'; Sub = '#A8AEB8'; Muted = '#7C838E'; Accent = To-Hex (Mix-Rgb $base $white 0.55)
        Bg = To-Hex (Mix-Rgb $base $black 0.45); Header = To-Hex (Mix-Rgb $base $black 0.25)
        Line = To-Hex (Mix-Rgb $base $white 0.18); Hover = To-Hex (Mix-Rgb $base $white 0.12)
        InputLine = To-Hex (Mix-Rgb $base $white 0.22); InputBg = To-Hex (Mix-Rgb $base $black 0.25); Hint = '#7C838E'
        Card = To-Hex (Mix-Rgb $base $black 0.3); CardLine = To-Hex (Mix-Rgb $base $white 0.12)
        Chip = To-Hex (Mix-Rgb $base $white 0.15); ChipText = '#D0D4DA'; Overdue = '#3A2626'
    }
}

function Get-Palette([string]$name) {
    if ($name -eq '自定义' -and $script:S.customColor) { return (New-PaletteFromColor ([string]$script:S.customColor)) }
    if ($Themes.Contains($name)) { return $Themes[$name] }
    return $Themes['便签黄']
}

function Apply-Theme([string]$name) {
    if ($name -ne '自定义' -and -not $Themes.Contains($name)) { $name = '便签黄' }
    $pal = Get-Palette $name
    foreach ($k in @($pal.Keys)) {
        $window.Resources[$k + 'Brush'] = $BC.ConvertFromString($pal[$k])
        $app.Resources[$k + 'Brush'] = $BC.ConvertFromString($pal[$k])
    }
    $script:S.theme = $name
    Render-Chips
    Render
    Save-Settings
}

function Update-Opacity {
    $o = [double]$script:S.opacity
    if ($o -lt 0.3 -or $o -gt 1) { $o = 1 }
    if ($script:S.hoverOpaque -and ($window.IsMouseOver -or $window.IsActive)) { $o = 1 }
    $window.Opacity = $o
}

$OnSwatch = {
    param($s, $e)
    $name = [string]$s.Tag
    if ($name -eq '__pick') {
        $cd = [System.Windows.Forms.ColorDialog]::new()
        $cd.FullOpen = $true
        if ($script:S.customColor) { try { $cd.Color = [System.Drawing.ColorTranslator]::FromHtml([string]$script:S.customColor) } catch {} }
        if ($cd.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }
        $script:S.customColor = '#{0:X2}{1:X2}{2:X2}' -f $cd.Color.R, $cd.Color.G, $cd.Color.B
        $name = '自定义'
    }
    Apply-Theme $name
    Render-Swatches
}

function New-Swatch([string]$name, $pal, [string]$label, [string]$tag) {
    $outer = [System.Windows.Controls.StackPanel]::new()
    $outer.Margin = Th 0 0 10 10
    $outer.Cursor = $Hand
    $outer.Tag = $tag
    $box = [System.Windows.Controls.Border]::new()
    $box.Width = 64; $box.Height = 44
    $box.CornerRadius = [System.Windows.CornerRadius]::new(6)
    $box.BorderThickness = Th 2 2 2 2
    $box.BorderBrush = if ($script:S.theme -eq $name) { Brush '#2F6FED' } else { Brush '#D0D0D0' }
    $box.ClipToBounds = $true
    if ($pal) {
        $box.Background = Brush $pal.Bg
        $inner = [System.Windows.Controls.StackPanel]::new()
        $top = [System.Windows.Controls.Border]::new()
        $top.Height = 13; $top.Background = Brush $pal.Header
        $top.CornerRadius = [System.Windows.CornerRadius]::new(4, 4, 0, 0)
        [void]$inner.Children.Add($top)
        $line = [System.Windows.Controls.Border]::new()
        $line.Height = 5; $line.Width = 34; $line.Margin = Th 6 6 0 0; $line.HorizontalAlignment = 'Left'
        $line.Background = Brush $pal.Accent; $line.CornerRadius = [System.Windows.CornerRadius]::new(2)
        [void]$inner.Children.Add($line)
        $line2 = [System.Windows.Controls.Border]::new()
        $line2.Height = 5; $line2.Width = 24; $line2.Margin = Th 6 4 0 0; $line2.HorizontalAlignment = 'Left'
        $line2.Background = Brush $pal.Muted; $line2.CornerRadius = [System.Windows.CornerRadius]::new(2)
        [void]$inner.Children.Add($line2)
        $box.Child = $inner
    } else {
        $box.Background = Brush '#FFFFFF'
        $plus = [System.Windows.Controls.TextBlock]::new()
        $plus.Text = '+'; $plus.FontSize = 20; $plus.Foreground = Brush '#888888'
        $plus.HorizontalAlignment = 'Center'; $plus.VerticalAlignment = 'Center'
        $box.Child = $plus
    }
    [void]$outer.Children.Add($box)
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $label; $t.FontSize = 12; $t.HorizontalAlignment = 'Center'; $t.Margin = Th 0 3 0 0
    [void]$outer.Children.Add($t)
    $outer.add_MouseLeftButtonUp($OnSwatch)
    return $outer
}

function Render-Swatches {
    $p = $script:SwatchPanel
    $p.Children.Clear()
    foreach ($name in $Themes.Keys) { [void]$p.Children.Add((New-Swatch $name $Themes[$name] $name $name)) }
    if ($script:S.customColor) {
        [void]$p.Children.Add((New-Swatch '自定义' (New-PaletteFromColor ([string]$script:S.customColor)) '自定义' '自定义'))
    }
    [void]$p.Children.Add((New-Swatch '__none' $null '选颜色…' '__pick'))
}

function Show-StyleDialog {
    if (-not $window.IsVisible) { Show-Main }
    $script:Dlg = New-DialogWindow '外观'
    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.Margin = Th 16 14 16 14
    $sp.Width = 330

    $h1 = [System.Windows.Controls.DockPanel]::new()
    $l1 = [System.Windows.Controls.TextBlock]::new()
    $l1.Text = '透明度'; $l1.FontWeight = [System.Windows.FontWeights]::SemiBold
    $script:OpLabel = [System.Windows.Controls.TextBlock]::new()
    $script:OpLabel.HorizontalAlignment = 'Right'; $script:OpLabel.Foreground = Brush '#666666'
    [System.Windows.Controls.DockPanel]::SetDock($script:OpLabel, 'Right')
    [void]$h1.Children.Add($script:OpLabel); [void]$h1.Children.Add($l1)
    [void]$sp.Children.Add($h1)

    $script:OpSlider = [System.Windows.Controls.Slider]::new()
    $script:OpSlider.Minimum = 30; $script:OpSlider.Maximum = 100
    $script:OpSlider.TickFrequency = 5; $script:OpSlider.IsSnapToTickEnabled = $true
    $script:OpSlider.Margin = Th 0 6 0 4
    $script:OpSlider.Value = [math]::Round([double]$script:S.opacity * 100)
    $script:OpLabel.Text = ('{0}%' -f [int]$script:OpSlider.Value)
    $script:OpSlider.add_ValueChanged({
        $script:S.opacity = [double]$script:OpSlider.Value / 100
        $script:OpLabel.Text = ('{0}%' -f [int]$script:OpSlider.Value)
        Update-Opacity
    })
    [void]$sp.Children.Add($script:OpSlider)

    $script:HoverCb = [System.Windows.Controls.CheckBox]::new()
    $script:HoverCb.Content = '鼠标移上去或正在打字时，自动变回不透明'
    $script:HoverCb.IsChecked = [bool]$script:S.hoverOpaque
    $script:HoverCb.Margin = Th 0 2 0 14
    $script:HoverCb.add_Click({ $script:S.hoverOpaque = [bool]$script:HoverCb.IsChecked; Update-Opacity })
    [void]$sp.Children.Add($script:HoverCb)
    $script:TimeCb = [System.Windows.Controls.CheckBox]::new()
    $script:TimeCb.Content = '显示几点几分（关掉后待办、完成、工作记录、周月视图都只显示内容和日期）'
    $script:TimeCb.IsChecked = Test-ShowTime
    $script:TimeCb.Margin = Th 0 0 0 14
    $script:TimeCb.add_Click({ Set-ShowTime ([bool]$script:TimeCb.IsChecked) })
    [void]$sp.Children.Add($script:TimeCb)

    $l2 = [System.Windows.Controls.TextBlock]::new()
    $l2.Text = '配色'; $l2.FontWeight = [System.Windows.FontWeights]::SemiBold; $l2.Margin = Th 0 0 0 8
    [void]$sp.Children.Add($l2)
    $script:SwatchPanel = [System.Windows.Controls.WrapPanel]::new()
    [void]$sp.Children.Add($script:SwatchPanel)
    Render-Swatches

    $ok = [System.Windows.Controls.Button]::new()
    $ok.Content = '完成'; $ok.Width = 72; $ok.Height = 28; $ok.IsDefault = $true; $ok.IsCancel = $true
    $ok.HorizontalAlignment = 'Right'; $ok.Margin = Th 0 6 0 0
    [void]$sp.Children.Add($ok)

    $script:Dlg.Content = $sp
    [void]$script:Dlg.ShowDialog()
    Save-Settings
    Update-Opacity
}

# ---------- 报告：按项目汇总 ----------
$CnNum = @('', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十', '十一', '十二', '十三', '十四', '十五', '十六', '十七', '十八', '十九', '二十')
function Get-CnNum([int]$n) { if ($n -lt $CnNum.Count) { return $CnNum[$n] }; return [string]$n }

function Get-PeriodRange([string]$kind, [datetime]$anchor) {
    $a = $anchor.Date
    switch ($kind) {
        'day'     { return @($a, $a, ('工作日报（{0} {1}）' -f $a.ToString('yyyy年M月d日'), $WeekNames[[int]$a.DayOfWeek])) }
        'week'    { $f = Get-WeekStart $a; $t = $f.AddDays(6); return @($f, $t, ('工作周报（{0}—{1}）' -f $f.ToString('yyyy年M月d日'), $t.ToString('M月d日'))) }
        'month'   { $f = [datetime]::new($a.Year, $a.Month, 1); return @($f, $f.AddMonths(1).AddDays(-1), ('{0}工作月报' -f $f.ToString('yyyy年M月'))) }
        'quarter' { $q = [math]::Floor(($a.Month - 1) / 3); $f = [datetime]::new($a.Year, $q * 3 + 1, 1); return @($f, $f.AddMonths(3).AddDays(-1), ('{0}年第{1}季度工作总结' -f $a.Year, (Get-CnNum ($q + 1)))) }
        'year'    { $f = [datetime]::new($a.Year, 1, 1); return @($f, $f.AddYears(1).AddDays(-1), ('{0}年度工作总结' -f $a.Year)) }
    }
    $f = Parse-Day ([string]$script:S.repFrom); $t = Parse-Day ([string]$script:S.repTo)
    if (-not $f) { $f = $a.AddDays(-6) }; if (-not $t) { $t = $a }
    if ($t -lt $f) { $x = $f; $f = $t; $t = $x }
    return @($f, $t, ('工作总结（{0}—{1}）' -f $f.ToString('yyyy年M月d日'), $t.ToString('yyyy年M月d日')))
}

function Get-ShortDate([datetime]$d) { return $d.ToString('M月d日') }

# 报告的内容先整理成一行一行（带层级），文字版和 Word 版都从这里生成
# 层级：title 标题 / h1 项目 / h2 小节 / item 条目 / sub 条目下的细项 / note 说明
function Build-ReportLines([datetime]$from, [datetime]$to, [string]$title, [bool]$withSubs, [string]$onlyProj = '') {
    $script:RepAcc = New-Object System.Collections.ArrayList
    $L = $script:RepAcc
    function Add-L([string]$lv, [string]$tx, [string]$ref = '') { [void]$script:RepAcc.Add([pscustomobject]@{ lv = $lv; text = $tx; ref = $ref }) }
    Add-L 'title' $title
    $inRange = { param($d) $d -and $d.Date -ge $from -and $d.Date -le $to }
    $projs = if ($onlyProj) { @(Find-Project $onlyProj) } else { @(Get-ProjectsOrdered $true) }
    $no = 0
    foreach ($p in $projs) {
        if (-not $p) { continue }
        $pid0 = $p.id
        $todos = @($script:Items | Where-Object { $_.project -eq $pid0 })
        $works = @($script:Work | Where-Object { $_.proj -eq $pid0 -and (& $inRange (Parse-Day $_.date)) } | Sort-Object date, time, seq)

        if ($pid0 -eq 'cat-reimb') {
            $rows = @($todos | Where-Object { $c = Parse-Dt $_.created; (& $inRange $c) -or ($_.rstate -ne '已到账' -and $c -and $c.Date -le $to) } | Sort-Object created)
            if ($rows.Count -eq 0) { continue }
            $no++; Add-L 'h1' ('{0}、报销' -f (Get-CnNum $no))
            $sum = @{ '待报' = 0.0; '已提交' = 0.0; '已到账' = 0.0 }; $total = 0.0; $i = 1
            foreach ($t in $rows) {
                $v = 0.0; [void][double]::TryParse([string]$t.amount, [ref]$v)
                $st = if ($t.rstate) { $t.rstate } else { '待报' }
                $sum[$st] += $v; $total += $v
                $c = Parse-Dt $t.created
                $old = if ($c -and $c.Date -lt $from) { '，之前的未结清' } else { '' }
                $m = Format-Money $t.amount
                Add-L 'item' ('{0}. {1}  {2}（{3}{4}）' -f $i, $t.text, $(if ($m) { $m + ' 元' } else { '金额未填' }), $st, $old) ('todo|' + $t.id); $i++
            }
            Add-L 'note' ('合计 {0:N2} 元：待报 {1:N2} 元，已提交 {2:N2} 元，已到账 {3:N2} 元' -f $total, $sum['待报'], $sum['已提交'], $sum['已到账'])
            continue
        }

        $done = @($todos | Where-Object { $_.done -and (& $inRange (Parse-Dt $_.doneAt)) } | Sort-Object doneAt)
        $worksByTodo = @{}
        foreach ($w in $works) { if ($w.src) { if (-not $worksByTodo.ContainsKey($w.src)) { $worksByTodo[$w.src] = New-Object System.Collections.ArrayList }; [void]$worksByTodo[$w.src].Add($w) } }
        $open = @($todos | Where-Object { -not $_.done -and ($c = Parse-Dt $_.created) -and $c.Date -le $to })
        $doing = @($open | Where-Object { $worksByTodo.ContainsKey($_.id) -or ((Parse-Day $_.start) -and (Parse-Day $_.start) -le $to) })
        $shown = @{}; foreach ($t0 in @($done) + @($doing)) { $shown[$t0.id] = $true }
        $loose = @($works | Where-Object { -not $_.src -or -not $shown.ContainsKey($_.src) })
        $next = @($open | Where-Object { $d = Parse-Dt $_.due; $d -and $d.Date -gt $to } | Sort-Object { Parse-Dt $_.due } | Select-Object -First 8)
        $pending = @($open | Where-Object { -not ($doing -contains $_) -and -not ($next -contains $_) })

        $isCat = $p.kind -eq 'category'
        if ($done.Count + $doing.Count + $loose.Count -eq 0 -and ($isCat -or $next.Count + $pending.Count -eq 0)) { continue }
        $no++
        Add-L 'h1' ('{0}、{1}' -f (Get-CnNum $no), $p.text) $(if ($p.kind -eq 'project') { 'proj|' + $pid0 } else { '' })

        $writeTodo = {
            param($t, [string]$head)
            Add-L 'item' $head ('todo|' + $t.id)
            $ws = if ($worksByTodo.ContainsKey($t.id)) { @($worksByTodo[$t.id]) } else { @() }
            $ws = @($ws | Where-Object { ([string]$_.text).Trim() -ne ([string]$t.text).Trim() })
            foreach ($w in $ws) { Add-L 'sub' ('{0}：{1}' -f (Get-ShortDate (Parse-Day $w.date)), $w.text) ('work|' + $w.id) }
            if ($withSubs -and $ws.Count -eq 0) { $rsubs = @($t.subs); $rls = Get-SubLines $t; for ($ri = 0; $ri -lt $rsubs.Count; $ri++) { Add-L 'sub' $rls[$ri] ('sub|' + $t.id + '|' + $rsubs[$ri].id) } }
        }
        $sec = 0
        if ($done.Count -gt 0) {
            $sec++; Add-L 'h2' ('（{0}）已完成' -f (Get-CnNum $sec))
            $i = 1
            foreach ($t in $done) { & $writeTodo $t ('{0}. {1}（{2}完成）' -f $i, $t.text, (Get-ShortDate (Parse-Dt $t.doneAt))); $i++ }
        }
        if ($doing.Count -gt 0) {
            $sec++; Add-L 'h2' ('（{0}）推进中' -f (Get-CnNum $sec))
            $i = 1
            foreach ($t in $doing) {
                $d = Parse-Dt $t.due
                $tail = if ($d) { '，计划' + (Get-ShortDate $d) + '前完成' } else { '' }
                $sd = Parse-Day $t.start
                $from0 = if ($sd) { (Get-ShortDate $sd) + '开始' } else { '进行中' }
                & $writeTodo $t ('{0}. {1}（{2}{3}）' -f $i, $t.text, $from0, $tail); $i++
            }
        }
        if ($loose.Count -gt 0) {
            $sec++; Add-L 'h2' ('（{0}）{1}' -f (Get-CnNum $sec), $(if ($isCat) { '工作记录' } else { '其他工作' }))
            foreach ($w in $loose) {
                $lt = if ($w.src) { Find-Todo $w.src } else { $null }
                $pre = if ($lt) { '【' + $lt.text + '】' } else { '' }
                Add-L 'item' ('{0}：{1}{2}' -f (Get-ShortDate (Parse-Day $w.date)), $pre, $w.text) ('work|' + $w.id)
            }
        }
        if (-not $isCat -and ($next.Count + $pending.Count) -gt 0) {
            $sec++; Add-L 'h2' ('（{0}）下一步' -f (Get-CnNum $sec))
            foreach ($t in $next) { Add-L 'item' ('{0}前：{1}' -f (Get-ShortDate (Parse-Dt $t.due)), $t.text) ('todo|' + $t.id) }
            foreach ($t in @($pending | Select-Object -First 8)) { Add-L 'item' ('待办：' + $t.text) ('todo|' + $t.id) }
        }
    }
    if ($no -eq 0) { Add-L 'note' '这段时间没有记录。' }
    return , $L
}

function Get-ReportText($lines) {
    $sb = New-Object System.Text.StringBuilder
    foreach ($x in $lines) {
        switch ($x.lv) {
            'title' { [void]$sb.AppendLine($x.text) }
            'h1'    { [void]$sb.AppendLine(''); [void]$sb.AppendLine($x.text) }
            'h2'    { [void]$sb.AppendLine($x.text) }
            'item'  { [void]$sb.AppendLine('  ' + $x.text) }
            'sub'   { [void]$sb.AppendLine('      ' + $x.text) }
            'note'  { [void]$sb.AppendLine('  ' + $x.text) }
        }
    }
    return $sb.ToString().TrimEnd()
}

function Get-XmlText([string]$t) { return [System.Security.SecurityElement]::Escape($t) }

# 生成 Word 文件（.docx），不依赖 Office，直接按 Word 的文件格式写
function Export-ReportDocx($lines, [string]$path) {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    $body = New-Object System.Text.StringBuilder
    foreach ($x in $lines) {
        $font = '仿宋'; $size = 32; $bold = $false; $jc = ''; $ind = 0; $first = 0; $before = 0; $after = 0
        switch ($x.lv) {
            'title' { $font = '黑体'; $size = 40; $bold = $true; $jc = 'center'; $after = 240 }
            'h1'    { $font = '黑体'; $size = 32; $before = 200; $after = 60; $first = 640 }
            'h2'    { $font = '楷体'; $size = 32; $bold = $true; $first = 640 }
            'item'  { $first = 640 }
            'sub'   { $ind = 1280; $size = 28 }
            'note'  { $first = 640 }
        }
        $ppr = '<w:pPr><w:spacing w:before="' + $before + '" w:after="' + $after + '" w:line="560" w:lineRule="exact"/>'
        if ($ind -gt 0 -or $first -gt 0) { $ppr += '<w:ind w:left="' + $ind + '" w:firstLine="' + $first + '"/>' }
        if ($jc) { $ppr += '<w:jc w:val="' + $jc + '"/>' }
        $ppr += '</w:pPr>'
        $rpr = '<w:rPr><w:rFonts w:ascii="' + $font + '" w:hAnsi="' + $font + '" w:eastAsia="' + $font + '"/>' + $(if ($bold) { '<w:b/>' } else { '' }) + '<w:sz w:val="' + $size + '"/><w:szCs w:val="' + $size + '"/></w:rPr>'
        [void]$body.Append('<w:p>' + $ppr + '<w:r>' + $rpr + '<w:t xml:space="preserve">' + (Get-XmlText $x.text) + '</w:t></w:r></w:p>')
    }
    $doc = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + $body.ToString() +
        '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="2098" w:right="1474" w:bottom="1984" w:left="1588" w:header="851" w:footer="992" w:gutter="0"/></w:sectPr></w:body></w:document>'
    $ct = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'
    $rels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $fs = [IO.File]::Open($path, [IO.FileMode]::CreateNew)
    try {
        $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
        foreach ($pair in @(@('[Content_Types].xml', $ct), @('_rels/.rels', $rels), @('word/document.xml', $doc))) {
            $en = $zip.CreateEntry($pair[0])
            $sw = New-Object IO.StreamWriter($en.Open(), (New-Object System.Text.UTF8Encoding($false)))
            $sw.Write($pair[1]); $sw.Dispose()
        }
        $zip.Dispose()
    } finally { $fs.Dispose() }
}

function Copy-WeekSummary {
    $r = Get-PeriodRange 'week' (Get-Date)
    $text = Get-ReportText (Build-ReportLines $r[0] $r[1] $r[2] $true)
    try { [System.Windows.Clipboard]::SetText($text) }
    catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。'; return }
    Notify '已复制本周周报' '按项目整理好的，直接到周报里粘贴。要导出 Word 或看月报、季报，打开工作台里的"报告"。'
}

# ---------- 工作台：周视图 / 月视图 / 报告 / 项目 ----------
$script:CalWin = $null
$script:CalMode = 'week'          # week / month / report / project
$script:CalAnchor = (Get-Date).Date
$script:CalSel = (Get-Date).Date
$script:CalShow = @{ todo = $true; done = $true; work = $true; ongoing = $true }
$script:CalOpen = @{}             # 展开了具体事项的事项 id
$script:CalEdit = $null           # 右侧正在编辑的：@{ kind = 'todo'|'work'; id = ... }
$script:RepKind = 'week'
$script:RepAnchor = (Get-Date).Date
$script:ProjSel = $null

function Get-WeekStart([datetime]$d) { return $d.Date.AddDays(-((([int]$d.DayOfWeek) + 6) % 7)) }

function Get-SubLines($it) {
    $out = New-Object System.Collections.Generic.List[string]
    if (-not $it) { return , $out }
    $i = 1
    foreach ($x in @($it.subs)) {
        $body = if ($x.kind -eq 'image') { '[图片]' } else { $x.text }
        $mark = if ($x.done) { '（已完成）' } else { '' }
        $out.Add(('{0}. {1}{2}' -f $i, $body, $mark)); $i++
    }
    return , $out
}

# 一段日期内每天的内容：提醒/截止、完成、进展（工作记录）、跨天进行中
function Build-CalMap([datetime]$from, [datetime]$to) {
    $map = @{}
    $seen = @{}   # 某天已经出现过的事项，就不再重复显示"进行中"
    function AddTo([string]$ds, $obj) {
        if (-not $map.ContainsKey($ds)) { $map[$ds] = New-Object System.Collections.ArrayList }
        [void]$map[$ds].Add($obj)
        if ($obj.todo) { $seen[$ds + '|' + $obj.todo.id] = $true }
    }
    $today = (Get-Date).Date
    foreach ($it in $script:Items) {
        if ($it.done) {
            $da = Parse-Dt $it.doneAt
            if ($da -and $da.Date -ge $from -and $da.Date -le $to -and $script:CalShow.done) {
                AddTo $da.ToString('yyyy-MM-dd') ([pscustomobject]@{ kind = 'done'; time = $da.ToString('HH:mm'); text = $it.text; tag = '完成'; todo = $it; work = $null; proj = $it.project })
            }
        } elseif ($script:CalShow.todo) {
            foreach ($ts in @($it.times)) {
                $d = Parse-Dt $ts
                if (-not $d -or $d.Date -lt $from -or $d.Date -gt $to) { continue }
                $tag = if ($ts -eq $it.due) { '截止' } else { '提醒' }
                AddTo $d.ToString('yyyy-MM-dd') ([pscustomobject]@{ kind = 'todo'; time = $d.ToString('HH:mm'); text = $it.text; tag = $tag; todo = $it; work = $null; proj = $it.project })
            }
        }
    }
    if ($script:CalShow.work) {
        foreach ($w in $script:Work) {
            $d = Parse-Day $w.date
            if (-not $d -or $d -lt $from -or $d -gt $to) { continue }
            $src = if ($w.src) { Find-Todo $w.src } else { $null }
            # 从待办导入的同名记录：和这天的"完成"是同一件事，只显示"完成"那条
            if ($src -and $src.done -and $script:CalShow.done -and ([string]$w.text).Trim() -eq ([string]$src.text).Trim()) {
                $da1 = Parse-Dt $src.doneAt
                if ($da1 -and $da1.ToString('yyyy-MM-dd') -eq $w.date) { continue }
            }
            AddTo $w.date ([pscustomobject]@{ kind = 'work'; time = $w.time; text = (Get-WorkShowText $w $src); tag = ''; todo = $src; work = $w; proj = $w.proj })
        }
    }
    if ($script:CalShow.ongoing) {
        $firstProg = @{}
        foreach ($w in $script:Work) { if ($w.src -and (-not $firstProg.ContainsKey($w.src) -or $w.date -lt $firstProg[$w.src])) { $firstProg[$w.src] = $w.date } }
        foreach ($it in $script:Items) {
            if ($it.done) { $da0 = Parse-Dt $it.doneAt; if (-not $da0 -or $da0.Date -lt $from) { continue } }
            $s0 = Parse-Day $it.start
            if (-not $s0 -and $firstProg.ContainsKey($it.id)) { $s0 = Parse-Day $firstProg[$it.id] }
            if (-not $s0) { continue }
            if ($it.done) { $e0 = (Parse-Dt $it.doneAt); if ($e0) { $e0 = $e0.Date } else { continue } }
            else { $dd = Parse-Dt $it.due; $e0 = if ($dd -and $dd.Date -gt $today) { $dd.Date } else { $today } }
            if ($e0 -lt $s0) { continue }
            $a = if ($s0 -gt $from) { $s0 } else { $from }
            $b = if ($e0 -lt $to) { $e0 } else { $to }
            for ($d = $a; $d -le $b; $d = $d.AddDays(1)) {
                $ds = $d.ToString('yyyy-MM-dd')
                if ($seen.ContainsKey($ds + '|' + $it.id)) { continue }
                AddTo $ds ([pscustomobject]@{ kind = 'ongoing'; time = ''; text = $it.text; tag = '进行中'; todo = $it; work = $null; proj = $it.project })
            }
        }
    }
    $kindOrder = @{ todo = 0; work = 1; done = 2; ongoing = 3 }
    foreach ($k in @($map.Keys)) { $map[$k] = @($map[$k] | Sort-Object @{ Expression = { $kindOrder[$_.kind] } }, @{ Expression = { if ($_.time) { $_.time } else { '99:99' } } }) }
    return $map
}

$CalKindColor = @{ todo = '#2F6FED'; done = '#9AA0A6'; work = '#2E9E5B'; ongoing = '#C9A227' }

function Get-CalRange {
    if ($script:CalMode -eq 'month') {
        $first = [datetime]::new($script:CalAnchor.Year, $script:CalAnchor.Month, 1)
        return @($first, $first.AddMonths(1).AddDays(-1))
    }
    $s = Get-WeekStart $script:CalAnchor
    return @($s, $s.AddDays(6))
}

function Flush-CalFocus {
    $fe0 = [System.Windows.Input.Keyboard]::FocusedElement
    if (($fe0 -is [System.Windows.Controls.TextBox]) -and -not $fe0.IsReadOnly -and $script:CalWin -and $script:CalWin.IsKeyboardFocusWithin) { [System.Windows.Input.Keyboard]::ClearFocus() }
}
function Update-CalHighlight {
    foreach ($c in $script:CalGrid.Children) {
        if ($c -is [System.Windows.Controls.Border] -and $c.Tag -is [datetime]) {
            $c.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, $(if (([datetime]$c.Tag).Date -eq $script:CalSel.Date) { 'HoverBrush' } else { 'CardBrush' }))
        }
    }
}
# 单击：只换右边的编辑区，不重画整个日历（这样双击能接着改）
$OnCalItemClick = {
    param($s, $e)
    $e.Handled = $true
    if ($script:CalInline) { return }
    $p = ([string]$s.Tag).Split('|')
    $script:CalSel = [datetime]::ParseExact($p[2], 'yyyy-MM-dd', $Inv)
    $script:CalEdit = @{ kind = $p[0]; id = $p[1] }
    $script:CalSideOpen = $true
    Flush-CalFocus
    if ($script:CalWin -and ($script:CalMode -eq 'week' -or $script:CalMode -eq 'month') -and $script:CalGrid.IsAncestorOf($s)) {
        Update-CalHighlight
        Render-CalDetail (Build-CalMap $script:CalSel $script:CalSel)
    } else { Render-Cal }
}
# 双击日历里的一条：就地改文字
$OnCalItemDbl = {
    param($s, $e)
    if ($e.ClickCount -ne 2) { return }
    $e.Handled = $true
    $tb0 = $s.Resources['tb']
    $p = ([string]$s.Tag).Split('|')
    if ($tb0) { Start-SwapEdit $tb0 ($p[0] + '|' + $p[1]) }
}
$OnCalSubDbl = {
    param($s, $e)
    if ($e.ClickCount -ne 2) { return }
    $e.Handled = $true
    Start-SwapEdit $s ([string]$s.Tag)
}
$OnCalToggle = {
    param($s, $e)
    $e.Handled = $true
    $id = [string]$s.Tag
    if ($script:CalOpen.ContainsKey($id)) { $script:CalOpen.Remove($id) } else { $script:CalOpen[$id] = $true }
    Render-Cal
}

function New-CalItemLine($e, [string]$ds, [double]$size, [int]$no = 0, [bool]$inPanel = $false, [bool]$compact = $false) {
    $b = [System.Windows.Controls.Border]::new()
    $b.BorderThickness = Th 3 0 0 0
    $b.BorderBrush = Brush $CalKindColor[$e.kind]
    $b.Padding = Th 4 1 2 1
    $b.Margin = Th 0 0 0 3
    $b.Cursor = $Hand
    $b.Background = [System.Windows.Media.Brushes]::Transparent
    if ($e.kind -eq 'work') { $b.Tag = 'work|' + $e.work.id + '|' + $ds } else { $b.Tag = 'todo|' + $e.todo.id + '|' + $ds }
    $kindName = switch ($e.kind) { 'todo' { '待办的' + $e.tag } 'done' { '这天完成的待办' } 'work' { '工作记录' } default { '进行中的待办' } }
    $tipHead = $kindName + ' · ' + (Get-ProjName $e.proj)
    $b.ToolTip = if ($inPanel) { $tipHead + "`n点一下修改" } elseif ($compact) { $tipHead + "`n" + $e.text + "`n点一下在右边修改；双击直接改文字" } else { $tipHead + "`n点一下在右边修改；双击直接改文字" }
    $b.add_MouseLeftButtonUp($OnCalItemClick)
    if (-not $inPanel) { $b.add_MouseLeftButtonDown($OnCalItemDbl) }
    if ($e.kind -eq 'ongoing') { $b.Opacity = 0.7 }

    $sp = [System.Windows.Controls.StackPanel]::new()
    $row = [System.Windows.Controls.DockPanel]::new()
    $hasSubs = (-not $compact) -and $e.todo -and @($e.todo.subs).Count -gt 0
    if ($no -gt 0) {
        $nl = [System.Windows.Controls.TextBlock]::new()
        $nl.Text = '{0}.' -f $no; $nl.FontSize = $size; $nl.MinWidth = 16; $nl.Margin = Th 0 0 3 0
        $nl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
        [System.Windows.Controls.DockPanel]::SetDock($nl, 'Left')
        [void]$row.Children.Add($nl)
    }
    if ($hasSubs) {
        $ar = [System.Windows.Controls.TextBlock]::new()
        $open = $script:CalOpen.ContainsKey($e.todo.id)
        $ar.Text = if ($open) { '▾' } else { '▸' }
        $ar.Width = 13; $ar.FontSize = $size; $ar.Tag = $e.todo.id; $ar.Cursor = $Hand
        $ar.ToolTip = if ($open) { '收起具体事项' } else { '展开具体事项' }
        $ar.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
        $ar.add_MouseLeftButtonUp($OnCalToggle)
        $ar.add_MouseLeftButtonDown($OnSwallow)
        [System.Windows.Controls.DockPanel]::SetDock($ar, 'Left')
        [void]$row.Children.Add($ar)
    }
    $t = [System.Windows.Controls.TextBlock]::new()
    if ($compact) { $t.TextWrapping = 'NoWrap'; $t.TextTrimming = 'CharacterEllipsis' } else { $t.TextWrapping = 'Wrap' }
    $t.FontSize = $size
    # 前面只放时间；提醒/截止、进行中才多一个字样；完成的打 ✓；工作记录什么都不加
    $head = @()
    if ($e.time -and (Test-ShowTime)) { $head += $e.time }
    if ($e.kind -eq 'todo' -or $e.kind -eq 'ongoing') { $head += $e.tag }
    $r1 = New-Object System.Windows.Documents.Run ($(if ($head.Count -gt 0) { ($head -join ' ') + '  ' } else { '' }))
    $r1.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'MutedBrush')
    $r2 = New-Object System.Windows.Documents.Run ($(if ($e.kind -eq 'done') { '✓ ' } else { '' }) + $e.text)
    if ($e.kind -eq 'done') { $r2.SetResourceReference([System.Windows.Documents.TextElement]::ForegroundProperty, 'SubBrush') }
    [void]$t.Inlines.Add($r1); [void]$t.Inlines.Add($r2)
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush')
    [void]$row.Children.Add($t)
    $b.Resources['tb'] = $t
    [void]$sp.Children.Add($row)
    if ($hasSubs -and $script:CalOpen.ContainsKey($e.todo.id)) {
        $csubs = @($e.todo.subs); $clines = Get-SubLines $e.todo
        for ($ci = 0; $ci -lt $csubs.Count; $ci++) {
            $line = $clines[$ci]
            $st = [System.Windows.Controls.TextBlock]::new()
            $st.Text = $line; $st.FontSize = $size - 1; $st.TextWrapping = 'Wrap'; $st.Margin = Th 13 1 0 0
            if (-not $inPanel -and $csubs[$ci].kind -ne 'image') { $st.Tag = 'sub|' + $e.todo.id + '|' + $csubs[$ci].id; $st.add_MouseLeftButtonDown($OnCalSubDbl) }
            $st.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
            [void]$sp.Children.Add($st)
        }
    }
    $b.Child = $sp
    return $b
}

$OnCalCellClick = { param($s, $e) $script:CalSel = [datetime]$s.Tag; $script:CalEdit = $null; $script:CalSideOpen = $true; Render-Cal }

function New-CalCell([datetime]$d, $items, [bool]$dim, [int]$maxItems) {
    $cell = [System.Windows.Controls.Border]::new()
    $cell.BorderThickness = Th 0 0 1 1
    $cell.SetResourceReference([System.Windows.Controls.Border]::BorderBrushProperty, 'LineBrush')
    $cell.Padding = Th 6 4 4 4
    $cell.Cursor = $Hand
    $cell.Tag = $d
    $cell.ClipToBounds = $true
    $isToday = $d.Date -eq (Get-Date).Date
    if ($d.Date -eq $script:CalSel.Date) { $cell.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'HoverBrush') }
    else { $cell.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'CardBrush') }
    $cell.add_MouseLeftButtonUp($OnCalCellClick)
    $dock = [System.Windows.Controls.DockPanel]::new()
    $hd = [System.Windows.Controls.TextBlock]::new()
    $hd.FontSize = 12; $hd.Margin = Th 0 0 0 4
    $hd.Text = if ($script:CalMode -eq 'week') { $WeekNames[[int]$d.DayOfWeek] + '  ' + $d.ToString('M/d') } else { [string]$d.Day }
    if ($isToday) { $hd.FontWeight = [System.Windows.FontWeights]::Bold; $hd.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush'); $hd.Text += '  今天' }
    else { $hd.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $(if ($dim) { 'MutedBrush' } else { 'SubBrush' })) }
    [System.Windows.Controls.DockPanel]::SetDock($hd, 'Top')
    [void]$dock.Children.Add($hd)
    $list = [System.Windows.Controls.StackPanel]::new()
    $arr = @(@($items) | Where-Object { $_ })
    $ds = $d.ToString('yyyy-MM-dd')
    if ($script:CalMode -eq 'week') {
        $ord = @{}; $k = 0; foreach ($pp in (Get-ProjectsOrdered $true)) { $ord[$pp.id] = $k; $k++ }
        foreach ($grp in @($arr | Group-Object proj | Sort-Object @{ Expression = { if ($ord.ContainsKey($_.Name)) { $ord[$_.Name] } else { 9999 } } })) {
            $gh = [System.Windows.Controls.TextBlock]::new()
            $gh.Text = Get-ProjName $grp.Name
            $gh.FontSize = 11; $gh.FontWeight = [System.Windows.FontWeights]::SemiBold; $gh.TextWrapping = 'Wrap'; $gh.Margin = Th 0 4 0 2
            $gh.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
            $gpj = Find-Project $grp.Name
            if ($gpj -and $gpj.kind -eq 'project') {
                $gh.Tag = $grp.Name; $gh.ToolTip = '双击修改项目名称'
                $gh.add_MouseLeftButtonDown($OnProjLabelDbl); $gh.add_MouseLeftButtonUp($OnSwallow)
            }
            [void]$list.Children.Add($gh)
            $gn = 0
            foreach ($e in $grp.Group) { $gn++; [void]$list.Children.Add((New-CalItemLine $e $ds 12 $gn)) }
        }
    } else {
        $show = $arr.Count
        if ($maxItems -gt 0 -and $arr.Count -gt $maxItems) { $show = [math]::Max(0, $maxItems - 1) }
        for ($i = 0; $i -lt $show; $i++) { [void]$list.Children.Add((New-CalItemLine $arr[$i] $ds 12 0 $false $true)) }
        if ($show -lt $arr.Count) {
            $more = [System.Windows.Controls.TextBlock]::new()
            $more.Text = '还有 {0} 条' -f ($arr.Count - $show)
            $more.FontSize = 11; $more.Margin = Th 4 0 0 0
            $more.ToolTip = '点这一天，右边列出全部'
            $more.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
            [void]$list.Children.Add($more)
        }
    }
    [void]$dock.Children.Add($list)
    if ($dim) { $dock.Opacity = 0.55 }
    $cell.Child = $dock
    return $cell
}

# ---------- 右侧：当天清单 ----------
function New-PanelTitle([string]$text, [double]$size) {
    $h = [System.Windows.Controls.TextBlock]::new()
    $h.Text = $text; $h.FontSize = $size; $h.FontWeight = [System.Windows.FontWeights]::SemiBold; $h.Margin = Th 0 0 0 8; $h.TextWrapping = 'Wrap'
    $h.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush')
    return $h
}
function New-PanelLabel([string]$text) {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $text; $t.FontSize = 12; $t.Margin = Th 0 10 0 4
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
    return $t
}
function New-LinkText([string]$text, $tag, [scriptblock]$h) {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $text; $t.Cursor = $Hand; $t.Tag = $tag; $t.Margin = Th 0 0 14 0
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'AccentBrush')
    $t.add_MouseLeftButtonUp($h)
    return $t
}

function Render-CalDay($map) {
    $p = $script:CalDetail
    $d = $script:CalSel
    [void]$p.Children.Add((New-PanelTitle ($d.ToString('yyyy年M月d日') + '  ' + $WeekNames[[int]$d.DayOfWeek]) 15))
    $ds = $d.ToString('yyyy-MM-dd')
    $m2 = if ($map.ContainsKey($ds)) { $map } else { Build-CalMap $d $d }
    $items = @(@($m2[$ds]) | Where-Object { $_ })
    if ($items.Count -eq 0) {
        $em = [System.Windows.Controls.TextBlock]::new()
        $em.Text = '这一天没有内容。'
        $em.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
        [void]$p.Children.Add($em)
    }
    $ord = @{}; $k = 0; foreach ($pp in (Get-ProjectsOrdered $true)) { $ord[$pp.id] = $k; $k++ }
    foreach ($g in @($items | Group-Object proj | Sort-Object @{ Expression = { if ($ord.ContainsKey($_.Name)) { $ord[$_.Name] } else { 9999 } } })) {
        $gt = New-PanelLabel ((Get-ProjName $g.Name) + '（' + $g.Count + '）')
        $gt.FontWeight = [System.Windows.FontWeights]::SemiBold
        $gpj = Find-Project $g.Name
        if ($gpj -and $gpj.kind -eq 'project') {
            $gt.Cursor = $Hand; $gt.Tag = $g.Name; $gt.ToolTip = '点一下修改项目名称'
            $gt.add_MouseLeftButtonUp($OnRenameProj)
        }
        [void]$p.Children.Add($gt)
        $gn = 0
        foreach ($e in $g.Group) { $gn++; [void]$p.Children.Add((New-CalItemLine $e $ds 13 $gn $true)) }
    }
    $tip = [System.Windows.Controls.TextBlock]::new()
    $tip.Text = '点上面任意一条可以修改；下面可以给这天补记一条工作（先记在「日常」，记下后可以改项目）。'
    $tip.TextWrapping = 'Wrap'; $tip.FontSize = 11; $tip.Margin = Th 0 14 0 4
    $tip.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
    [void]$p.Children.Add($tip)
    $box = New-HintBox $ds ('给 ' + $d.ToString('M月d日') + ' 补记一条工作，回车保存') $false
    $box[1].add_KeyDown({
        param($s, $e)
        if ($e.Key -ne [System.Windows.Input.Key]::Return -or -not $s.Text.Trim()) { return }
        $e.Handled = $true
        [void]$script:Work.Add((New-WorkEntry $s.Text.Trim() ([string]$s.Tag) '' 'cat-daily' ''))
        Save-Work; Render
    })
    $box[1].add_TextChanged($OnSubAddText)
    [void]$p.Children.Add($box[0])
}

# ---------- 右侧：编辑一件事项 ----------
function Commit-TextBox($tb, [scriptblock]$apply) {
    $tb.add_LostKeyboardFocus($apply)
    $tb.add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return -and -not $s.AcceptsReturn) { $e.Handled = $true; [System.Windows.Input.Keyboard]::ClearFocus() } })
}

function New-EditBox([string]$text, [bool]$multi) {
    $tb = [System.Windows.Controls.TextBox]::new()
    $tb.Text = $text; $tb.TextWrapping = 'Wrap'; $tb.Padding = Th 4 3 4 3
    if ($multi) { $tb.AcceptsReturn = $true; $tb.MinHeight = 44 }
    $tb.SetResourceReference([System.Windows.Controls.Control]::BackgroundProperty, 'InputBgBrush')
    $tb.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'TextBrush')
    $tb.SetResourceReference([System.Windows.Controls.Control]::BorderBrushProperty, 'InputLineBrush')
    return $tb
}

function New-ProjCombo([string]$current, [scriptblock]$onChange, [string]$tag) {
    $cb = [System.Windows.Controls.ComboBox]::new()
    $cb.Tag = $tag
    $list = @(Get-ProjectsOrdered $true)
    $script:ComboProjIds = @{}
    foreach ($p in $list) { $label = $p.text; if ($p.status -eq 'closed') { $label += '（已结束）' }; [void]$cb.Items.Add($label) }
    $ids = @($list | ForEach-Object { $_.id })
    $cb.SelectedIndex = [array]::IndexOf($ids, $current)
    $cb.Resources['ids'] = $ids
    $cb.add_SelectionChanged($onChange)
    return $cb
}

function Get-ComboProj($cb) { $ids = $cb.Resources['ids']; if ($cb.SelectedIndex -ge 0) { return [string]$ids[$cb.SelectedIndex] }; return $null }

function Render-CalTodoEditor($it) {
    $p = $script:CalDetail
    [void]$p.Children.Add((New-LinkText $(if ($script:CalMode -eq 'week' -or $script:CalMode -eq 'month') { '‹ 返回当天清单' } else { '× 关闭' }) '' { $script:CalEdit = $null; Render-Cal }))
    $p.Children[$p.Children.Count - 1].Margin = Th 0 0 0 10
    [void]$p.Children.Add((New-PanelTitle '待办' 15))

    $title = New-EditBox $it.text $true
    $title.Tag = $it.id
    Commit-TextBox $title { param($s, $e) $t = Find-Todo ([string]$s.Tag); if ($t -and $s.Text.Trim() -and $s.Text.Trim() -ne $t.text) { $t.text = $s.Text.Trim(); Save-Items; Render } }
    [void]$p.Children.Add($title)

    [void]$p.Children.Add((New-PanelLabel '项目'))
    $pc = New-ProjCombo $it.project { param($s, $e) $t = Find-Todo ([string]$s.Tag); $np = Get-ComboProj $s; if ($t -and $np -and $np -ne $t.project) { $t.project = $np; foreach ($w in @($script:Work | Where-Object { $_.src -eq $t.id })) { $w.proj = $np }; if ($np -eq 'cat-reimb' -and -not $t.rstate) { $t.rstate = '待报' }; Save-Items; Save-Work; Render } } $it.id
    [void]$p.Children.Add($pc)
    Add-RenameLink $p $it.project

    $row = [System.Windows.Controls.WrapPanel]::new()
    $row.Margin = Th 0 10 0 0
    $done = [System.Windows.Controls.CheckBox]::new()
    $done.Content = '已完成'; $done.IsChecked = [bool]$it.done; $done.Tag = $it.id; $done.Margin = Th 0 0 16 0
    $done.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'TextBrush')
    $done.add_Click({ param($s, $e) Set-Done ([string]$s.Tag) ([bool]$s.IsChecked) })
    [void]$row.Children.Add($done)
    $sdTxt = if ($it.start) { '开始：' + (Parse-Day $it.start).ToString('M月d日') } else { '设开始日期' }
    [void]$row.Children.Add((New-LinkText $sdTxt ($it.id + '|set') $OnTodoStart))
    if ($it.start) { [void]$row.Children.Add((New-LinkText '清除' ($it.id + '|clear') $OnTodoStart)) }
    [void]$p.Children.Add($row)
    $row2 = [System.Windows.Controls.WrapPanel]::new()
    $row2.Margin = Th 0 6 0 0
    $crd = Parse-Dt $it.created
    [void]$row2.Children.Add((New-LinkText ('记于 ' + $(if ($crd) { $crd.ToString('M月d日 HH:mm') } else { '未知' }) + '（改）') $it.id $OnEditCreated))
    if ($it.done) {
        $dna = Parse-Dt $it.doneAt
        [void]$row2.Children.Add((New-LinkText ('完成于 ' + $(if ($dna) { $dna.ToString('M月d日 HH:mm') } else { '未知' }) + '（改）') $it.id $OnEditDoneAt))
    }
    [void]$p.Children.Add($row2)

    [void]$p.Children.Add((New-PanelLabel '提醒 / 截止时间'))
    $tw = [System.Windows.Controls.WrapPanel]::new()
    foreach ($ts in @($it.times)) {
        $dt = Parse-Dt $ts
        $lab = $dt.ToString('M月d日 HH:mm') + $(if ($ts -eq $it.due) { '（截止）' } else { '' })
        $tl = New-LinkText $lab ($it.id + '|' + $ts) $OnEditDue
        $tl.Margin = Th 0 0 3 0; $tl.ToolTip = '点一下修改这个时间'
        [void]$tw.Children.Add($tl)
        $tx = New-LinkText '×' ($it.id + '|del|' + $ts) { param($s, $e) $q = ([string]$s.Tag).Split([char[]]'|', 3); Set-Due $q[0] $q[1] $q[2] }
        $tx.ToolTip = '删除这个时间'
        [void]$tw.Children.Add($tx)
    }
    [void]$tw.Children.Add((New-LinkText '＋加时间' ($it.id + '|pick|') { param($s, $e) $q = ([string]$s.Tag).Split([char[]]'|', 3); Set-Due $q[0] $q[1] $q[2] }))
    [void]$p.Children.Add($tw)

    if (Test-Reimb $it) {
        [void]$p.Children.Add((New-PanelLabel '报销金额（元）和状态'))
        $rr = [System.Windows.Controls.DockPanel]::new()
        $amt = New-EditBox ([string]$it.amount) $false
        $amt.Width = 110; $amt.Tag = $it.id; $amt.Margin = Th 0 0 10 0
        Commit-TextBox $amt { param($s, $e) $t = Find-Todo ([string]$s.Tag); $v = 0.0; if ($t -and $s.Text.Trim() -ne $t.amount -and (-not $s.Text.Trim() -or [double]::TryParse($s.Text.Trim(), [ref]$v))) { $t.amount = $s.Text.Trim(); Save-Items; Render } }
        [System.Windows.Controls.DockPanel]::SetDock($amt, 'Left')
        [void]$rr.Children.Add($amt)
        $rs = [System.Windows.Controls.ComboBox]::new()
        foreach ($x in $ReimbStates) { [void]$rs.Items.Add($x) }
        $rs.SelectedItem = if ($it.rstate) { $it.rstate } else { '待报' }
        $rs.Tag = $it.id; $rs.Width = 100; $rs.HorizontalAlignment = 'Left'
        $rs.add_SelectionChanged({ param($s, $e) $t = Find-Todo ([string]$s.Tag); if ($t -and [string]$s.SelectedItem -ne $t.rstate) { $t.rstate = [string]$s.SelectedItem; if ($t.rstate -eq '已到账' -and -not $t.done) { $t.done = $true; $t.doneAt = (Get-Date).ToString($Fmt) }; Save-Items; Render } })
        [void]$rr.Children.Add($rs)
        [void]$p.Children.Add($rr)
    }

    [void]$p.Children.Add((New-PanelLabel '具体事项（勾掉一项，当天自动记一条进展）'))
    $i = 0
    foreach ($x in @($it.subs)) {
        $i++
        $r = [System.Windows.Controls.DockPanel]::new()
        $r.Margin = Th 0 0 0 4
        $del = New-LinkText '×' ($it.id + '|' + $x.id) $OnSubDelete
        $del.Margin = Th 6 0 0 0
        [System.Windows.Controls.DockPanel]::SetDock($del, 'Right'); [void]$r.Children.Add($del)
        $cb = [System.Windows.Controls.CheckBox]::new()
        $cb.IsChecked = [bool]$x.done; $cb.Tag = $it.id + '|' + $x.id; $cb.Margin = Th 0 3 6 0; $cb.VerticalAlignment = 'Top'
        $cb.add_Click($OnSubCheck)
        [System.Windows.Controls.DockPanel]::SetDock($cb, 'Left'); [void]$r.Children.Add($cb)
        $num = New-NumLabel $i
        [System.Windows.Controls.DockPanel]::SetDock($num, 'Left'); [void]$r.Children.Add($num)
        if ($x.kind -eq 'image') {
            [void]$r.Children.Add((New-Thumb $x.file ($it.id + '|' + $x.id) $OnSubImgOpen 140))
        } else {
            $tb = New-EditBox $x.text $false
            $tb.Tag = $it.id + '|' + $x.id
            Commit-TextBox $tb { param($s, $e) $q = ([string]$s.Tag).Split('|'); $t = Find-Todo $q[0]; if (-not $t) { return }; $y = Find-Sub $t $q[1]; if ($y -and $s.Text.Trim() -and $s.Text.Trim() -ne $y.text) { $old = $y.text; $y.text = $s.Text.Trim(); Save-Items; Sync-AutoText $t $y $old; Render } }
            [void]$r.Children.Add($tb)
        }
        [void]$p.Children.Add($r)
    }
    $add = New-HintBox $it.id '加一条具体事项，回车' $false
    $add[1].add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return -and $s.Text.Trim()) { $e.Handled = $true; Add-Sub ([string]$s.Tag) $s.Text } })
    $add[1].add_TextChanged($OnSubAddText)
    [void]$p.Children.Add($add[0])

    # 进展记到哪一天：打开这件事时默认是周/月视图里点的那天（报告、项目页里默认今天），可以点"换日期"改
    if ($script:ProgDateFor -ne $it.id -or -not $script:ProgDate) {
        $script:ProgDateFor = $it.id
        $script:ProgDate = if ($script:CalMode -eq 'week' -or $script:CalMode -eq 'month') { $script:CalSel.Date } else { (Get-Date).Date }
    }
    $ds = $script:ProgDate.ToString('yyyy-MM-dd')
    [void]$p.Children.Add((New-PanelLabel '每日进展'))
    $pdRow = [System.Windows.Controls.WrapPanel]::new()
    $pdRow.Margin = Th 0 0 0 4
    $pdl = [System.Windows.Controls.TextBlock]::new()
    $pdl.Text = '记到：' + $script:ProgDate.ToString('M月d日') + ' ' + $WeekNames[[int]$script:ProgDate.DayOfWeek] + $(if ($script:ProgDate -eq (Get-Date).Date) { '（今天）' } else { '' })
    $pdl.Margin = Th 0 0 12 0
    $pdl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush')
    [void]$pdRow.Children.Add($pdl)
    [void]$pdRow.Children.Add((New-LinkText '换日期' '' {
        $d = Show-DatePickerDialog $script:ProgDate.AddHours(9) '进展记到哪一天'
        if ($d) { $script:ProgDate = $d.Date; Render-Cal }
    }))
    if ($script:ProgDate -ne (Get-Date).Date) { [void]$pdRow.Children.Add((New-LinkText '改回今天' '' { $script:ProgDate = (Get-Date).Date; Render-Cal })) }
    [void]$p.Children.Add($pdRow)
    $pg = New-HintBox ($it.id + '|' + $ds) ('这天在这件事上做了什么，回车记下') $false
    $pg[1].add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Return -and $s.Text.Trim()) { $e.Handled = $true; $q = ([string]$s.Tag).Split('|'); $t = Find-Todo $q[0]; if ($t) { Add-Progress $t $s.Text $q[1]; Render } } })
    $pg[1].add_TextChanged($OnSubAddText)
    [void]$p.Children.Add($pg[0])
    $prog = @(Get-TodoProgress $it)
    [array]::Reverse($prog)
    foreach ($w in $prog) {
        $r = [System.Windows.Controls.DockPanel]::new()
        $r.Margin = Th 0 4 0 0
        $del = New-LinkText '×' $w.id { param($s, $e) $w0 = Find-Work ([string]$s.Tag); if ($w0) { [void]$script:Work.Remove($w0); Save-Work; Render } }
        $del.Margin = Th 6 0 0 0
        [System.Windows.Controls.DockPanel]::SetDock($del, 'Right'); [void]$r.Children.Add($del)
        $dtxt = (Parse-Day $w.date).ToString('M月d日') + $(if ((Test-ShowTime) -and $w.time) { ' ' + $w.time } else { '' })
        $dl = New-LinkText $dtxt $w.id $OnWorkMenuDate
        $dl.FontSize = 12; $dl.VerticalAlignment = 'Center'; $dl.Margin = Th 0 0 8 0
        $dl.ToolTip = '点一下修改这条进展的日期和时间'
        [System.Windows.Controls.DockPanel]::SetDock($dl, 'Left'); [void]$r.Children.Add($dl)
        $tb = New-EditBox $w.text $false
        $tb.Tag = $w.id
        Commit-TextBox $tb { param($s, $e) $w0 = Find-Work ([string]$s.Tag); if ($w0 -and $s.Text.Trim() -and $s.Text.Trim() -ne $w0.text) { $w0.text = $s.Text.Trim(); Save-Work; Render } }
        [void]$r.Children.Add($tb)
        [void]$p.Children.Add($r)
    }
    if ($prog.Count -eq 0) {
        $em = [System.Windows.Controls.TextBlock]::new()
        $em.Text = '还没有记过进展。'; $em.FontSize = 12; $em.Margin = Th 0 4 0 0
        $em.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
        [void]$p.Children.Add($em)
    }
}

# ---------- 右侧：编辑一条工作记录 ----------
function Render-CalWorkEditor($w) {
    $p = $script:CalDetail
    [void]$p.Children.Add((New-LinkText $(if ($script:CalMode -eq 'week' -or $script:CalMode -eq 'month') { '‹ 返回当天清单' } else { '× 关闭' }) '' { $script:CalEdit = $null; Render-Cal }))
    $p.Children[$p.Children.Count - 1].Margin = Th 0 0 0 10
    [void]$p.Children.Add((New-PanelTitle '工作记录' 15))
    $wh = [System.Windows.Controls.TextBlock]::new()
    $wh.Text = '工作记录就是某天做了什么。属于哪件待办的，就是那件待办的进展；不属于任何待办也可以，只记在项目下。'
    $wh.TextWrapping = 'Wrap'; $wh.FontSize = 11; $wh.Margin = Th 0 0 0 8
    $wh.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
    [void]$p.Children.Add($wh)
    $tb = New-EditBox $w.text $true
    $tb.Tag = $w.id
    Commit-TextBox $tb { param($s, $e) $w0 = Find-Work ([string]$s.Tag); if ($w0 -and $s.Text.Trim() -and $s.Text.Trim() -ne $w0.text) { $w0.text = $s.Text.Trim(); Save-Work; Render } }
    [void]$p.Children.Add($tb)

    [void]$p.Children.Add((New-PanelLabel '日期和时间'))
    $dr = [System.Windows.Controls.WrapPanel]::new()
    $dtx = (Parse-Day $w.date).ToString('yyyy年M月d日') + $(if ($w.time) { ' ' + $w.time } else { '' })
    $dl = [System.Windows.Controls.TextBlock]::new(); $dl.Text = $dtx; $dl.Margin = Th 0 0 14 0
    $dl.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush')
    [void]$dr.Children.Add($dl)
    [void]$dr.Children.Add((New-LinkText '改日期…' $w.id { param($s, $e) Set-WorkDate ([string]$s.Tag); $w0 = Find-Work ([string]$s.Tag); if ($w0) { $script:CalSel = Parse-Day $w0.date; Render-Cal } }))
    [void]$p.Children.Add($dr)

    [void]$p.Children.Add((New-PanelLabel '项目'))
    [void]$p.Children.Add((New-ProjCombo $w.proj { param($s, $e) $w0 = Find-Work ([string]$s.Tag); $np = Get-ComboProj $s; if ($w0 -and $np -and $np -ne $w0.proj) { $w0.proj = $np; if ($w0.src) { $t = Find-Todo $w0.src; if ($t -and $t.project -ne $np) { $w0.src = '' } }; Save-Work; Render } } $w.id))
    Add-RenameLink $p $w.proj

    [void]$p.Children.Add((New-PanelLabel '属于哪件待办（可以不选）'))
    $tc = [System.Windows.Controls.ComboBox]::new()
    $cands = @($script:Items | Where-Object { (Get-TodoProj $_) -eq $w.proj } | Sort-Object @{ Expression = { [int][bool]$_.done } }, @{ Expression = { $_.created }; Descending = $true } | Select-Object -First 40)
    if ($w.src -and -not @($cands | Where-Object { $_.id -eq $w.src }).Count) { $st1 = Find-Todo $w.src; if ($st1) { $cands = @($st1) + $cands } }
    [void]$tc.Items.Add('（不属于哪件待办）')
    foreach ($t in $cands) { [void]$tc.Items.Add($(if ($t.text.Length -gt 30) { $t.text.Substring(0, 30) + '…' } else { $t.text }) + $(if ($t.done) { '（已完成）' } else { '' })) }
    $ids = @('') + @($cands | ForEach-Object { $_.id })
    $tc.SelectedIndex = [math]::Max(0, [array]::IndexOf($ids, [string]$w.src))
    $tc.Resources['ids'] = $ids; $tc.Tag = $w.id
    $tc.add_SelectionChanged({ param($s, $e) $w0 = Find-Work ([string]$s.Tag); $ids0 = $s.Resources['ids']; if ($w0 -and $s.SelectedIndex -ge 0) { $nid = [string]$ids0[$s.SelectedIndex]; if ($nid -ne $w0.src) { $w0.src = $nid; Save-Work; Render } } })
    [void]$p.Children.Add($tc)
    if ($w.src -and (Find-Todo $w.src)) {
        $ol = New-LinkText '打开这件待办 ›' $w.src { param($s, $e) $script:CalEdit = @{ kind = 'todo'; id = [string]$s.Tag }; Render-Cal }
        $ol.Margin = Th 0 6 0 0
        [void]$p.Children.Add($ol)
    }

    $dl2 = New-LinkText '删除这条记录' $w.id { param($s, $e) $w0 = Find-Work ([string]$s.Tag); if ($w0 -and ([System.Windows.MessageBox]::Show($script:CalWin, "删除这条记录？`n`n" + $w0.text, '桌面待办', 'YesNo', 'Question') -eq 'Yes')) { [void]$script:Work.Remove($w0); Save-Work; $script:CalEdit = $null; Render } }
    $dl2.Margin = Th 0 18 0 0
    $dl2.Foreground = Brush '#D0402F'
    [void]$p.Children.Add($dl2)
}

# 右边一栏：平时收起，点了某一天或某一条才打开
function Set-CalSide([bool]$open) {
    $script:CalSideOpen = $open
    try {
        $cd = $script:CalPanel.ColumnDefinitions[1]
        if ($open) { $cd.Width = [System.Windows.GridLength]::new(360); $script:CalSide.Visibility = 'Visible' }
        else { $cd.Width = [System.Windows.GridLength]::new(0); $script:CalSide.Visibility = 'Collapsed' }
    } catch {}
}
function Render-CalDetail($map) {
    Set-CalSide ([bool]$script:CalSideOpen)
    $script:CalDetail.Children.Clear()
    if ($script:CalEdit) {
        if ($script:CalEdit.kind -eq 'todo') { $t = Find-Todo $script:CalEdit.id; if ($t) { Render-CalTodoEditor $t; return } }
        else { $w = Find-Work $script:CalEdit.id; if ($w) { Render-CalWorkEditor $w; return } }
        $script:CalEdit = $null
    }
    Render-CalDay $map
}

function Get-CalText {
    $r = Get-CalRange
    $map = Build-CalMap $r[0] $r[1]
    $lines = New-Object System.Collections.Generic.List[string]
    if ($script:CalMode -eq 'week') { $lines.Add(('{0}—{1} 清单' -f $r[0].ToString('M月d日'), $r[1].ToString('M月d日'))) }
    else { $lines.Add(('{0} 清单' -f $r[0].ToString('yyyy年M月'))) }
    for ($d = $r[0]; $d -le $r[1]; $d = $d.AddDays(1)) {
        $ds = $d.ToString('yyyy-MM-dd')
        if (-not $map.ContainsKey($ds)) { continue }
        $lines.Add('')
        $lines.Add(($d.ToString('M月d日') + ' ' + $WeekNames[[int]$d.DayOfWeek]))
        $cn = 0
        foreach ($e in $map[$ds]) {
            $cn++
            $head = @(); if ($e.time -and (Test-ShowTime)) { $head += $e.time }; if ($e.tag) { $head += $e.tag }
            $hd0 = if ($head.Count -gt 0) { '【' + ($head -join ' ') + '】' } else { '' }
            $lines.Add(('  {0}. {1}{2}（{3}）' -f $cn, $hd0, $e.text, (Get-ProjName $e.proj)))
            if ($e.todo -and $e.kind -ne 'ongoing' -and $script:CalOpen.ContainsKey($e.todo.id)) { foreach ($l in (Get-SubLines $e.todo)) { $lines.Add('      ' + $l) } }
        }
    }
    if ($lines.Count -eq 1) { $lines.Add(''); $lines.Add('（这段时间没有内容）') }
    return ($lines -join "`r`n")
}

function Render-CalGrid {
    $r = Get-CalRange
    $isWeek = $script:CalMode -eq 'week'
    $script:CalTitle.Text = if ($isWeek) { '{0} — {1}' -f $r[0].ToString('yyyy年M月d日'), $r[1].ToString('M月d日') } else { $r[0].ToString('yyyy年M月') }
    $g = $script:CalGrid
    $g.Children.Clear(); $g.RowDefinitions.Clear(); $g.ColumnDefinitions.Clear()
    for ($c = 0; $c -lt 7; $c++) { $cd = [System.Windows.Controls.ColumnDefinition]::new(); $cd.Width = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star); [void]$g.ColumnDefinitions.Add($cd) }
    if ($isWeek) {
        $script:CalScroll.VerticalScrollBarVisibility = 'Auto'
        $g.MinHeight = [math]::Max(0, $script:CalScroll.ActualHeight - 2)
        $map = Build-CalMap $r[0] $r[1]
        $rd = [System.Windows.Controls.RowDefinition]::new(); $rd.Height = [System.Windows.GridLength]::new(1, [System.Windows.GridUnitType]::Star); [void]$g.RowDefinitions.Add($rd)
        for ($i = 0; $i -lt 7; $i++) {
            $d = $r[0].AddDays($i)
            $cell = New-CalCell $d $map[$d.ToString('yyyy-MM-dd')] $false 0
            [System.Windows.Controls.Grid]::SetColumn($cell, $i)
            [void]$g.Children.Add($cell)
        }
    } else {
        $start = Get-WeekStart $r[0]
        $end = (Get-WeekStart $r[1]).AddDays(6)
        $map = Build-CalMap $start $end
        $hdr = [System.Windows.Controls.RowDefinition]::new(); $hdr.Height = [System.Windows.GridLength]::Auto; [void]$g.RowDefinitions.Add($hdr)
        $names = @('周一', '周二', '周三', '周四', '周五', '周六', '周日')
        for ($c = 0; $c -lt 7; $c++) {
            $t = [System.Windows.Controls.TextBlock]::new()
            $t.Text = $names[$c]; $t.FontSize = 12; $t.Margin = Th 6 2 0 4
            $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
            [System.Windows.Controls.Grid]::SetColumn($t, $c)
            [void]$g.Children.Add($t)
        }
        $weeks = [int]((($end - $start).TotalDays + 1) / 7)
        $script:CalScroll.VerticalScrollBarVisibility = 'Disabled'
        $g.MinHeight = 0
        $H = [double]$script:CalScroll.ActualHeight
        if ($H -lt 150) { $H = [math]::Max(360.0, [double]$script:CalWin.ActualHeight - 180) }
        $script:CalLastH = $H
        $H -= 24
        $lineH = 21.0; $over = 30.0
        $need = @()
        for ($w = 0; $w -lt $weeks; $w++) {
            $mx = 0
            for ($c = 0; $c -lt 7; $c++) { $k0 = $start.AddDays($w * 7 + $c).ToString('yyyy-MM-dd'); if ($map.ContainsKey($k0)) { $mx = [math]::Max($mx, @($map[$k0] | Where-Object { $_ }).Count) } }
            $need += ($over + $lineH * [math]::Max(1, [math]::Min($mx, 8)))
        }
        $sumNeed = 0.0; foreach ($nd in $need) { $sumNeed += $nd }
        $hs = @(foreach ($nd in $need) { if ($sumNeed -le $H) { $nd + ($H - $sumNeed) / $weeks } else { [math]::Max($over + $lineH, $H * $nd / $sumNeed) } })
        for ($w = 0; $w -lt $weeks; $w++) {
            $rd = [System.Windows.Controls.RowDefinition]::new(); $rd.Height = [System.Windows.GridLength]::new($hs[$w], [System.Windows.GridUnitType]::Star); [void]$g.RowDefinitions.Add($rd)
            $fit = [math]::Max(1, [int][math]::Floor(($hs[$w] - $over) / $lineH))
            for ($c = 0; $c -lt 7; $c++) {
                $d = $start.AddDays($w * 7 + $c)
                $cell = New-CalCell $d $map[$d.ToString('yyyy-MM-dd')] ($d.Month -ne $r[0].Month) $fit
                [System.Windows.Controls.Grid]::SetRow($cell, $w + 1)
                [System.Windows.Controls.Grid]::SetColumn($cell, $c)
                [void]$g.Children.Add($cell)
            }
        }
    }
    Render-CalDetail $map
    $cnt = 0
    for ($d = $r[0]; $d -le $r[1]; $d = $d.AddDays(1)) { $k = $d.ToString('yyyy-MM-dd'); if ($map.ContainsKey($k)) { $cnt += @($map[$k] | Where-Object { $_.kind -ne 'ongoing' }).Count } }
    $script:CalCount.Text = if ($isWeek) { '本周 {0} 条' -f $cnt } else { '本月 {0} 条' -f $cnt }
    # 全部展开/收起：看当前范围里有具体事项的事项是不是都展开了
    $withSubs = @($map.Values | ForEach-Object { $_ } | Where-Object { $_.todo -and @($_.todo.subs).Count -gt 0 } | ForEach-Object { $_.todo.id } | Select-Object -Unique)
    $script:CalVisibleTodos = $withSubs
    $closed = @($withSubs | Where-Object { -not $script:CalOpen.ContainsKey($_) })
    $script:CalExpandBtn.Content = if ($withSubs.Count -gt 0 -and $closed.Count -eq 0) { '全部收起' } else { '全部展开' }
}

# ---------- 报告页 ----------
function Update-RepHighlight($panel) {
    foreach ($c in $panel.Children) {
        if ($c -is [System.Windows.Controls.Border] -and $c.Tag) {
            $q = ([string]$c.Tag).Split('|')
            $k = if ($q[0] -eq 'sub') { 'todo' } else { $q[0] }
            if ($script:CalEdit -and $script:CalEdit.kind -eq $k -and $script:CalEdit.id -eq $q[1]) { $c.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'HoverBrush') }
            else { $c.Background = [System.Windows.Media.Brushes]::Transparent }
        }
    }
}
# 单击：右边打开修改；不重画报告，方便接着双击
$OnRepLineClick = {
    param($s, $e)
    if ($script:CalInline) { return }
    $q = ([string]$s.Tag).Split('|')
    if ($q[0] -eq 'proj') { return }
    $k = if ($q[0] -eq 'sub') { 'todo' } else { $q[0] }
    $script:CalEdit = @{ kind = $k; id = $q[1] }
    if ($k -eq 'work') { $w = Find-Work $q[1]; if ($w) { $d = Parse-Day $w.date; if ($d) { $script:CalSel = $d } } }
    Flush-CalFocus
    Update-RepHighlight $s.Parent
    Render-SideEditor '点左边任意一条，在这里修改内容、日期和时间。'
}
# 双击：就地改文字；项目标题双击改项目名
$OnRepLineDbl = {
    param($s, $e)
    if ($e.ClickCount -ne 2) { return }
    $e.Handled = $true
    $q = ([string]$s.Tag).Split('|')
    if ($q[0] -eq 'proj') { Rename-ProjectById $q[1]; return }
    if ($s.Child) { Start-SwapEdit $s.Child ([string]$s.Tag) }
}

# 把报告一行一行画出来；能对应到事项或进展的行，点一下在右边修改
function Show-ReportLines($lines, $panel) {
    $panel.Children.Clear()
    foreach ($x in $lines) {
        $b = [System.Windows.Controls.Border]::new()
        $b.CornerRadius = [System.Windows.CornerRadius]::new(3)
        $b.Padding = Th 4 2 4 2
        $t = [System.Windows.Controls.TextBlock]::new()
        $t.Text = $x.text; $t.TextWrapping = 'Wrap'
        $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'TextBrush')
        switch ($x.lv) {
            'title' { $t.FontSize = 18; $t.FontWeight = [System.Windows.FontWeights]::Bold; $b.Margin = Th 0 0 0 10 }
            'h1'    { $t.FontSize = 15; $t.FontWeight = [System.Windows.FontWeights]::SemiBold; $b.Margin = Th 0 12 0 2 }
            'h2'    { $t.FontSize = 14; $t.FontWeight = [System.Windows.FontWeights]::SemiBold; $b.Margin = Th 10 4 0 2 }
            'item'  { $t.FontSize = 13.5; $b.Margin = Th 22 1 0 1 }
            'sub'   { $t.FontSize = 12.5; $b.Margin = Th 46 0 0 1; $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush') }
            default { $t.FontSize = 13; $b.Margin = Th 22 4 0 1 }
        }
        if ($x.ref) {
            $b.Cursor = $Hand
            $b.Tag = $x.ref
            $b.ToolTip = if ($x.ref.StartsWith('proj|')) { '双击修改项目名称' } else { '点一下在右边修改；双击直接改文字' }
            $q = $x.ref.Split('|')
            $k0 = if ($q[0] -eq 'sub') { 'todo' } else { $q[0] }
            if ($script:CalEdit -and $script:CalEdit.kind -eq $k0 -and $script:CalEdit.id -eq $q[1]) { $b.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'HoverBrush') }
            else { $b.Background = [System.Windows.Media.Brushes]::Transparent }
            $b.add_MouseLeftButtonUp($OnRepLineClick)
            $b.add_MouseLeftButtonDown($OnRepLineDbl)
        }
        $b.Child = $t
        [void]$panel.Children.Add($b)
    }
}

# 报告页、项目页右侧：正在改的那一条；没选中时给个提示
function Render-SideEditor([string]$tip) {
    $p = $script:CalDetail
    $p.Children.Clear()
    if ($script:CalEdit) {
        if ($script:CalEdit.kind -eq 'todo') { $t = Find-Todo $script:CalEdit.id; if ($t) { Render-CalTodoEditor $t; return } }
        else { $w = Find-Work $script:CalEdit.id; if ($w) { Render-CalWorkEditor $w; return } }
        $script:CalEdit = $null
    }
    $tb = [System.Windows.Controls.TextBlock]::new()
    $tb.Text = $tip; $tb.TextWrapping = 'Wrap'
    $tb.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
    [void]$p.Children.Add($tb)
}

function Render-Report {
    $r = Get-PeriodRange $script:RepKind $script:RepAnchor
    $script:RepTitle.Text = $r[2]
    $lines = Build-ReportLines $r[0] $r[1] $r[2] ([bool]$script:RepSubs.IsChecked)
    $script:RepLines = $lines
    Show-ReportLines $lines $script:RepView
    Render-SideEditor '点左边报告里任意一条待办或进展，在这里修改内容、日期和时间；点项目标题可以改项目名。改完报告、周月视图、小窗和另一台电脑都会一起更新。'
    $isCustom = $script:RepKind -eq 'custom'
    $script:RepCustom.Visibility = if ($isCustom) { 'Visible' } else { 'Collapsed' }
    $script:RepNav.Visibility = if ($isCustom) { 'Collapsed' } else { 'Visible' }
    foreach ($b in $script:RepKindBtns) {
        $on = [string]$b.Tag -eq $script:RepKind
        $b.FontWeight = if ($on) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
        $b.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $(if ($on) { 'AccentBrush' } else { 'SubBrush' }))
    }
}

function Move-Report([int]$dir) {
    switch ($script:RepKind) {
        'day' { $script:RepAnchor = $script:RepAnchor.AddDays($dir) }
        'week' { $script:RepAnchor = $script:RepAnchor.AddDays(7 * $dir) }
        'month' { $script:RepAnchor = $script:RepAnchor.AddMonths($dir) }
        'quarter' { $script:RepAnchor = $script:RepAnchor.AddMonths(3 * $dir) }
        'year' { $script:RepAnchor = $script:RepAnchor.AddYears($dir) }
    }
    Render-Report
}

function Save-ReportDocx {
    $r = Get-PeriodRange $script:RepKind $script:RepAnchor
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Word 文档|*.docx'
    $dlg.FileName = ($r[2] -replace '[\\/:*?"<>|]', '') + '.docx'
    if (-not $dlg.ShowDialog($script:CalWin)) { return }
    try {
        Export-ReportDocx $script:RepLines $dlg.FileName
        Start-Process $dlg.FileName
    } catch {
        Write-Log ('导出 Word 失败：' + $_.Exception.ToString())
        [System.Windows.MessageBox]::Show($script:CalWin, '导出失败：' + $_.Exception.Message + "`n`n如果这个文件正在 Word 里打开着，先关掉再导出。", '桌面待办') | Out-Null
    }
}

# ---------- 项目页 ----------
function Render-ProjectsTab {
    $lb = $script:ProjList
    $lb.Children.Clear()
    $month0 = [datetime]::new((Get-Date).Year, (Get-Date).Month, 1)
    foreach ($p in @(Get-ProjectsOrdered $true)) {
        $open = @($script:Items | Where-Object { $_.project -eq $p.id -and -not $_.done }).Count
        $all = @($script:Items | Where-Object { $_.project -eq $p.id }).Count
        $last = @($script:Work | Where-Object { $_.proj -eq $p.id } | Sort-Object date -Descending | Select-Object -First 1)
        $b = [System.Windows.Controls.Border]::new()
        $b.Padding = Th 10 7 10 7; $b.Margin = Th 0 0 0 4; $b.CornerRadius = [System.Windows.CornerRadius]::new(5); $b.Cursor = $Hand; $b.Tag = $p.id
        if ($script:ProjSel -eq $p.id) { $b.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'HoverBrush') }
        else { $b.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'CardBrush') }
        $sp = [System.Windows.Controls.StackPanel]::new()
        $t1 = [System.Windows.Controls.TextBlock]::new()
        $t1.Text = $p.text + $(if ($p.kind -eq 'category') { '　（固定类别）' } elseif ($p.status -eq 'closed') { '　（已结束）' } else { '' })
        $t1.FontWeight = [System.Windows.FontWeights]::SemiBold
        $t1.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $(if ($p.status -eq 'closed') { 'MutedBrush' } else { 'TextBrush' }))
        $t2 = [System.Windows.Controls.TextBlock]::new()
        $t2.FontSize = 11
        $t2.Text = ('未完成 {0} / 共 {1} 件' -f $open, $all) + $(if ($last.Count -gt 0) { '　最近进展 ' + (Parse-Day $last[0].date).ToString('M月d日') } else { '' })
        $t2.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'SubBrush')
        [void]$sp.Children.Add($t1); [void]$sp.Children.Add($t2)
        $b.Child = $sp
        $b.ToolTip = if ($p.kind -eq 'project') { '双击修改项目名称' } else { '固定类别' }
        $b.add_MouseLeftButtonUp({
            param($s, $e)
            if ($script:ProjSel -eq [string]$s.Tag) { return }
            $script:ProjSel = [string]$s.Tag; $script:CalEdit = $null
            foreach ($c in $s.Parent.Children) { $c.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, $(if ([string]$c.Tag -eq $script:ProjSel) { 'HoverBrush' } else { 'CardBrush' })) }
            Render-ProjView
        })
        $b.add_MouseLeftButtonDown({ param($s, $e) if ($e.ClickCount -eq 2) { $e.Handled = $true; $script:ProjSel = [string]$s.Tag; Rename-ProjectById ([string]$s.Tag) } })
        [void]$lb.Children.Add($b)
    }
    Render-ProjView
}

function Render-ProjView {
    $p = if ($script:ProjSel) { Find-Project $script:ProjSel } else { $null }
    $script:ProjRename.IsEnabled = [bool]($p -and $p.kind -eq 'project')
    $script:ProjClose.IsEnabled = [bool]($p -and $p.kind -eq 'project')
    $script:ProjClose.Content = if ($p -and $p.status -eq 'closed') { '恢复为进行中' } else { '结束项目' }
    if ($p) {
        $all0 = @($script:Items | Where-Object { $_.project -eq $p.id }) + @($script:Work | Where-Object { $_.proj -eq $p.id })
        $dates = @($all0 | ForEach-Object { if ($_.date) { Parse-Day $_.date } else { (Parse-Dt $_.created) } } | Where-Object { $_ } | Sort-Object)
        $f = if ($dates.Count -gt 0) { $dates[0].Date } else { (Get-Date).Date }
        $t = (Get-Date).Date
        foreach ($x in $dates) { if ($x.Date -gt $t) { $t = $x.Date } }
        $script:ProjLines = Build-ReportLines $f $t ($p.text + '　全部记录（' + $f.ToString('yyyy年M月d日') + '—' + $t.ToString('M月d日') + '）') $true $p.id
        Show-ReportLines $script:ProjLines $script:ProjView
    } else {
        $script:ProjLines = $null
        $script:ProjView.Children.Clear()
        $ph = [System.Windows.Controls.TextBlock]::new()
        $ph.Text = '在左边点一个项目，这里显示它从开始到现在的全部待办和每日进展。'
        $ph.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, 'MutedBrush')
        [void]$script:ProjView.Children.Add($ph)
    }
    Render-SideEditor '点中间任意一条待办或进展，在这里修改内容、日期和时间，改完各处一起更新。'
}

# ---------- 总的刷新 ----------
function Render-Cal {
    if (-not $script:CalWin -or -not $script:CalWin.IsVisible) { return }
    # 右侧编辑框里还没保存的文字，先让它失去焦点触发保存，再重画
    $fe0 = [System.Windows.Input.Keyboard]::FocusedElement
    if (($fe0 -is [System.Windows.Controls.TextBox]) -and -not $fe0.IsReadOnly -and $script:CalWin.IsKeyboardFocusWithin -and -not $script:CalRendering) {
        $script:CalRendering = $true
        try { [System.Windows.Input.Keyboard]::ClearFocus() } finally { $script:CalRendering = $false }
    }
    $m = $script:CalMode
    foreach ($b in $script:CalModeBtns) {
        $on = [string]$b.Tag -eq $m
        $b.FontWeight = if ($on) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
        $b.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $(if ($on) { 'AccentBrush' } else { 'SubBrush' }))
    }
    if ($script:ChkTime) { $script:ChkTime.IsChecked = Test-ShowTime }
    $isCal = $m -eq 'week' -or $m -eq 'month'
    $script:CalBar.Visibility = if ($isCal) { 'Visible' } else { 'Collapsed' }
    $script:CalPanel.Visibility = if ($isCal) { 'Visible' } else { 'Collapsed' }
    $script:RepPanel.Visibility = if ($m -eq 'report') { 'Visible' } else { 'Collapsed' }
    $script:ProjPanel.Visibility = if ($m -eq 'project') { 'Visible' } else { 'Collapsed' }
    $script:ReimbPanel.Visibility = if ($m -eq 'reimb') { 'Visible' } else { 'Collapsed' }
    if ($isCal) { $script:CalDetail = $script:CalDetailCal; Render-CalGrid }
    elseif ($m -eq 'report') { $script:CalDetail = $script:RepDetail; Render-Report }
    elseif ($m -eq 'reimb') { Render-Reimb }
    else { $script:CalDetail = $script:ProjDetail; Render-ProjectsTab }
}

function Show-Calendar([string]$mode) {
    if ($mode) { $script:CalMode = $mode }
    if ($script:CalWin -and $script:CalWin.IsLoaded) {
        $script:CalWin.Show(); [void]$script:CalWin.Activate()
        if ($script:CalWin.WindowState -eq 'Minimized') { $script:CalWin.WindowState = 'Normal' }
        Render-Cal
        return
    }
    [xml]$cx = @'
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="桌面待办 · 工作台" Width="1240" Height="800" MinWidth="820" MinHeight="520"
        WindowStartupLocation="CenterScreen" FontFamily="Microsoft YaHei UI" FontSize="13"
        Background="{DynamicResource BgBrush}" Foreground="{DynamicResource TextBrush}">
  <Grid Margin="14,10,14,12">
    <Grid.RowDefinitions><RowDefinition Height="Auto"/><RowDefinition Height="Auto"/><RowDefinition Height="*"/></Grid.RowDefinitions>
    <StackPanel Grid.Row="0" Orientation="Horizontal" Margin="0,0,0,10">
      <TextBlock x:Name="ModeWeek" Tag="week" Text="周视图" FontSize="15" Cursor="Hand" Margin="0,0,22,0"/>
      <TextBlock x:Name="ModeMonth" Tag="month" Text="月视图" FontSize="15" Cursor="Hand" Margin="0,0,22,0"/>
      <TextBlock x:Name="ModeReport" Tag="report" Text="报告" FontSize="15" Cursor="Hand" Margin="0,0,22,0"/>
      <TextBlock x:Name="ModeProject" Tag="project" Text="项目" FontSize="15" Cursor="Hand" Margin="0,0,22,0"/>
      <TextBlock x:Name="ModeReimb" Tag="reimb" Text="报销" FontSize="15" Cursor="Hand"/>
    </StackPanel>
    <DockPanel x:Name="CalBar" Grid.Row="1" Margin="0,0,0,10">
      <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
        <CheckBox x:Name="ChkWork" Margin="0,0,10,0" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center" ToolTip="每天做了什么（绿条）：自己记的工作、待办的进展、勾掉具体事项自动记的">
          <StackPanel Orientation="Horizontal"><Border Width="3" Height="13" Background="#2E9E5B" Margin="0,0,4,0"/><TextBlock Text="工作记录"/></StackPanel></CheckBox>
        <CheckBox x:Name="ChkDone" Margin="0,0,10,0" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center" ToolTip="这天打勾做完的待办（灰条，前面有 ✓）">
          <StackPanel Orientation="Horizontal"><Border Width="3" Height="13" Background="#9AA0A6" Margin="0,0,4,0"/><TextBlock Text="完成"/></StackPanel></CheckBox>
        <CheckBox x:Name="ChkTodo" Margin="0,0,10,0" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center" ToolTip="这天有提醒或截止的待办（蓝条）">
          <StackPanel Orientation="Horizontal"><Border Width="3" Height="13" Background="#2F6FED" Margin="0,0,4,0"/><TextBlock Text="提醒/截止"/></StackPanel></CheckBox>
        <CheckBox x:Name="ChkOngoing" Margin="0,0,16,0" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center" ToolTip="开始了还没做完的待办（黄条），在它进行的每一天都显示">
          <StackPanel Orientation="Horizontal"><Border Width="3" Height="13" Background="#C9A227" Margin="0,0,4,0"/><TextBlock Text="进行中"/></StackPanel></CheckBox>
        <CheckBox x:Name="ChkTime" Content="几点几分" Margin="0,0,12,0" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center" ToolTip="显示几点几分"/>
        <Button x:Name="CalExpand" Content="全部展开" Padding="10,3" Margin="0,0,8,0" Visibility="Collapsed"/>
        <Button x:Name="CalCopy" Content="复制清单" Padding="10,3" Visibility="Collapsed"/>
        <Button x:Name="CalMore" Content="⋯" Width="34" Padding="0,1" ToolTip="全部展开 / 收起、复制清单"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
        <Button x:Name="CalPrev" Content="‹" Width="30" Padding="0,1"/>
        <Button x:Name="CalToday" Content="今天" Margin="6,0" Padding="10,1"/>
        <Button x:Name="CalNext" Content="›" Width="30" Padding="0,1"/>
        <TextBlock x:Name="CalTitle" FontSize="17" FontWeight="SemiBold" Margin="16,0,0,0" VerticalAlignment="Center"/>
        <TextBlock x:Name="CalCount" Margin="16,0,0,0" FontSize="12" Foreground="{DynamicResource SubBrush}" VerticalAlignment="Center"/>
      </StackPanel>
    </DockPanel>
    <Grid x:Name="CalPanel" Grid.Row="2">
      <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="360"/></Grid.ColumnDefinitions>
      <Border Grid.Column="0" BorderThickness="1,1,0,0" BorderBrush="{DynamicResource LineBrush}">
        <ScrollViewer x:Name="CalScroll" VerticalScrollBarVisibility="Auto" HorizontalScrollBarVisibility="Disabled"><Grid x:Name="CalGrid"/></ScrollViewer>
      </Border>
      <Border x:Name="CalSide" Grid.Column="1" Margin="12,0,0,0" Padding="14,10,14,12" CornerRadius="6" Background="{DynamicResource CardBrush}" BorderThickness="1" BorderBrush="{DynamicResource LineBrush}">
        <DockPanel>
          <TextBlock x:Name="CalSideClose" DockPanel.Dock="Top" HorizontalAlignment="Right" Text="✕" FontSize="13" Cursor="Hand" Margin="0,0,0,2" ToolTip="关掉右边这一栏（点日历里的任意一天或一条，会再打开）" Foreground="{DynamicResource SubBrush}"/>
          <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="CalDetail"/></ScrollViewer>
        </DockPanel>
      </Border>
    </Grid>
    <DockPanel x:Name="RepPanel" Grid.Row="1" Grid.RowSpan="2" Visibility="Collapsed">
      <DockPanel DockPanel.Dock="Top" Margin="0,0,0,10">
        <StackPanel DockPanel.Dock="Right" Orientation="Horizontal" VerticalAlignment="Center">
          <CheckBox x:Name="RepSubs" Content="待办没记进展时，列出它的具体事项" IsChecked="True" Margin="0,0,14,0" Foreground="{DynamicResource TextBrush}" VerticalAlignment="Center"/>
          <Button x:Name="RepCopy" Content="复制文字" Padding="12,3" Margin="0,0,8,0"/>
          <Button x:Name="RepWord" Content="导出 Word" Padding="12,3"/>
        </StackPanel>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
          <TextBlock x:Name="RkDay" Tag="day" Text="日报" Cursor="Hand" Margin="0,0,14,0" FontSize="14"/>
          <TextBlock x:Name="RkWeek" Tag="week" Text="周报" Cursor="Hand" Margin="0,0,14,0" FontSize="14"/>
          <TextBlock x:Name="RkMonth" Tag="month" Text="月报" Cursor="Hand" Margin="0,0,14,0" FontSize="14"/>
          <TextBlock x:Name="RkQuarter" Tag="quarter" Text="季度" Cursor="Hand" Margin="0,0,14,0" FontSize="14"/>
          <TextBlock x:Name="RkYear" Tag="year" Text="年度" Cursor="Hand" Margin="0,0,14,0" FontSize="14"/>
          <TextBlock x:Name="RkCustom" Tag="custom" Text="自选日期" Cursor="Hand" Margin="0,0,20,0" FontSize="14"/>
          <StackPanel x:Name="RepNav" Orientation="Horizontal">
            <Button x:Name="RepPrev" Content="‹" Width="30"/>
            <Button x:Name="RepNow" Content="当前" Margin="6,0" Padding="10,1"/>
            <Button x:Name="RepNext" Content="›" Width="30"/>
          </StackPanel>
          <StackPanel x:Name="RepCustom" Orientation="Horizontal" Visibility="Collapsed">
            <DatePicker x:Name="RepFrom" Width="130"/>
            <TextBlock Text="至" Margin="8,0" VerticalAlignment="Center"/>
            <DatePicker x:Name="RepTo" Width="130"/>
          </StackPanel>
          <TextBlock x:Name="RepTitle" FontSize="16" FontWeight="SemiBold" Margin="18,0,0,0" VerticalAlignment="Center"/>
        </StackPanel>
      </DockPanel>
      <Grid>
        <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="360"/></Grid.ColumnDefinitions>
        <Border Grid.Column="0" Padding="18,14" CornerRadius="6" Background="{DynamicResource CardBrush}" BorderThickness="1" BorderBrush="{DynamicResource LineBrush}">
          <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="RepView"/></ScrollViewer>
        </Border>
        <Border Grid.Column="1" Margin="12,0,0,0" Padding="14,12" CornerRadius="6" Background="{DynamicResource CardBrush}" BorderThickness="1" BorderBrush="{DynamicResource LineBrush}">
          <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="RepDetail"/></ScrollViewer>
        </Border>
      </Grid>
    </DockPanel>
    <Grid x:Name="ProjPanel" Grid.Row="1" Grid.RowSpan="2" Visibility="Collapsed">
      <Grid.ColumnDefinitions><ColumnDefinition Width="340"/><ColumnDefinition Width="*"/></Grid.ColumnDefinitions>
      <DockPanel Grid.Column="0">
        <WrapPanel DockPanel.Dock="Top" Margin="0,0,0,10">
          <Button x:Name="ProjNew" Content="新建项目" Padding="10,3" Margin="0,0,8,0"/>
          <Button x:Name="ProjRename" Content="重命名" Padding="10,3" Margin="0,0,8,0"/>
          <Button x:Name="ProjClose" Content="结束项目" Padding="10,3"/>
        </WrapPanel>
        <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="ProjList"/></ScrollViewer>
      </DockPanel>
      <DockPanel Grid.Column="1" Margin="12,0,0,0">
        <StackPanel DockPanel.Dock="Top" Orientation="Horizontal" Margin="0,0,0,10">
          <Button x:Name="ProjCopy" Content="复制文字" Padding="12,3" Margin="0,0,8,0"/>
          <Button x:Name="ProjWord" Content="导出 Word" Padding="12,3"/>
        </StackPanel>
        <Grid>
          <Grid.ColumnDefinitions><ColumnDefinition Width="*"/><ColumnDefinition Width="340"/></Grid.ColumnDefinitions>
          <Border Grid.Column="0" Padding="18,14" CornerRadius="6" Background="{DynamicResource CardBrush}" BorderThickness="1" BorderBrush="{DynamicResource LineBrush}">
            <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="ProjView"/></ScrollViewer>
          </Border>
          <Border Grid.Column="1" Margin="12,0,0,0" Padding="14,12" CornerRadius="6" Background="{DynamicResource CardBrush}" BorderThickness="1" BorderBrush="{DynamicResource LineBrush}">
            <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="ProjDetail"/></ScrollViewer>
          </Border>
        </Grid>
      </DockPanel>
    </Grid>
    <Grid x:Name="ReimbPanel" Grid.Row="1" Grid.RowSpan="2" Visibility="Collapsed" AllowDrop="True" Background="Transparent">
      <Grid.ColumnDefinitions><ColumnDefinition Width="270"/><ColumnDefinition Width="*"/><ColumnDefinition Width="330"/></Grid.ColumnDefinitions>
      <DockPanel Grid.Column="0">
        <WrapPanel DockPanel.Dock="Top" Margin="0,0,0,8">
          <Button x:Name="RbNew" Content="新建报销单" Padding="10,3" Margin="0,0,8,4"/>
          <Button x:Name="RbLedger" Content="导出台账" Padding="10,3" Margin="0,0,8,4" ToolTip="所有报销单、全部明细、按项目汇总，导成一个 Excel"/>
          <Button x:Name="RbSet" Content="设置" Padding="10,3" Margin="0,0,0,4" ToolTip="姓名、默认成本中心、其他事项说明里固定的那句话"/>
        </WrapPanel>
        <WrapPanel x:Name="RbFilter" DockPanel.Dock="Top" Margin="2,0,0,4"/>
        <TextBlock x:Name="RbSum" DockPanel.Dock="Top" FontSize="11" TextWrapping="Wrap" Margin="2,0,0,8" Foreground="{DynamicResource SubBrush}"/>
        <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="RbList"/></ScrollViewer>
      </DockPanel>
      <DockPanel Grid.Column="1" Margin="12,0,0,0">
        <WrapPanel DockPanel.Dock="Top" Margin="0,0,0,6">
          <Button x:Name="RbImport" Content="导入发票…" Padding="10,3" Margin="0,0,8,4" ToolTip="选下载好的发票（PDF / OFD / 图片），也可以直接拖进这个窗口"/>
          <Button x:Name="RbTabList" Content="费用明细" Padding="10,3" Margin="0,0,4,4"/>
          <Button x:Name="RbTabOA" Content="OA 填报" Padding="10,3" Margin="0,0,8,4" ToolTip="按 OA 的格子排好，点一下复制"/>
          <Button x:Name="RbExcel" Content="导出 Excel" Padding="10,3" Margin="0,0,8,4"/>
          <Button x:Name="RbPrint" Content="打印 ▾" Padding="10,3" Margin="0,0,8,4" ToolTip="粘贴单 + 每张发票一页 A4，拼成一个 PDF"/>
          <Button x:Name="RbFolder" Content="发票文件夹" Padding="10,3" Margin="0,0,0,4"/>
        </WrapPanel>
        <StackPanel x:Name="RbQuickHost" DockPanel.Dock="Top" Margin="0,0,0,8"/>
        <Border Padding="16,12" CornerRadius="6" Background="{DynamicResource CardBrush}" BorderThickness="1" BorderBrush="{DynamicResource LineBrush}">
          <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="RbViewPanel"/></ScrollViewer>
        </Border>
      </DockPanel>
      <Border Grid.Column="2" Margin="12,0,0,0" Padding="14,12" CornerRadius="6" Background="{DynamicResource CardBrush}" BorderThickness="1" BorderBrush="{DynamicResource LineBrush}">
        <ScrollViewer VerticalScrollBarVisibility="Auto"><StackPanel x:Name="RbSide"/></ScrollViewer>
      </Border>
    </Grid>
  </Grid>
</Window>
'@
    $win = [Windows.Markup.XamlReader]::Load((New-Object System.Xml.XmlNodeReader $cx))
    Set-WindowIcon $win
    $script:CalWin = $win
    foreach ($n in 'CalSide', 'CalSideClose', 'CalPanel', 'CalScroll', 'CalGrid', 'CalDetail', 'CalTitle', 'CalCount', 'CalBar', 'CalPanel', 'RepPanel', 'ProjPanel', 'RepView', 'RepDetail', 'ProjView', 'ProjDetail', 'RepTitle', 'RepSubs', 'RepNav', 'RepCustom', 'ProjList', 'ProjRename', 'ProjClose') {
        Set-Variable -Name $n -Value $win.FindName($n) -Scope Script
    }
    $script:CalExpandBtn = $win.FindName('CalExpand')
    Init-ReimbPanel $win
    $script:CalSideOpen = $false
    $script:CalSideClose.add_MouseLeftButtonUp({ Flush-CalFocus; $script:CalEdit = $null; Set-CalSide $false })
    $win.FindName('CalMore').add_Click({
        param($s, $e)
        $m = [System.Windows.Controls.ContextMenu]::new()
        [void]$m.Items.Add((New-MenuItem ([string]$script:CalExpandBtn.Content + '具体事项') '' { $script:CalExpandBtn.RaiseEvent([System.Windows.RoutedEventArgs]::new([System.Windows.Controls.Primitives.ButtonBase]::ClickEvent)) }))
        [void]$m.Items.Add((New-MenuItem '复制清单（按天列好）' '' { try { [System.Windows.Clipboard]::SetText((Get-CalText)); Notify '已复制清单' '按天整理好的清单已经在剪贴板里。' } catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。' } }))
        $m.PlacementTarget = $s; $m.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
        $m.IsOpen = $true
    })
    $script:CalScroll.add_SizeChanged({
        param($s, $e)
        if (-not $e.HeightChanged) { return }
        if ($script:CalMode -eq 'week') { $script:CalGrid.MinHeight = [math]::Max(0, $s.ActualHeight - 2); return }
        if ($script:CalMode -ne 'month' -or $script:CalInline) { return }
        $fe = [System.Windows.Input.Keyboard]::FocusedElement
        if (($fe -is [System.Windows.Controls.TextBox]) -and -not $fe.IsReadOnly) { return }
        if ([math]::Abs($s.ActualHeight - [double]$script:CalLastH) -gt 30) { Render-CalGrid }
    })
    $script:CalDetailCal = $script:CalDetail
    $script:CalModeBtns = @('ModeWeek', 'ModeMonth', 'ModeReport', 'ModeProject', 'ModeReimb' | ForEach-Object { $win.FindName($_) })
    foreach ($b in $script:CalModeBtns) { $b.add_MouseLeftButtonUp({ param($s, $e) $script:CalMode = [string]$s.Tag; $script:CalEdit = $null; if ($script:CalMode -eq 'week' -or $script:CalMode -eq 'month') { $script:CalAnchor = $script:CalSel }; Render-Cal }) }
    $script:RepKindBtns = @('RkDay', 'RkWeek', 'RkMonth', 'RkQuarter', 'RkYear', 'RkCustom' | ForEach-Object { $win.FindName($_) })
    foreach ($b in $script:RepKindBtns) { $b.add_MouseLeftButtonUp({ param($s, $e) $script:RepKind = [string]$s.Tag; Render-Report }) }

    $move = {
        param([int]$dir)
        if ($script:CalMode -eq 'week') { $script:CalAnchor = $script:CalAnchor.AddDays(7 * $dir); $script:CalSel = Get-WeekStart $script:CalAnchor }
        else { $script:CalAnchor = $script:CalAnchor.AddMonths($dir); $script:CalSel = [datetime]::new($script:CalAnchor.Year, $script:CalAnchor.Month, 1) }
        $script:CalEdit = $null
        Render-Cal
    }
    $script:CalMove = $move
    $win.FindName('CalPrev').add_Click({ & $script:CalMove -1 })
    $win.FindName('CalNext').add_Click({ & $script:CalMove 1 })
    $win.FindName('CalToday').add_Click({ $script:CalAnchor = (Get-Date).Date; $script:CalSel = (Get-Date).Date; $script:CalEdit = $null; Render-Cal })
    foreach ($pair in @(@('ChkTodo', 'todo'), @('ChkDone', 'done'), @('ChkWork', 'work'), @('ChkOngoing', 'ongoing'))) {
        $cb = $win.FindName($pair[0]); $cb.Tag = $pair[1]
        $cb.IsChecked = [bool]$script:CalShow[$pair[1]]
        $cb.add_Click({ param($s, $e) $script:CalShow[[string]$s.Tag] = [bool]$s.IsChecked; Render-Cal })
    }
    $script:ChkTime = $win.FindName('ChkTime')
    $script:ChkTime.IsChecked = Test-ShowTime
    $script:ChkTime.add_Click({ param($s, $e) Set-ShowTime ([bool]$s.IsChecked) })
    $script:CalExpandBtn.add_Click({
        $closed = @($script:CalVisibleTodos | Where-Object { -not $script:CalOpen.ContainsKey($_) })
        if ($closed.Count -gt 0) { foreach ($id in $script:CalVisibleTodos) { $script:CalOpen[$id] = $true } } else { $script:CalOpen = @{} }
        Render-Cal
    })
    $win.FindName('CalCopy').add_Click({ try { [System.Windows.Clipboard]::SetText((Get-CalText)); Notify '已复制清单' '按天整理好的清单已经在剪贴板里。' } catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。' } })

    $win.FindName('RepPrev').add_Click({ Move-Report -1 })
    $win.FindName('RepNext').add_Click({ Move-Report 1 })
    $win.FindName('RepNow').add_Click({ $script:RepAnchor = (Get-Date).Date; Render-Report })
    $script:RepSubs.add_Click({ Render-Report })
    $win.FindName('RepCopy').add_Click({ try { [System.Windows.Clipboard]::SetText((Get-ReportText $script:RepLines)); Notify '已复制报告' '可以直接粘到微信或文档里。' } catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。' } })
    $win.FindName('RepWord').add_Click({ Save-ReportDocx })
    $rf = $win.FindName('RepFrom'); $rt = $win.FindName('RepTo')
    $f0 = Parse-Day ([string]$script:S.repFrom); $t0 = Parse-Day ([string]$script:S.repTo)
    $rf.SelectedDate = if ($f0) { $f0 } else { (Get-Date).Date.AddDays(-6) }
    $rt.SelectedDate = if ($t0) { $t0 } else { (Get-Date).Date }
    $rf.add_SelectedDateChanged({ param($s, $e) if ($s.SelectedDate) { $script:S.repFrom = ([datetime]$s.SelectedDate).ToString('yyyy-MM-dd'); Save-Settings; Render-Report } })
    $rt.add_SelectedDateChanged({ param($s, $e) if ($s.SelectedDate) { $script:S.repTo = ([datetime]$s.SelectedDate).ToString('yyyy-MM-dd'); Save-Settings; Render-Report } })
    if (-not $script:S.repFrom) { $script:S.repFrom = ([datetime]$rf.SelectedDate).ToString('yyyy-MM-dd'); $script:S.repTo = ([datetime]$rt.SelectedDate).ToString('yyyy-MM-dd') }

    $win.FindName('ProjNew').add_Click({ $p = New-Project-Dialog; if ($p) { $script:ProjSel = $p.id; Render } })
    $script:ProjRename.add_Click({ if ($script:ProjSel) { Rename-ProjectById $script:ProjSel } })
    $script:ProjClose.add_Click({
        $p = Find-Project $script:ProjSel; if (-not $p -or $p.kind -ne 'project') { return }
        $p.status = if ($p.status -eq 'closed') { 'active' } else { 'closed' }
        Save-Projects; Render
    })
    $win.FindName('ProjCopy').add_Click({ if ($script:ProjLines) { try { [System.Windows.Clipboard]::SetText((Get-ReportText $script:ProjLines)); Notify '已复制' '项目的全部记录已经在剪贴板里。' } catch {} } })
    $win.FindName('ProjWord').add_Click({
        if (-not $script:ProjLines) { return }
        $p = Find-Project $script:ProjSel
        $dlg = New-Object Microsoft.Win32.SaveFileDialog
        $dlg.Filter = 'Word 文档|*.docx'; $dlg.FileName = ($p.text -replace '[\\/:*?"<>|]', '') + '_全部记录.docx'
        if ($dlg.ShowDialog($script:CalWin)) { try { Export-ReportDocx $script:ProjLines $dlg.FileName; Start-Process $dlg.FileName } catch { [System.Windows.MessageBox]::Show($script:CalWin, '导出失败：' + $_.Exception.Message, '桌面待办') | Out-Null } }
    })

    $win.add_Closed({ $script:CalWin = $null })
    $win.Show()
    Render-Cal
}

# 报销单据粘贴单（单位给的模板，打印时当底图）
$PasteSheetB64 = 'JVBERi0xLjcKJb/3ov4KMSAwIG9iago8PCAvTGFuZyAoemgpIC9NYXJrSW5mbyA8PCAvTWFya2VkIHRydWUgPj4gL01ldGFkYXRhIDMgMCBSIC9PdXRwdXRJbnRlbnRzIFsgPDwgL0Rlc3RPdXRwdXRQcm9maWxlIDQgMCBSIC9JbmZvIChDcmVhdG9yOiBIUCAgICAgTWFudWZhY3R1cmVyOklFQyAgICBNb2RlbDpzUkdCKSAvT3V0cHV0Q29uZGl0aW9uSWRlbnRpZmllciAoc1JHQikgL1JlZ2lzdHJ5TmFtZSAoaHR0cDovL3d3dy5jb2xvci5vcmcpIC9TIC9HVFNfUERGQTEgL1R5cGUgL091dHB1dEludGVudCA+PiBdIC9QYWdlcyA1IDAgUiAvU3RydWN0VHJlZVJvb3QgNiAwIFIgL1R5cGUgL0NhdGFsb2cgL1ZpZXdlclByZWZlcmVuY2VzIDcgMCBSID4+CmVuZG9iagoyIDAgb2JqCjw8ID4+CmVuZG9iagozIDAgb2JqCjw8IC9TdWJ0eXBlIC9YTUwgL1R5cGUgL01ldGFkYXRhIC9MZW5ndGggMjY5NCA+PgpzdHJlYW0KPD94cGFja2V0IGJlZ2luPSLvu78iIGlkPSJXNU0wTXBDZWhpSHpyZVN6TlRjemtjOWQiPz4KPHg6eG1wbWV0YSB4bWxuczp4PSJhZG9iZTpuczptZXRhLyIgeDp4bXB0az0iMy4xLTcwMSI+CjxyZGY6UkRGIHhtbG5zOnJkZj0iaHR0cDovL3d3dy53My5vcmcvMTk5OS8wMi8yMi1yZGYtc3ludGF4LW5zIyI+CjxyZGY6RGVzY3JpcHRpb24geG1sbnM6cGRmPSJodHRwOi8vbnMuYWRvYmUuY29tL3BkZi8xLjMvIiByZGY6YWJvdXQ9IiI+CjxwZGY6UHJvZHVjZXI+TWljcm9zb2Z0wq4gV29yZCAyMDIxPC9wZGY6UHJvZHVjZXI+PC9yZGY6RGVzY3JpcHRpb24+CjxyZGY6RGVzY3JpcHRpb24geG1sbnM6eG1wPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvIiByZGY6YWJvdXQ9IiI+Cjx4bXA6Q3JlYXRvclRvb2w+TWljcm9zb2Z0wq4gV29yZCAyMDIxPC94bXA6Q3JlYXRvclRvb2w+PC9yZGY6RGVzY3JpcHRpb24+CjxyZGY6RGVzY3JpcHRpb24geG1sbnM6cGRmYWlkPSJodHRwOi8vd3d3LmFpaW0ub3JnL3BkZmEvbnMvaWQvIiByZGY6YWJvdXQ9IiI+CjxwZGZhaWQ6cGFydD4zPC9wZGZhaWQ6cGFydD48cGRmYWlkOmNvbmZvcm1hbmNlPkE8L3BkZmFpZDpjb25mb3JtYW5jZT48L3JkZjpEZXNjcmlwdGlvbj4KICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCiAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAKICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgIAogICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgCjwvcmRmOlJERj48L3g6eG1wbWV0YT4KCjw/eHBhY2tldCBlbmQ9InciPz4KCmVuZHN0cmVhbQplbmRvYmoKNCAwIG9iago8PCAvRmlsdGVyIC9GbGF0ZURlY29kZSAvTiAzIC9MZW5ndGggMjU5MyA+PgpzdHJlYW0KeJydlndUVNcWh8+9d3qhzTB0GHqvUgYQ6R2kV1EYZgYYygDDDNgLIioQUUSkKYIEBQwYDUViRRQLAVEBe0CCgBKDUWyoZEbWSnx5ee/l5ffHPd/aZ+9z99l737UuACQvPy4vHZYCII0n4Ad7utAjo6Lp2H4AAzzAAHMAmKysDP8Qj1Agkre7Kz1L5AT+Ra+HASRebxl7BdLp4P+TNCuDLwAAChTxEjYniyXiPBGn5ggyxPZZEVPjU8QMo8TMFyUoYnkxJy6y0WefRXYSMzuNxxaxOOcMdhpbzD0i3pEt5IgY8RNxfjaXkyPi2yLWShWmcUX8VhybxmFmAYAiie0CDitJxGYiJvFDg11FvBQAHCnxC47/ggWc1QLxpVzTM9bwuYlJAroeS59ubmvLoHtxclI5AoFxIJOVwuSz6a7paRlM3hoAFu/8WTLi2tJFRbY2t7W2NrYwMf+iUP91829K3NtFehn0uWcQre8P21/5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ehGjpfJq72DX4pJzvHBI2PGhYgJqdEKZLai54zTtd9+mroNCBkKP7Tv0PXM2A/f2jYeekNvHWdhHwZc23tAuBudZWeNhK3Rj1nhGzZFhvCFlw4c6Yrwe1EONFzfGU6S4h4yXurdvPKJHSOQZzw297eA8K2s85WmFNfnIM4vLvPHc0OLOhzpivB7UQ8/1DhXwfFyplPNIoY3tWxSmU/0HZEGnmwplbmRzdHJlYW0KZW5kb2JqCjEzIDAgb2JqCjw8IC9CTSAvTm9ybWFsIC9UeXBlIC9FeHRHU3RhdGUgL2NhIDEgPj4KZW5kb2JqCjE0IDAgb2JqCjw8IC9CTSAvTm9ybWFsIC9DQSAxIC9UeXBlIC9FeHRHU3RhdGUgPj4KZW5kb2JqCjE1IDAgb2JqCjw8IC9CYXNlRm9udCAvQkNERUVFK1NpbVN1biAvRGVzY2VuZGFudEZvbnRzIDI3IDAgUiAvRW5jb2RpbmcgL0lkZW50aXR5LUggL1N1YnR5cGUgL1R5cGUwIC9Ub1VuaWNvZGUgMjggMCBSIC9UeXBlIC9Gb250ID4+CmVuZG9iagoxNiAwIG9iago8PCAvQmFzZUZvbnQgL0JDREZFRStDYWxpYnJpLUJvbGQgL0VuY29kaW5nIC9XaW5BbnNpRW5jb2RpbmcgL0ZpcnN0Q2hhciAzMiAvRm9udERlc2NyaXB0b3IgMjkgMCBSIC9MYXN0Q2hhciAzMiAvTmFtZSAvRjIgL1N1YnR5cGUgL1RydWVUeXBlIC9UeXBlIC9Gb250IC9XaWR0aHMgMzAgMCBSID4+CmVuZG9iagoxNyAwIG9iago8PCAvQmFzZUZvbnQgL0JDREdFRStDYWxpYnJpIC9FbmNvZGluZyAvV2luQW5zaUVuY29kaW5nIC9GaXJzdENoYXIgMzIgL0ZvbnREZXNjcmlwdG9yIDMxIDAgUiAvTGFzdENoYXIgMzIgL05hbWUgL0YzIC9TdWJ0eXBlIC9UcnVlVHlwZSAvVHlwZSAvRm9udCAvV2lkdGhzIDMyIDAgUiA+PgplbmRvYmoKMTggMCBvYmoKPDwgL0sgWyAwIDMzIDAgUiAyIDM0IDAgUiA0IDM1IDAgUiA2IDM2IDAgUiA4IDM3IDAgUiAxMCAzOCAwIFIgMTIgMzkgMCBSIF0gL1AgOSAwIFIgL1BnIDggMCBSIC9TIC9QIC9UeXBlIC9TdHJ1Y3RFbGVtID4+CmVuZG9iagoxOSAwIG9iago8PCAvSyBbIDQwIDAgUiA0MSAwIFIgNDIgMCBSIDQzIDAgUiAxOCA0NCAwIFIgMjAgNDUgMCBSIDQ2IDAgUiA0NyAwIFIgNDggMCBSIDQ5IDAgUiAyNiA1MCAwIFIgNTEgMCBSIDUyIDAgUiAzMCA1MyAwIFIgXSAvUCA5IDAgUiAvUGcgOCAwIFIgL1MgL1AgL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjIwIDAgb2JqCjw8IC9LIFsgNTQgMCBSIDU1IDAgUiA1NiAwIFIgXSAvUCA5IDAgUiAvUGcgOCAwIFIgL1MgL1RleHRib3ggL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjIxIDAgb2JqCjw8IC9LIFsgNTcgMCBSIDU4IDAgUiBdIC9QIDkgMCBSIC9QZyA4IDAgUiAvUyAvVGV4dGJveCAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKMjIgMCBvYmoKPDwgL0sgWyA1OSAwIFIgNjAgMCBSIF0gL1AgOSAwIFIgL1BnIDggMCBSIC9TIC9UZXh0Ym94IC9UeXBlIC9TdHJ1Y3RFbGVtID4+CmVuZG9iagoyMyAwIG9iago8PCAvQWx0ICgpIC9LIFsgNTIgXSAvUCA5IDAgUiAvUGcgOCAwIFIgL1MgL0ZpZ3VyZSAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKMjQgMCBvYmoKPDwgL0sgWyA2MSAwIFIgNjIgMCBSIDYzIDAgUiAzNSA2NCAwIFIgMzcgNjUgMCBSIDY2IDAgUiA2NyAwIFIgNDEgNjggMCBSIDY5IDAgUiA3MCAwIFIgNzEgMCBSIF0gL1AgOSAwIFIgL1BnIDggMCBSIC9TIC9QIC9UeXBlIC9TdHJ1Y3RFbGVtID4+CmVuZG9iagoyNSAwIG9iago8PCAvSyBbIDcyIDAgUiA3MyAwIFIgNzQgMCBSIDQ5IDc1IDAgUiA3NiAwIFIgXSAvUCA5IDAgUiAvUGcgOCAwIFIgL1MgL1AgL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjI2IDAgb2JqClsgMTggMCBSIDMzIDAgUiAxOCAwIFIgMzQgMCBSIDE4IDAgUiAzNSAwIFIgMTggMCBSIDM2IDAgUiAxOCAwIFIgMzcgMCBSIDE4IDAgUiAzOCAwIFIgMTggMCBSIDM5IDAgUiA0MCAwIFIgNDEgMCBSIDQyIDAgUiA0MyAwIFIgMTkgMCBSIDQ0IDAgUiAxOSAwIFIgNDUgMCBSIDQ2IDAgUiA0NyAwIFIgNDggMCBSIDQ5IDAgUiAxOSAwIFIgNTAgMCBSIDUxIDAgUiA1MiAwIFIgMTkgMCBSIDUzIDAgUiA2MSAwIFIgNjIgMCBSIDYzIDAgUiAyNCAwIFIgNjQgMCBSIDI0IDAgUiA2NSAwIFIgNjYgMCBSIDY3IDAgUiAyNCAwIFIgNjggMCBSIDY5IDAgUiA3MCAwIFIgNzEgMCBSIDcyIDAgUiA3MyAwIFIgNzQgMCBSIDI1IDAgUiA3NSAwIFIgNzYgMCBSIDIzIDAgUiA1OSAwIFIgNjAgMCBSIDc3IDAgUiA1NyAwIFIgNTggMCBSIDc4IDAgUiA1NCAwIFIgNTUgMCBSIDc5IDAgUiA4MCAwIFIgXQplbmRvYmoKMjcgMCBvYmoKWyA4MSAwIFIgXQplbmRvYmoKMjggMCBvYmoKPDwgL0ZpbHRlciAvRmxhdGVEZWNvZGUgL0xlbmd0aCAzNzIgPj4Kc3RyZWFtCnicXdJfa4MwEADwdz9FHreHYhITjVAETSL0YX9Ytw9g9doJM0pqH/rtl0tLBwso/Mjd5cgl1Tuzc+NK0nc/93tYyXF0g4fzfPE9kAOcRpcwSoaxX++K/37qliQNyfvreYVp545zst2S9CNsnld/JU/1MB/gOUnf/AB+dCfy9KX3wfvLsvzABG4lNKkqMsAxFHrpltduApLGtM1uCPvjet2EnL+Iz+sChEezWzP9PMB56XrwnTtBsqVhVWTbhlUl4IZ/+/yedjj2350P4dzkMZzVVVBmVBCltUQJWwcJ28Q9pW2Q5KVFlUqiMhkja15EtUUQkxJbkC2nKGUwMuexJitLrJJnOVbJuMbzCl0qlLAtyjZtVHuTjpIC+yxalqNsizWV0lhTMMNRTc1RWS5QJhOoguEJJbV4glCUoZhhqCbDPktJsU+hb5HyFmm1QanSxGu83xdeKM79Ma3+4n0YVHwccUI4m9HB4/0s84JZ+P0CWGKpjQplbmRzdHJlYW0KZW5kb2JqCjI5IDAgb2JqCjw8IC9Bc2NlbnQgNzUwIC9BdmdXaWR0aCA1MzYgL0NhcEhlaWdodCA3NTAgL0Rlc2NlbnQgLTI1MCAvRmxhZ3MgMzIgL0ZvbnRCQm94IFsgLTUxOSAtMjUwIDEyNjMgNzUwIF0gL0ZvbnRGaWxlMiA4MiAwIFIgL0ZvbnROYW1lIC9CQ0RGRUUrQ2FsaWJyaS1Cb2xkIC9Gb250V2VpZ2h0IDcwMCAvSXRhbGljQW5nbGUgMCAvTWF4V2lkdGggMTc4MSAvU3RlbVYgNTMgL1R5cGUgL0ZvbnREZXNjcmlwdG9yIC9YSGVpZ2h0IDI1MCA+PgplbmRvYmoKMzAgMCBvYmoKWyAyMjYgXQplbmRvYmoKMzEgMCBvYmoKPDwgL0FzY2VudCA3NTAgL0F2Z1dpZHRoIDUyMSAvQ2FwSGVpZ2h0IDc1MCAvRGVzY2VudCAtMjUwIC9GbGFncyAzMiAvRm9udEJCb3ggWyAtNTAzIC0yNTAgMTI0MCA3NTAgXSAvRm9udEZpbGUyIDgzIDAgUiAvRm9udE5hbWUgL0JDREdFRStDYWxpYnJpIC9Gb250V2VpZ2h0IDQwMCAvSXRhbGljQW5nbGUgMCAvTWF4V2lkdGggMTc0MyAvU3RlbVYgNTIgL1R5cGUgL0ZvbnREZXNjcmlwdG9yIC9YSGVpZ2h0IDI1MCA+PgplbmRvYmoKMzIgMCBvYmoKWyAyMjYgXQplbmRvYmoKMzMgMCBvYmoKPDwgL0sgWyAxIF0gL0xhbmcgKGVuKSAvUCAxOCAwIFIgL1BnIDggMCBSIC9TIC9TcGFuIC9UeXBlIC9TdHJ1Y3RFbGVtID4+CmVuZG9iagozNCAwIG9iago8PCAvSyBbIDMgXSAvTGFuZyAoZW4pIC9QIDE4IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjM1IDAgb2JqCjw8IC9LIFsgNSBdIC9MYW5nIChlbikgL1AgMTggMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKMzYgMCBvYmoKPDwgL0sgWyA3IF0gL0xhbmcgKGVuKSAvUCAxOCAwIFIgL1BnIDggMCBSIC9TIC9TcGFuIC9UeXBlIC9TdHJ1Y3RFbGVtID4+CmVuZG9iagozNyAwIG9iago8PCAvSyBbIDkgXSAvTGFuZyAoZW4pIC9QIDE4IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjM4IDAgb2JqCjw8IC9LIFsgMTEgXSAvTGFuZyAoZW4pIC9QIDE4IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjM5IDAgb2JqCjw8IC9LIFsgMTMgXSAvTGFuZyAoZW4pIC9QIDE4IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQwIDAgb2JqCjw8IC9LIFsgMTQgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQxIDAgb2JqCjw8IC9LIFsgMTUgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQyIDAgb2JqCjw8IC9LIFsgMTYgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQzIDAgb2JqCjw8IC9LIFsgMTcgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQ0IDAgb2JqCjw8IC9LIFsgMTkgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQ1IDAgb2JqCjw8IC9LIFsgMjEgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQ2IDAgb2JqCjw8IC9LIFsgMjIgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQ3IDAgb2JqCjw8IC9LIFsgMjMgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQ4IDAgb2JqCjw8IC9LIFsgMjQgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjQ5IDAgb2JqCjw8IC9LIFsgMjUgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjUwIDAgb2JqCjw8IC9LIFsgMjcgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjUxIDAgb2JqCjw8IC9LIFsgMjggXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjUyIDAgb2JqCjw8IC9LIFsgMjkgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjUzIDAgb2JqCjw8IC9LIFsgMzEgXSAvTGFuZyAoZW4pIC9QIDE5IDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjU0IDAgb2JqCjw8IC9LIDU5IC9QIDIwIDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjU1IDAgb2JqCjw8IC9LIFsgNjAgNzkgMCBSIF0gL1AgMjAgMCBSIC9QZyA4IDAgUiAvUyAvUCAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNTYgMCBvYmoKPDwgL0sgWyA4MCAwIFIgXSAvUCAyMCAwIFIgL1BnIDggMCBSIC9TIC9QIC9UeXBlIC9TdHJ1Y3RFbGVtID4+CmVuZG9iago1NyAwIG9iago8PCAvSyA1NiAvUCAyMSAwIFIgL1BnIDggMCBSIC9TIC9TcGFuIC9UeXBlIC9TdHJ1Y3RFbGVtID4+CmVuZG9iago1OCAwIG9iago8PCAvSyBbIDU3IDc4IDAgUiBdIC9QIDIxIDAgUiAvUGcgOCAwIFIgL1MgL1AgL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjU5IDAgb2JqCjw8IC9LIDUzIC9QIDIyIDAgUiAvUGcgOCAwIFIgL1MgL1NwYW4gL1R5cGUgL1N0cnVjdEVsZW0gPj4KZW5kb2JqCjYwIDAgb2JqCjw8IC9LIFsgNTQgNzcgMCBSIF0gL1AgMjIgMCBSIC9QZyA4IDAgUiAvUyAvUCAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjEgMCBvYmoKPDwgL0sgWyAzMiBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjIgMCBvYmoKPDwgL0sgWyAzMyBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjMgMCBvYmoKPDwgL0sgWyAzNCBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjQgMCBvYmoKPDwgL0sgWyAzNiBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjUgMCBvYmoKPDwgL0sgWyAzOCBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjYgMCBvYmoKPDwgL0sgWyAzOSBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjcgMCBvYmoKPDwgL0sgWyA0MCBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjggMCBvYmoKPDwgL0sgWyA0MiBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNjkgMCBvYmoKPDwgL0sgWyA0MyBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzAgMCBvYmoKPDwgL0sgWyA0NCBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzEgMCBvYmoKPDwgL0sgWyA0NSBdIC9MYW5nIChlbikgL1AgMjQgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzIgMCBvYmoKPDwgL0sgWyA0NiBdIC9MYW5nIChlbikgL1AgMjUgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzMgMCBvYmoKPDwgL0sgWyA0NyBdIC9MYW5nIChlbikgL1AgMjUgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzQgMCBvYmoKPDwgL0sgWyA0OCBdIC9MYW5nIChlbikgL1AgMjUgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzUgMCBvYmoKPDwgL0sgWyA1MCBdIC9MYW5nIChlbikgL1AgMjUgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzYgMCBvYmoKPDwgL0sgWyA1MSBdIC9MYW5nIChlbikgL1AgMjUgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlwZSAvU3RydWN0RWxlbSA+PgplbmRvYmoKNzcgMCBvYmoKPDwgL0sgWyA1NSBdIC9MYW5nIChlbikgL1AgNjAgMCBSIC9QZyA4IDAgUiAvUyAvU3BhbiAvVHlw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MDAgbiAKMDAwMDAxMDc5MiAwMDAwMCBuIAowMDAwMDExMDUyIDAwMDAwIG4gCjAwMDAwMTEwNzYgMDAwMDAgbiAKMDAwMDAxMTMzMSAwMDAwMCBuIAowMDAwMDExMzU1IDAwMDAwIG4gCjAwMDAwMTE0NDQgMDAwMDAgbiAKMDAwMDAxMTUzMyAwMDAwMCBuIAowMDAwMDExNjIyIDAwMDAwIG4gCjAwMDAwMTE3MTEgMDAwMDAgbiAKMDAwMDAxMTgwMCAwMDAwMCBuIAowMDAwMDExODkwIDAwMDAwIG4gCjAwMDAwMTE5ODAgMDAwMDAgbiAKMDAwMDAxMjA3MCAwMDAwMCBuIAowMDAwMDEyMTYwIDAwMDAwIG4gCjAwMDAwMTIyNTAgMDAwMDAgbiAKMDAwMDAxMjM0MCAwMDAwMCBuIAowMDAwMDEyNDMwIDAwMDAwIG4gCjAwMDAwMTI1MjAgMDAwMDAgbiAKMDAwMDAxMjYxMCAwMDAwMCBuIAowMDAwMDEyNzAwIDAwMDAwIG4gCjAwMDAwMTI3OTAgMDAwMDAgbiAKMDAwMDAxMjg4MCAwMDAwMCBuIAowMDAwMDEyOTcwIDAwMDAwIG4gCjAwMDAwMTMwNjAgMDAwMDAgbiAKMDAwMDAxMzE1MCAwMDAwMCBuIAowMDAwMDEzMjQwIDAwMDAwIG4gCjAwMDAwMTMzMTUgMDAwMDAgbiAKMDAwMDAxMzM5OCAwMDAwMCBuIAowMDAwMDEzNDc4IDAwMDAwIG4gCjAwMDAwMTM1NTMgMDAwMDAgbiAKMDAwMDAxMzYzNiAwMDAwMCBuIAowMDAwMDEzNzExIDAwMDAwIG4gCjAwMDAwMTM3OTQgMDAwMDAgbiAKMDAwMDAxMzg4NCAwMDAwMCBuIAowMDAwMDEzOTc0IDAwMDAwIG4gCjAwMDAwMTQwNjQgMDAwMDAgbiAKMDAwMDAxNDE1NCAwMDAwMCBuIAowMDAwMDE0MjQ0IDAwMDAwIG4gCjAwMDAwMTQzMzQgMDAwMDAgbiAKMDAwMDAxNDQyNCAwMDAwMCBuIAowMDAwMDE0NTE0IDAwMDAwIG4gCjAwMDAwMTQ2MDQgMDAwMDAgbiAKMDAwMDAxNDY5NCAwMDAwMCBuIAowMDAwMDE0Nzg0IDAwMDAwIG4gCjAwMDAwMTQ4NzQgMDAwMDAgbiAKMDAwMDAxNDk2NCAwMDAwMCBuIAowMDAwMDE1MDU0IDAwMDAwIG4gCjAwMDAwMTUxNDQgMDAwMDAgbiAKMDAwMDAxNTIzNCAwMDAwMCBuIAowMDAwMDE1MzI0IDAwMDAwIG4gCjAwMDAwMTU0MTQgMDAwMDAgbiAKMDAwMDAxNTUwNCAwMDAwMCBuIAowMDAwMDE1NTk0IDAwMDAwIG4gCjAwMDAwMTU3NjMgMDAwMDAgbiAKMDAwMDAzNDU2NyAwMDAwMCBuIAowMDAwMDUzOTQ4IDAwMDAwIG4gCjAwMDAwNTQwMjMgMDAwMDAgbiAKMDAwMDA1NDI4OCAwMDAwMCBuIAowMDAwMDU0MzMxIDAwMDAwIG4gCjAwMDAwNTQ3MTMgMDAwMDAgbiAKMDAwMDA1NTA5NSAwMDAwMCBuIAowMDAwMDg3MjM0IDAwMDAwIG4gCnRyYWlsZXIgPDwgL0luZm8gMiAwIFIgL1Jvb3QgMSAwIFIgL1NpemUgOTEgL0lEIFs8MjI0ZDYzZjFmMTIzMDk0Zjg4NTcwMmNiMjY5NDY0NDA+PDE0OWUwNWM2MGI5NTRiZTdmODA3NjdjM2VjOTQ5Njk4Pl0gPj4Kc3RhcnR4cmVmCjg3NjE2CiUlRU9GCg=='
# ---------- 读发票、拼打印 PDF 用的 C#（第一次用到报销功能时才编译） ----------
$DtPdfSource = @'
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;

// 桌面待办用的极简 PDF 工具：读发票上的文字（带位置）、把几张发票拼成一份 A4 打印文件
public class DtPdfRef { public int Num; public int Gen; public DtPdfRef(int n, int g) { Num = n; Gen = g; } }
public class DtPdfName { public string V; public DtPdfName(string v) { V = v; } public override string ToString() { return V; } }
public class DtPdfKw { public string V; public DtPdfKw(string v) { V = v; } public override string ToString() { return V; } }
public class DtPdfObj { public int Num; public object Val; public byte[] Raw; }
public class DtPdfItem { public double X; public double Y; public double Size; public string Text; public int Page; }

public class DtPdfDoc {
    public Dictionary<int, DtPdfObj> Objs = new Dictionary<int, DtPdfObj>();
    byte[] data;
    static Encoding Latin = Encoding.GetEncoding(28591);

    public DtPdfDoc(byte[] bytes) {
        data = bytes;
        string s = Latin.GetString(bytes);
        foreach (Match m in Regex.Matches(s, @"(?<![0-9])(\d+)\s+(\d+)\s+obj\b")) {
            try {
                int num = int.Parse(m.Groups[1].Value);
                int p = m.Index + m.Length;
                var ps = new DtPdfParser(bytes, p);
                object v = ps.ReadObject();
                var o = new DtPdfObj { Num = num, Val = v };
                int save = ps.Pos;
                object kw = ps.ReadObject();
                if (kw is DtPdfKw && ((DtPdfKw)kw).V == "stream") {
                    int q = ps.Pos;
                    if (q < bytes.Length && bytes[q] == '\r') q++;
                    if (q < bytes.Length && bytes[q] == '\n') q++;
                    int len = -1;
                    var d = v as Dictionary<string, object>;
                    if (d != null && d.ContainsKey("Length")) {
                        object lv = d["Length"];
                        if (lv is double) len = (int)(double)lv;
                    }
                    if (len < 0 || q + len > bytes.Length || !EndsStream(s, q + len)) {
                        int e = s.IndexOf("endstream", q, StringComparison.Ordinal);
                        if (e < 0) e = bytes.Length;
                        int ee = e;
                        if (ee > q && bytes[ee - 1] == '\n') ee--;
                        if (ee > q && bytes[ee - 1] == '\r') ee--;
                        len = ee - q;
                    }
                    o.Raw = new byte[len];
                    Array.Copy(bytes, q, o.Raw, 0, len);
                }
                Objs[num] = o;
            } catch { }
        }
        // 压缩对象流里的对象
        foreach (var o in new List<DtPdfObj>(Objs.Values)) {
            var d = o.Val as Dictionary<string, object>;
            if (d == null || o.Raw == null || Name(d, "Type") != "ObjStm") continue;
            try {
                byte[] dec = Decode(d, o.Raw);
                int n = (int)Num(d, "N"), first = (int)Num(d, "First");
                var hp = new DtPdfParser(dec, 0);
                var nums = new List<int>(); var offs = new List<int>();
                for (int i = 0; i < n; i++) { nums.Add((int)(double)hp.ReadObject()); offs.Add((int)(double)hp.ReadObject()); }
                for (int i = 0; i < n; i++) {
                    if (Objs.ContainsKey(nums[i]) && Objs[nums[i]].Raw != null) continue;
                    var vp = new DtPdfParser(dec, first + offs[i]);
                    Objs[nums[i]] = new DtPdfObj { Num = nums[i], Val = vp.ReadObject() };
                }
            } catch { }
        }
    }
    static bool EndsStream(string s, int at) {
        int i = at; while (i < s.Length && (s[i] == '\r' || s[i] == '\n' || s[i] == ' ')) i++;
        return string.CompareOrdinal(s, i, "endstream", 0, 9) == 0;
    }
    public object R(object v) {
        int guard = 0;
        while (v is DtPdfRef && guard++ < 20) { DtPdfObj o; if (!Objs.TryGetValue(((DtPdfRef)v).Num, out o)) return null; v = o.Val; }
        return v;
    }
    public DtPdfObj RObj(object v) { var r = v as DtPdfRef; DtPdfObj o; if (r != null && Objs.TryGetValue(r.Num, out o)) return o; return null; }
    public Dictionary<string, object> D(object v) { return R(v) as Dictionary<string, object>; }
    public static string Name(Dictionary<string, object> d, string k) { object v; if (d != null && d.TryGetValue(k, out v) && v is DtPdfName) return ((DtPdfName)v).V; return null; }
    public double Num(Dictionary<string, object> d, string k) { object v; if (d != null && d.TryGetValue(k, out v)) { v = R(v); if (v is double) return (double)v; } return 0; }

    public static byte[] Inflate(byte[] b) {
        foreach (int skip in new int[] { 2, 0 }) {
            try {
                if (b.Length <= skip) continue;
                using (var ms = new MemoryStream(b, skip, b.Length - skip))
                using (var ds = new DeflateStream(ms, CompressionMode.Decompress))
                using (var outp = new MemoryStream()) {
                    var buf = new byte[8192]; int n;
                    try { while ((n = ds.Read(buf, 0, buf.Length)) > 0) outp.Write(buf, 0, n); } catch { if (outp.Length == 0) throw; }
                    return outp.ToArray();
                }
            } catch { }
        }
        return new byte[0];
    }
    public byte[] Decode(Dictionary<string, object> d, byte[] raw) {
        object f; if (!d.TryGetValue("Filter", out f)) return raw;
        f = R(f);
        var names = new List<string>();
        if (f is DtPdfName) names.Add(((DtPdfName)f).V);
        else if (f is List<object>) foreach (var x in (List<object>)f) { var nx = R(x) as DtPdfName; if (nx != null) names.Add(nx.V); }
        byte[] cur = raw;
        foreach (var n in names) {
            if (n == "FlateDecode" || n == "Fl") cur = Inflate(cur);
            else return null;  // 图片等其它压缩方式：读文字用不上
        }
        object dp; d.TryGetValue("DecodeParms", out dp); var dpd = D(dp);
        if (dpd != null && Num(dpd, "Predictor") >= 10) cur = UnPng(cur, (int)Math.Max(1, Num(dpd, "Columns")));
        return cur;
    }
    static byte[] UnPng(byte[] b, int cols) {
        int row = cols + 1; var outp = new MemoryStream(); var prev = new byte[cols];
        for (int i = 0; i + row <= b.Length; i += row) {
            int t = b[i]; var cur = new byte[cols];
            for (int j = 0; j < cols; j++) {
                int x = b[i + 1 + j], a = j > 0 ? cur[j - 1] : 0, up = prev[j], c = j > 0 ? prev[j - 1] : 0;
                if (t == 1) x += a; else if (t == 2) x += up; else if (t == 3) x += (a + up) / 2;
                else if (t == 4) { int p = a + up - c, pa = Math.Abs(p - a), pb = Math.Abs(p - up), pc = Math.Abs(p - c); x += (pa <= pb && pa <= pc) ? a : (pb <= pc ? up : c); }
                cur[j] = (byte)x;
            }
            outp.Write(cur, 0, cols); prev = cur;
        }
        return outp.ToArray();
    }

    // ---------- 页面 ----------
    public List<Dictionary<string, object>> Pages() {
        var list = new List<Dictionary<string, object>>();
        Dictionary<string, object> cat = null;
        foreach (var o in Objs.Values) { var d = o.Val as Dictionary<string, object>; if (Name(d, "Type") == "Catalog") { cat = d; } }
        if (cat != null) { object p; if (cat.TryGetValue("Pages", out p)) Walk(D(p), list, 0); }
        if (list.Count == 0) {
            var keys = new List<int>(Objs.Keys); keys.Sort();
            foreach (var k in keys) { var d = Objs[k].Val as Dictionary<string, object>; if (Name(d, "Type") == "Page") list.Add(d); }
        }
        return list;
    }
    void Walk(Dictionary<string, object> node, List<Dictionary<string, object>> list, int depth) {
        if (node == null || depth > 30) return;
        if (Name(node, "Type") == "Page" || !node.ContainsKey("Kids")) { list.Add(node); return; }
        var kids = R(node["Kids"]) as List<object>;
        if (kids != null) foreach (var k in kids) Walk(D(k), list, depth + 1);
    }
    public object Inherited(Dictionary<string, object> page, string key) {
        var cur = page; int g = 0;
        while (cur != null && g++ < 30) { object v; if (cur.TryGetValue(key, out v)) return v; object p; if (!cur.TryGetValue("Parent", out p)) break; cur = D(p); }
        return null;
    }
    public double[] Box(Dictionary<string, object> page) {
        var a = R(Inherited(page, "CropBox")) as List<object>;
        if (a == null) a = R(Inherited(page, "MediaBox")) as List<object>;
        if (a == null || a.Count < 4) return new double[] { 0, 0, 595.28, 841.89 };
        var r = new double[4]; for (int i = 0; i < 4; i++) { object v = R(a[i]); r[i] = v is double ? (double)v : 0; }
        return new double[] { Math.Min(r[0], r[2]), Math.Min(r[1], r[3]), Math.Max(r[0], r[2]), Math.Max(r[1], r[3]) };
    }
    public byte[] PageContent(Dictionary<string, object> page) {
        object c; if (!page.TryGetValue("Contents", out c)) return new byte[0];
        var parts = new List<object>();
        object rc = R(c);
        if (rc is List<object>) parts.AddRange((List<object>)rc); else parts.Add(c);
        var ms = new MemoryStream();
        foreach (var p in parts) {
            var o = RObj(p); if (o == null || o.Raw == null) continue;
            var dec = Decode(o.Val as Dictionary<string, object>, o.Raw); if (dec == null) continue;
            ms.Write(dec, 0, dec.Length); ms.WriteByte((byte)'\n');
        }
        return ms.ToArray();
    }

    // ---------- 字体：把字节翻成文字 ----------
    class FontInfo { public Dictionary<string, string> Map; public HashSet<int> Lens = new HashSet<int>(); public Encoding Enc; public bool TwoByte; }
    Dictionary<object, FontInfo> fontCache = new Dictionary<object, FontInfo>();
    static Encoding Gbk() { try { return Encoding.GetEncoding(936); } catch { return null; } }
    FontInfo GetFont(object fref) {
        object key = fref is DtPdfRef ? (object)((DtPdfRef)fref).Num : fref;
        FontInfo fi; if (key != null && fontCache.TryGetValue(key, out fi)) return fi;
        fi = new FontInfo();
        var d = D(fref);
        if (d != null) {
            string sub = Name(d, "Subtype");
            object enc; d.TryGetValue("Encoding", out enc); enc = R(enc);
            string en = enc is DtPdfName ? ((DtPdfName)enc).V : "";
            if (sub == "Type0") {
                if (en.StartsWith("Identity")) fi.TwoByte = true;
                else if (en.Contains("UCS2") || en.Contains("UTF16")) fi.Enc = Encoding.BigEndianUnicode;
                else if (en.Contains("GB") && en.Contains("EUC") || en.Contains("GBK")) fi.Enc = Gbk();
                else fi.TwoByte = true;
            }
            object tu; if (d.TryGetValue("ToUnicode", out tu)) {
                var to = RObj(tu);
                if (to != null && to.Raw != null) { var dec = Decode(to.Val as Dictionary<string, object>, to.Raw); if (dec != null) ParseCMap(Latin.GetString(dec), fi); }
            }
        }
        if (key != null) fontCache[key] = fi;
        return fi;
    }
    static string HexToStr(string h) {
        h = Regex.Replace(h, @"\s", ""); if (h.Length % 4 != 0) h = h.PadRight((h.Length + 3) / 4 * 4, '0');
        var sb = new StringBuilder();
        for (int i = 0; i + 4 <= h.Length; i += 4) sb.Append((char)Convert.ToInt32(h.Substring(i, 4), 16));
        return sb.ToString();
    }
    static void ParseCMap(string t, FontInfo fi) {
        fi.Map = new Dictionary<string, string>();
        foreach (Match blk in Regex.Matches(t, @"beginbfchar(.*?)endbfchar", RegexOptions.Singleline))
            foreach (Match m in Regex.Matches(blk.Groups[1].Value, @"<([0-9A-Fa-f\s]+)>\s*<([0-9A-Fa-f\s]*)>")) {
                string src = Regex.Replace(m.Groups[1].Value, @"\s", "").ToUpperInvariant();
                fi.Map[src] = HexToStr(m.Groups[2].Value); fi.Lens.Add(src.Length / 2);
            }
        foreach (Match blk in Regex.Matches(t, @"beginbfrange(.*?)endbfrange", RegexOptions.Singleline))
            foreach (Match m in Regex.Matches(blk.Groups[1].Value, @"<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*(<[0-9A-Fa-f\s]*>|\[[^\]]*\])")) {
                string lo = m.Groups[1].Value, hi = m.Groups[2].Value; int w = lo.Length;
                long a = Convert.ToInt64(lo, 16), b = Convert.ToInt64(hi, 16);
                if (b - a > 65535) continue;
                fi.Lens.Add(w / 2);
                string dst = m.Groups[3].Value;
                if (dst.StartsWith("[")) {
                    int i = 0;
                    foreach (Match x in Regex.Matches(dst, @"<([0-9A-Fa-f\s]*)>")) { if (a + i > b) break; fi.Map[(a + i).ToString("X" + w)] = HexToStr(x.Groups[1].Value); i++; }
                } else {
                    string ds = HexToStr(dst.Trim('<', '>'));
                    for (long c = a; c <= b; c++) {
                        string v = ds.Length == 0 ? "" : ds.Substring(0, ds.Length - 1) + (char)(ds[ds.Length - 1] + (c - a));
                        fi.Map[c.ToString("X" + w)] = v;
                    }
                }
            }
    }
    static string DecodeStr(byte[] b, FontInfo fi) {
        if (fi.Map != null && fi.Map.Count > 0) {
            var sb = new StringBuilder(); int i = 0;
            var lens = new List<int>(fi.Lens); if (lens.Count == 0) lens.Add(fi.TwoByte ? 2 : 1);
            lens.Sort(); lens.Reverse();
            if (fi.TwoByte && !lens.Contains(2)) lens.Insert(0, 2);
            while (i < b.Length) {
                bool hit = false;
                foreach (int l in lens) {
                    if (i + l > b.Length) continue;
                    var k = new StringBuilder(); for (int j = 0; j < l; j++) k.Append(b[i + j].ToString("X2"));
                    string v; if (fi.Map.TryGetValue(k.ToString(), out v)) { sb.Append(v); i += l; hit = true; break; }
                }
                if (!hit) { if (fi.Enc != null) { sb.Append(fi.Enc.GetString(b, i, Math.Min(2, b.Length - i))); i += 2; } else { i += fi.TwoByte ? 2 : 1; } }
            }
            return sb.ToString();
        }
        if (fi.Enc != null) return fi.Enc.GetString(b);
        if (fi.TwoByte) return "";
        return Latin.GetString(b);
    }

    // ---------- 读文字 ----------
    static double[] Mul(double[] A, double[] B) {
        return new double[] {
            A[0] * B[0] + A[1] * B[2], A[0] * B[1] + A[1] * B[3],
            A[2] * B[0] + A[3] * B[2], A[2] * B[1] + A[3] * B[3],
            A[4] * B[0] + A[5] * B[2] + B[4], A[4] * B[1] + A[5] * B[3] + B[5] };
    }
    public List<DtPdfItem> Text(int maxPages) {
        var items = new List<DtPdfItem>();
        var pages = Pages();
        for (int pi = 0; pi < pages.Count && pi < maxPages; pi++) {
            var pg = pages[pi];
            var res = D(Inherited(pg, "Resources"));
            Run(PageContent(pg), res, new double[] { 1, 0, 0, 1, 0, 0 }, items, pi, 0);
        }
        return items;
    }
    void Run(byte[] content, Dictionary<string, object> res, double[] ctm0, List<DtPdfItem> items, int page, int depth) {
        if (depth > 8 || content == null) return;
        var fonts = res != null ? D(res.ContainsKey("Font") ? res["Font"] : null) : null;
        var xobjs = res != null ? D(res.ContainsKey("XObject") ? res["XObject"] : null) : null;
        var stack = new Stack<double[]>();
        double[] ctm = ctm0, tm = { 1, 0, 0, 1, 0, 0 }, tlm = { 1, 0, 0, 1, 0, 0 };
        double fs = 10, tl = 0; FontInfo font = new FontInfo();
        var ops = new List<object>();
        var p = new DtPdfParser(content, 0);
        while (true) {
            object t;
            try { t = p.ReadObject(); } catch { break; }
            if (t == null && p.AtEnd) break;
            var kw = t as DtPdfKw;
            if (kw == null) { ops.Add(t); continue; }
            string op = kw.V;
            try {
                switch (op) {
                    case "q": stack.Push(ctm); break;
                    case "Q": if (stack.Count > 0) ctm = stack.Pop(); break;
                    case "cm": if (ops.Count >= 6) ctm = Mul(Nums(ops, 6), ctm); break;
                    case "BT": tm = new double[] { 1, 0, 0, 1, 0, 0 }; tlm = tm; break;
                    case "Tf": if (ops.Count >= 2) { var fn = ops[ops.Count - 2] as DtPdfName; fs = ToD(ops[ops.Count - 1]); object fr = null; if (fn != null && fonts != null) fonts.TryGetValue(fn.V, out fr); font = fr != null ? GetFont(fr) : new FontInfo(); } break;
                    case "Tm": if (ops.Count >= 6) { tm = Nums(ops, 6); tlm = tm; } break;
                    case "Td": if (ops.Count >= 2) { var n = Nums(ops, 2); tlm = Mul(new double[] { 1, 0, 0, 1, n[0], n[1] }, tlm); tm = tlm; } break;
                    case "TD": if (ops.Count >= 2) { var n = Nums(ops, 2); tl = -n[1]; tlm = Mul(new double[] { 1, 0, 0, 1, n[0], n[1] }, tlm); tm = tlm; } break;
                    case "TL": if (ops.Count >= 1) tl = ToD(ops[ops.Count - 1]); break;
                    case "T*": tlm = Mul(new double[] { 1, 0, 0, 1, 0, -tl }, tlm); tm = tlm; break;
                    case "Tj": case "'": case "\"": case "TJ": {
                        if (op == "'" || op == "\"") { tlm = Mul(new double[] { 1, 0, 0, 1, 0, -tl }, tlm); tm = tlm; }
                        var sb = new StringBuilder();
                        if (op == "TJ") {
                            var arr = ops.Count > 0 ? ops[ops.Count - 1] as List<object> : null;
                            if (arr != null) foreach (var x in arr) {
                                if (x is byte[]) sb.Append(DecodeStr((byte[])x, font));
                                else if (x is double && (double)x < -350) sb.Append(' ');
                            }
                        } else if (ops.Count > 0 && ops[ops.Count - 1] is byte[]) sb.Append(DecodeStr((byte[])ops[ops.Count - 1], font));
                        string s = sb.ToString();
                        if (s.Trim().Length > 0) {
                            var m = Mul(tm, ctm);
                            double sc = Math.Sqrt(Math.Abs(m[0] * m[3] - m[1] * m[2]));
                            items.Add(new DtPdfItem { X = m[4], Y = m[5], Size = fs * sc, Text = s, Page = page });
                        }
                        // 没有字宽，往右挪一点，让同一行后面的字排在后面
                        tm = Mul(new double[] { 1, 0, 0, 1, s.Length * fs * 0.5, 0 }, tm);
                        break;
                    }
                    case "Do": {
                        var xn = ops.Count > 0 ? ops[ops.Count - 1] as DtPdfName : null;
                        object xr = null; if (xn != null && xobjs != null) xobjs.TryGetValue(xn.V, out xr);
                        var xo = RObj(xr);
                        var xd = xo != null ? xo.Val as Dictionary<string, object> : null;
                        if (xd != null && Name(xd, "Subtype") == "Form" && xo.Raw != null) {
                            var mx = R(xd.ContainsKey("Matrix") ? xd["Matrix"] : null) as List<object>;
                            double[] fm = { 1, 0, 0, 1, 0, 0 };
                            if (mx != null && mx.Count == 6) for (int i = 0; i < 6; i++) fm[i] = ToD(R(mx[i]));
                            var xres = D(xd.ContainsKey("Resources") ? xd["Resources"] : null) ?? res;
                            Run(Decode(xd, xo.Raw), xres, Mul(fm, ctm), items, page, depth + 1);
                        }
                        break;
                    }
                    case "BI": p.SkipInlineImage(); break;
                }
            } catch { }
            ops.Clear();
        }
    }
    static double ToD(object o) { return o is double ? (double)o : 0; }
    static double[] Nums(List<object> ops, int n) { var r = new double[n]; for (int i = 0; i < n; i++) r[i] = ToD(ops[ops.Count - n + i]); return r; }

    // 按行拼成文字：同一行（y 接近）的按 x 排
    public static string Lines(List<DtPdfItem> items) {
        var list = new List<DtPdfItem>(items);
        list.Sort((a, b) => a.Page != b.Page ? a.Page.CompareTo(b.Page) : b.Y.CompareTo(a.Y));
        var sb = new StringBuilder(); var line = new List<DtPdfItem>(); double ly = double.NaN; int lp = -1;
        foreach (var it in list) {
            double tol = Math.Max(2.0, Math.Min(it.Size, 12) * 0.45);
            if (line.Count > 0 && (it.Page != lp || Math.Abs(it.Y - ly) > tol)) { Flush(sb, line); line.Clear(); }
            if (line.Count == 0) { ly = it.Y; lp = it.Page; }
            line.Add(it);
        }
        Flush(sb, line);
        return sb.ToString();
    }
    static void Flush(StringBuilder sb, List<DtPdfItem> line) {
        if (line.Count == 0) return;
        var l = new List<DtPdfItem>(line);
        for (int i = 0; i < l.Count; i++) for (int j = i + 1; j < l.Count; j++) if (l[j].X < l[i].X - 0.01) { var t = l[i]; l[i] = l[j]; l[j] = t; }
        for (int i = 0; i < l.Count; i++) { if (i > 0) sb.Append(' '); sb.Append(l[i].Text.Trim()); }
        sb.Append('\n');
    }
    public static string ItemsJson(List<DtPdfItem> items) {
        var sb = new StringBuilder("[");
        for (int i = 0; i < items.Count; i++) {
            var it = items[i]; if (i > 0) sb.Append(',');
            sb.Append("{\"x\":").Append(it.X.ToString("0.##", CultureInfo.InvariantCulture)).Append(",\"y\":").Append(it.Y.ToString("0.##", CultureInfo.InvariantCulture))
              .Append(",\"s\":").Append(it.Size.ToString("0.##", CultureInfo.InvariantCulture)).Append(",\"p\":").Append(it.Page).Append(",\"t\":\"")
              .Append(it.Text.Replace("\\", "\\\\").Replace("\"", "\\\"").Replace("\r", " ").Replace("\n", " ")).Append("\"}");
        }
        return sb.Append(']').ToString();
    }
}

public class DtPdfParser {
    byte[] b; public int Pos;
    public DtPdfParser(byte[] bytes, int pos) { b = bytes; Pos = pos; }
    public bool AtEnd { get { SkipWs(); return Pos >= b.Length; } }
    static bool Ws(byte c) { return c == 0 || c == 9 || c == 10 || c == 12 || c == 13 || c == 32; }
    static bool Delim(byte c) { return c == '(' || c == ')' || c == '<' || c == '>' || c == '[' || c == ']' || c == '{' || c == '}' || c == '/' || c == '%'; }
    void SkipWs() {
        while (Pos < b.Length) {
            if (Ws(b[Pos])) { Pos++; continue; }
            if (b[Pos] == '%') { while (Pos < b.Length && b[Pos] != '\n' && b[Pos] != '\r') Pos++; continue; }
            break;
        }
    }
    public object ReadObject() {
        SkipWs();
        if (Pos >= b.Length) return null;
        byte c = b[Pos];
        if (c == '<' && Pos + 1 < b.Length && b[Pos + 1] == '<') {
            Pos += 2; var d = new Dictionary<string, object>();
            while (true) {
                SkipWs(); if (Pos >= b.Length) break;
                if (b[Pos] == '>' && Pos + 1 < b.Length && b[Pos + 1] == '>') { Pos += 2; break; }
                object k = ReadObject();
                var kn = k as DtPdfName; if (kn == null) { if (k == null) break; continue; }
                object v = ReadObject();
                d[kn.V] = v;
            }
            return d;
        }
        if (c == '<') {
            Pos++; var hex = new StringBuilder();
            while (Pos < b.Length && b[Pos] != '>') { char ch = (char)b[Pos++]; if (Uri.IsHexDigit(ch)) hex.Append(ch); }
            Pos++;
            if (hex.Length % 2 == 1) hex.Append('0');
            var r = new byte[hex.Length / 2];
            for (int i = 0; i < r.Length; i++) r[i] = Convert.ToByte(hex.ToString(i * 2, 2), 16);
            return r;
        }
        if (c == '(') {
            Pos++; var ms = new MemoryStream(); int depth = 1;
            while (Pos < b.Length) {
                byte x = b[Pos++];
                if (x == '\\' && Pos < b.Length) {
                    byte e = b[Pos++];
                    switch (e) {
                        case (byte)'n': ms.WriteByte(10); break; case (byte)'r': ms.WriteByte(13); break;
                        case (byte)'t': ms.WriteByte(9); break; case (byte)'b': ms.WriteByte(8); break; case (byte)'f': ms.WriteByte(12); break;
                        case (byte)'\r': if (Pos < b.Length && b[Pos] == '\n') Pos++; break;
                        case (byte)'\n': break;
                        default:
                            if (e >= '0' && e <= '7') {
                                int v = e - '0'; for (int k = 0; k < 2 && Pos < b.Length && b[Pos] >= '0' && b[Pos] <= '7'; k++) v = v * 8 + (b[Pos++] - '0');
                                ms.WriteByte((byte)v);
                            } else ms.WriteByte(e);
                            break;
                    }
                    continue;
                }
                if (x == '(') depth++;
                if (x == ')') { depth--; if (depth == 0) break; }
                ms.WriteByte(x);
            }
            return ms.ToArray();
        }
        if (c == '[') {
            Pos++; var l = new List<object>();
            while (true) {
                SkipWs(); if (Pos >= b.Length) break;
                if (b[Pos] == ']') { Pos++; break; }
                object v = ReadObject(); if (v == null && Pos >= b.Length) break;
                l.Add(v);
            }
            return l;
        }
        if (c == '/') {
            Pos++; var sb = new StringBuilder();
            while (Pos < b.Length && !Ws(b[Pos]) && !Delim(b[Pos])) {
                if (b[Pos] == '#' && Pos + 2 < b.Length) { try { sb.Append((char)Convert.ToByte(Encoding.ASCII.GetString(b, Pos + 1, 2), 16)); Pos += 3; continue; } catch { } }
                sb.Append((char)b[Pos++]);
            }
            // 名字里可能是 UTF-8 / GBK 字节，读文字时用不上，原样保留
            return new DtPdfName(sb.ToString());
        }
        if (c == ']' || c == '>' || c == ')' || c == '{' || c == '}') { Pos++; return new DtPdfKw(((char)c).ToString()); }
        if ((c >= '0' && c <= '9') || c == '-' || c == '+' || c == '.') {
            int st = Pos; Pos++;
            while (Pos < b.Length && ((b[Pos] >= '0' && b[Pos] <= '9') || b[Pos] == '.' || b[Pos] == '-')) Pos++;
            double v; double.TryParse(Encoding.ASCII.GetString(b, st, Pos - st), NumberStyles.Float, CultureInfo.InvariantCulture, out v);
            // 引用：12 0 R
            int save = Pos;
            if (v >= 0 && v == Math.Floor(v)) {
                int q = Pos; while (q < b.Length && Ws(b[q])) q++;
                int g0 = q; while (q < b.Length && b[q] >= '0' && b[q] <= '9') q++;
                if (q > g0) {
                    int g1 = q; while (q < b.Length && Ws(b[q])) q++;
                    if (q < b.Length && b[q] == 'R' && (q + 1 >= b.Length || Ws(b[q + 1]) || Delim(b[q + 1]))) {
                        Pos = q + 1; return new DtPdfRef((int)v, int.Parse(Encoding.ASCII.GetString(b, g0, g1 - g0)));
                    }
                }
            }
            Pos = save; return v;
        }
        {
            int st = Pos;
            while (Pos < b.Length && !Ws(b[Pos]) && !Delim(b[Pos])) Pos++;
            if (Pos == st) { Pos++; return new DtPdfKw(((char)c).ToString()); }
            string w = Encoding.ASCII.GetString(b, st, Pos - st);
            if (w == "true") return true; if (w == "false") return false; if (w == "null") return null;
            return new DtPdfKw(w);
        }
    }
    public void SkipInlineImage() {
        // BI ... ID <二进制> EI
        while (Pos < b.Length) {
            object t = ReadObject();
            if (t is DtPdfKw && ((DtPdfKw)t).V == "ID") break;
            if (t == null && Pos >= b.Length) return;
        }
        Pos++;
        while (Pos + 2 < b.Length) {
            if (b[Pos] == 'E' && b[Pos + 1] == 'I' && (Pos == 0 || Ws(b[Pos - 1])) && (Pos + 2 >= b.Length || Ws(b[Pos + 2]))) { Pos += 2; return; }
            Pos++;
        }
        Pos = b.Length;
    }
}

// OFD：其实是个压缩包，文字在 Pages/Tpls 的 Content.xml 里，单位是毫米、y 向下
public static class DtOfd {
    public static List<DtPdfItem> Text(byte[] bytes) {
        var items = new List<DtPdfItem>();
        using (var ms = new MemoryStream(bytes))
        using (var z = new System.IO.Compression.ZipArchive(ms, System.IO.Compression.ZipArchiveMode.Read)) {
            foreach (var e in z.Entries) {
                string n = e.FullName.Replace('\\', '/');
                if (!n.EndsWith("Content.xml") || !(n.Contains("/Pages/") || n.Contains("/Tpls/"))) continue;
                var doc = new System.Xml.XmlDocument();
                using (var s = e.Open()) doc.Load(s);
                foreach (System.Xml.XmlNode to in doc.SelectNodes("//*[local-name()='TextObject']")) {
                    double bx = 0, by = 0, size = 3;
                    var ba = to.Attributes["Boundary"];
                    if (ba != null) { var p = ba.Value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries); if (p.Length >= 2) { double.TryParse(p[0], NumberStyles.Float, CultureInfo.InvariantCulture, out bx); double.TryParse(p[1], NumberStyles.Float, CultureInfo.InvariantCulture, out by); } }
                    var sa = to.Attributes["Size"]; if (sa != null) double.TryParse(sa.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out size);
                    foreach (System.Xml.XmlNode tc in to.ChildNodes) {
                        if (tc.LocalName != "TextCode") continue;
                        double x = 0, y = 0;
                        var xa = tc.Attributes["X"]; if (xa != null) double.TryParse(xa.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out x);
                        var ya = tc.Attributes["Y"]; if (ya != null) double.TryParse(ya.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out y);
                        string t = tc.InnerText; if (t.Trim().Length == 0) continue;
                        items.Add(new DtPdfItem { X = (bx + x) * 2.8346, Y = -(by + y) * 2.8346, Size = size * 2.8346, Text = t, Page = 0 });
                    }
                }
            }
        }
        return items;
    }
}

// 写一个打印用的 PDF：每张发票各占一页 A4（缩放到合适大小、靠上居中），前面可以加一页粘贴单
public class DtPdfOut {
    List<byte[]> objs = new List<byte[]>();
    List<int> pages = new List<int>();
    int pagesNum, fontNum = 0;
    static Encoding Latin = Encoding.GetEncoding(28591);
    public DtPdfOut() { pagesNum = Reserve(); }
    int Reserve() { objs.Add(null); return objs.Count; }
    void Set(int num, string body) { objs[num - 1] = Latin.GetBytes(body); }
    void SetStream(int num, string dict, byte[] data) {
        var ms = new MemoryStream(); var a = Latin.GetBytes(dict + "\nstream\n"); ms.Write(a, 0, a.Length); ms.Write(data, 0, data.Length);
        var b = Latin.GetBytes("\nendstream"); ms.Write(b, 0, b.Length); objs[num - 1] = ms.ToArray();
    }
    static string F(double d) { return d.ToString("0.####", CultureInfo.InvariantCulture); }
    static string NameOut(string n) {
        var sb = new StringBuilder("/");
        foreach (char c in n) { int v = c; if (v > 32 && v < 127 && "()<>[]{}/%#".IndexOf(c) < 0) sb.Append(c); else sb.Append('#').Append((v & 0xFF).ToString("X2")); }
        return sb.ToString();
    }
    string Ser(object v, DtPdfDoc src, Dictionary<int, int> map, int depth) {
        if (depth > 60) return "null";
        if (v == null) return "null";
        if (v is bool) return (bool)v ? "true" : "false";
        if (v is double) return F((double)v);
        if (v is DtPdfName) return NameOut(((DtPdfName)v).V);
        if (v is byte[]) { var sb = new StringBuilder("<"); foreach (var x in (byte[])v) sb.Append(x.ToString("X2")); return sb.Append('>').ToString(); }
        if (v is DtPdfRef) return Copy((DtPdfRef)v, src, map, depth + 1) + " 0 R";
        if (v is List<object>) { var sb = new StringBuilder("["); foreach (var x in (List<object>)v) sb.Append(Ser(x, src, map, depth + 1)).Append(' '); return sb.Append(']').ToString(); }
        if (v is Dictionary<string, object>) {
            var sb = new StringBuilder("<<");
            foreach (var kv in (Dictionary<string, object>)v) { if (kv.Key == "Parent") continue; sb.Append(NameOut(kv.Key)).Append(' ').Append(Ser(kv.Value, src, map, depth + 1)).Append(' '); }
            return sb.Append(">>").ToString();
        }
        if (v is DtPdfKw) return ((DtPdfKw)v).V;
        return "null";
    }
    int Copy(DtPdfRef r, DtPdfDoc src, Dictionary<int, int> map, int depth) {
        int nn; if (map.TryGetValue(r.Num, out nn)) return nn;
        nn = Reserve(); map[r.Num] = nn;
        DtPdfObj o; if (!src.Objs.TryGetValue(r.Num, out o)) { Set(nn, "null"); return nn; }
        if (o.Raw != null) {
            var d = new Dictionary<string, object>(o.Val as Dictionary<string, object> ?? new Dictionary<string, object>());
            d["Length"] = (double)o.Raw.Length;
            SetStream(nn, Ser(d, src, map, depth), o.Raw);
        } else Set(nn, Ser(o.Val, src, map, depth));
        return nn;
    }
    // 把源 PDF 的第一页做成一个可以缩放摆放的 Form
    int PageForm(DtPdfDoc doc, out double[] box) {
        var pg = doc.Pages()[0];
        box = doc.Box(pg);
        var map = new Dictionary<int, int>();
        string res = Ser(doc.Inherited(pg, "Resources") ?? new Dictionary<string, object>(), doc, map, 0);
        byte[] content = doc.PageContent(pg);
        int f = Reserve();
        SetStream(f, "<< /Type /XObject /Subtype /Form /BBox [" + F(box[0]) + " " + F(box[1]) + " " + F(box[2]) + " " + F(box[3]) + "] /Resources " + res + " /Length " + content.Length + " >>", content);
        return f;
    }
    void AddPage(double w, double h, string content, string xobjs, bool withFont) {
        int c = Reserve(); var cb = Latin.GetBytes(content);
        SetStream(c, "<< /Length " + cb.Length + " >>", cb);
        int p = Reserve();
        string res = "<< /XObject << " + xobjs + " >>" + (withFont ? " /Font << /FZ " + Font() + " 0 R >>" : "") + " >>";
        Set(p, "<< /Type /Page /Parent " + pagesNum + " 0 R /MediaBox [0 0 " + F(w) + " " + F(h) + "] /Resources " + res + " /Contents " + c + " 0 R >>");
        pages.Add(p);
    }
    int Font() {
        if (fontNum > 0) return fontNum;
        int fd = Reserve();
        Set(fd, "<< /Type /FontDescriptor /FontName /STSong-Light /Flags 6 /FontBBox [-25 -254 1000 880] /ItalicAngle 0 /Ascent 880 /Descent -120 /CapHeight 880 /StemV 93 >>");
        int cid = Reserve();
        Set(cid, "<< /Type /Font /Subtype /CIDFontType0 /BaseFont /STSong-Light /CIDSystemInfo << /Registry (Adobe) /Ordering (GB1) /Supplement 4 >> /FontDescriptor " + fd + " 0 R /DW 1000 /W [1 95 500] >>");
        fontNum = Reserve();
        Set(fontNum, "<< /Type /Font /Subtype /Type0 /BaseFont /STSong-Light /Encoding /UniGB-UCS2-H /DescendantFonts [" + cid + " 0 R] >>");
        return fontNum;
    }
    static string Hex(string s) { var sb = new StringBuilder("<"); foreach (char c in s) sb.Append(((int)c).ToString("X4")); return sb.Append('>').ToString(); }

    // 一张发票一页 A4 竖版：靠上、左右居中，最大不放大
    public void AddInvoice(byte[] pdf, string caption) {
        var doc = new DtPdfDoc(pdf);
        if (doc.Pages().Count == 0) throw new Exception("这个文件里没找到页面");
        double[] b; int f = PageForm(doc, out b);
        double W = 595.28, H = 841.89, m = 36;
        double w = b[2] - b[0], h = b[3] - b[1];
        double s = Math.Min(1.0, Math.Min((W - 2 * m) / w, (H - 2 * m - 20) / h));
        double tx = (W - w * s) / 2 - b[0] * s, ty = H - m - h * s - b[1] * s;
        string c = "q " + F(s) + " 0 0 " + F(s) + " " + F(tx) + " " + F(ty) + " cm /Fm0 Do Q\n";
        if (!string.IsNullOrEmpty(caption)) c += "BT /FZ 9 Tf 0.4 g " + F(m) + " " + F(ty + b[1] * s - 16) + " Td " + Hex(caption) + " Tj ET\n";
        AddPage(W, H, c, "/Fm0 " + f + " 0 R", !string.IsNullOrEmpty(caption));
    }
    // 粘贴单：用模板 PDF 当底，上面写字。texts: x, y, size, 文字
    public void AddTemplate(byte[] pdf, List<DtPdfItem> texts) {
        var doc = new DtPdfDoc(pdf);
        double[] b; int f = PageForm(doc, out b);
        var sb = new StringBuilder("q 1 0 0 1 " + F(-b[0]) + " " + F(-b[1]) + " cm /Fm0 Do Q\n");
        foreach (var t in texts) sb.Append("BT /FZ " + F(t.Size) + " Tf 0 g " + F(t.X) + " " + F(t.Y) + " Td " + Hex(t.Text) + " Tj ET\n");
        AddPage(b[2] - b[0], b[3] - b[1], sb.ToString(), "/Fm0 " + f + " 0 R", true);
    }
    // 没有模板时，自己画一张简单的粘贴单（A4 横版）
    public void AddBlankSheet(List<DtPdfItem> texts, List<double[]> lines) {
        var sb = new StringBuilder("0 g 0.8 w\n");
        foreach (var l in lines) sb.Append(F(l[0]) + " " + F(l[1]) + " m " + F(l[2]) + " " + F(l[3]) + " l S\n");
        foreach (var t in texts) sb.Append("BT /FZ " + F(t.Size) + " Tf " + F(t.X) + " " + F(t.Y) + " Td " + Hex(t.Text) + " Tj ET\n");
        AddPage(841.89, 595.28, sb.ToString(), "", true);
    }
    // 图片（PS 那边统一转成 RGB 的 JPG 再给过来）：按 A4 竖版放，最大铺满可用区域
    public void AddJpeg(byte[] jpg, int pw, int ph, string caption) {
        int im = Reserve();
        SetStream(im, "<< /Type /XObject /Subtype /Image /Width " + pw + " /Height " + ph + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length " + jpg.Length + " >>", jpg);
        double W = 595.28, H = 841.89, m = 36;
        double w = pw * 0.75, h = ph * 0.75;
        double s = Math.Min((W - 2 * m) / w, (H - 2 * m - 20) / h);
        if (s > 1.6) s = 1.6;
        double dw = w * s, dh = h * s, tx = (W - dw) / 2, ty = H - m - dh;
        string c = "q " + F(dw) + " 0 0 " + F(dh) + " " + F(tx) + " " + F(ty) + " cm /Im0 Do Q\n";
        if (!string.IsNullOrEmpty(caption)) c += "BT /FZ 9 Tf 0.4 g " + F(m) + " " + F(ty - 16) + " Td " + Hex(caption) + " Tj ET\n";
        AddPage(W, H, c, "/Im0 " + im + " 0 R", !string.IsNullOrEmpty(caption));
    }
    public int PageCount { get { return pages.Count; } }
    public byte[] Save() {
        var kids = new StringBuilder(); foreach (var p in pages) kids.Append(p + " 0 R ");
        Set(pagesNum, "<< /Type /Pages /Kids [" + kids + "] /Count " + pages.Count + " >>");
        int cat = Reserve(); Set(cat, "<< /Type /Catalog /Pages " + pagesNum + " 0 R >>");
        var ms = new MemoryStream();
        var head = new byte[] { 0x25, 0x50, 0x44, 0x46, 0x2D, 0x31, 0x2E, 0x37, 0x0A, 0x25, 0xE2, 0xE3, 0xCF, 0xD3, 0x0A };
        ms.Write(head, 0, head.Length);
        var offs = new long[objs.Count];
        for (int i = 0; i < objs.Count; i++) {
            offs[i] = ms.Position;
            var a = Latin.GetBytes((i + 1) + " 0 obj\n"); ms.Write(a, 0, a.Length);
            var body = objs[i] ?? Latin.GetBytes("null"); ms.Write(body, 0, body.Length);
            var z = Latin.GetBytes("\nendobj\n"); ms.Write(z, 0, z.Length);
        }
        long x = ms.Position;
        var sb = new StringBuilder("xref\n0 " + (objs.Count + 1) + "\n0000000000 65535 f \n");
        foreach (var o in offs) sb.Append(o.ToString("D10") + " 00000 n \n");
        sb.Append("trailer\n<< /Size " + (objs.Count + 1) + " /Root " + cat + " 0 R >>\nstartxref\n" + x + "\n%%EOF\n");
        var t = Latin.GetBytes(sb.ToString()); ms.Write(t, 0, t.Length);
        return ms.ToArray();
    }
}
'@
# ---------- 报销：读发票 ----------
function ConvertTo-IsoDay([string]$s) {
    if ($s -match '(\d{4})\s*[年\-/.]\s*(\d{1,2})\s*[月\-/.]\s*(\d{1,2})') { return ('{0}-{1:D2}-{2:D2}' -f [int]$Matches[1], [int]$Matches[2], [int]$Matches[3]) }
    return ''
}
function Get-CleanPlace([string]$s) { return ($s.Trim() -replace '\s+', '' -replace '站$', '') }

# lines：按版面从上到下排好的文字行。返回一个 hashtable，认不出来的字段就空着
function Parse-InvoiceText([string[]]$lines) {
    $r = [ordered]@{ invType = ''; invNo = ''; invDate = ''; amount = ''; person = ''; cat = ''; date = ''; time = ''; from = ''; to = ''; trip = ''; seller = ''; item = ''; nights = ''; legs = 0; date2 = ''; remark = '' }
    $all = ($lines -join "`n")
    $flat = $all -replace '[ \t　]+', ' '
    if ($flat -match '发票号码\s*[:：]?\s*(\d{20})') { $r.invNo = $Matches[1] }
    elseif ($flat -match '(?<!\d)(\d{20})(?!\d)') { $r.invNo = $Matches[1] }
    if ($flat -match '(开票日期|填开日期)\s*[:：]?\s*(\d{4}\s*年\s*\d{1,2}\s*月\s*\d{1,2}\s*日|\d{4}-\d{2}-\d{2})') { $r.invDate = ConvertTo-IsoDay $Matches[2] }

    if ($flat -match '铁路电子客票') {
        $r.invType = '铁路电子客票'; $r.cat = '城市间交通费'
        foreach ($l in $lines) {
            if (-not $r.from -and $l -match '^\s*(\S+)\s*站\s+([GDCKTZYSL]?\d{1,5})\s+(\S+)\s*站') { $r.from = Get-CleanPlace $Matches[1]; $r.trip = $Matches[2]; $r.to = Get-CleanPlace $Matches[3] }
            if (-not $r.date -and $l -match '(\d{4})年(\d{1,2})月(\d{1,2})日\s*(\d{1,2}:\d{2})\s*开') { $r.date = ConvertTo-IsoDay $l; $r.time = $Matches[4] }
            if (-not $r.person -and $l -match '\*{3,}\w{3,4}\s+(\S{2,6})\s*$') { $r.person = $Matches[1] }
        }
        if ($flat -match '(票价|退票费|改签费|改签差额)\s*[:：]?\s*[￥¥]?\s*([\d,]+\.\d{2})') { $r.amount = $Matches[2] -replace ',', ''; if ($Matches[1] -ne '票价') { $r.item = $Matches[1]; $r.cat = '其他' } }
        if ($r.trip -and $r.item -eq '') { $r.item = '高铁/火车 ' + $r.trip }
    }
    elseif ($flat -match '航空运输电子客票行程单|航空.*行程单') {
        $r.invType = '机票电子行程单'; $r.cat = '城市间交通费'
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $l = $lines[$i]
            if (-not $r.from -and $l -match '^\s*自\s*[:：]\s*(.+)$') {
                $rest = $Matches[1]
                if ($rest -match '^(.*?)\s+(\S+)\s+([A-Z0-9]{2}\d{3,4})\b') { $tk = @($Matches[1] -split '\s+' | Where-Object { $_ }); $r.from = $tk[0]; $r.trip = $Matches[3] }
                if ($rest -match '(\d{4})年(\d{1,2})月(\d{1,2})日\s*(\d{1,2}:\d{2})') { $r.date = ConvertTo-IsoDay $rest; $r.time = $Matches[4] }
                for ($j = $i + 1; $j -lt $lines.Count; $j++) {
                    if ($lines[$j] -match '^\s*至\s*[:：]\s*([^\sC][^\s]*)') {
                        $r.legs++
                        if ($r.legs -eq 1) { $r.to = $Matches[1] }
                    }
                }
            }
            if (-not $r.person -and $l -match '^\s*(\S{2,6})\s+\S*\*{4,}\S*') { $r.person = $Matches[1] }
        }
        $cny = [regex]::Matches($flat, 'CNY\s*([\d,]+\.\d{2})')
        if ($cny.Count -gt 0) { $r.amount = $cny[$cny.Count - 1].Groups[1].Value -replace ',', '' }
        if ($flat -match '填开单位\s*[:：]\s*(\S+)') { $r.seller = $Matches[1] }
        if ($r.trip) { $r.item = '飞机 ' + $r.trip }
    }
    else {
        $r.invType = if ($flat -match '专用发票') { '电子专用发票' } elseif ($flat -match '普通发票|电子发票') { '电子普通发票' } else { '其他票据' }
        if ($flat -match '[（(]\s*小写\s*[)）]\s*[￥¥]?\s*([\d,]+\.\d{2})') { $r.amount = $Matches[1] -replace ',', '' }
        else {
            $mx = 0.0
            foreach ($m in [regex]::Matches($flat, '[￥¥]\s*([\d,]+\.\d{2})')) { $v = [double]($m.Groups[1].Value -replace ',', ''); if ($v -gt $mx) { $mx = $v } }
            if ($mx -gt 0) { $r.amount = $mx.ToString('0.00', [Globalization.CultureInfo]::InvariantCulture) }
        }
        $names = [regex]::Matches($flat, '名\s*称\s*[:：]\s*(\S+)')
        if ($names.Count -ge 2) { $r.seller = $names[1].Groups[1].Value }
        if ($flat -match '\*([^*\s]+)\*(\S+)') { $r.item = $Matches[2]; $kind = $Matches[1] } else { $kind = '' }
        $hay = $kind + ' ' + $r.item + ' ' + $r.seller
        if ($hay -match '住宿|酒店|宾馆|饭店|旅馆') { $r.cat = '住宿费' }
        elseif ($hay -match '餐饮|餐费|食品') { $r.cat = '餐费' }
        elseif ($hay -match '会议|会务') { $r.cat = '会议/会务费' }
        elseif ($hay -match '印刷|复印|资料|图书') { $r.cat = '资料/印刷费' }
        elseif ($hay -match '租车|包车|租赁') { $r.cat = '租车/包车费' }
        elseif ($hay -match '运输|客运|出行|出租|打车|交通') { $r.cat = '市内交通费' }
        else { $r.cat = '其他' }
        # 旅客运输服务的发票下面有出行人表格：出行人 证件号 出行日期 出发地 到达地 等级 交通工具
        for ($i = 0; $i -lt $lines.Count; $i++) {
            $l = $lines[$i]
            if ($l -match '^\s*(\S{2,6})\s+\S*\*{4,}\S*\s+(\d{4}-\d{2}-\d{2})\s+(\S+)\s+(\S+)(.*)$') {
                if (-not $r.person) {
                    $r.person = $Matches[1]; $r.date = $Matches[2]; $r.from = $Matches[3]; $r.to = $Matches[4]
                    $tail = $Matches[5].Trim()
                    # 到达地太长会折到下一行
                    if ($i + 1 -lt $lines.Count) {
                        $nx = $lines[$i + 1].Trim()
                        if ($nx -and $nx -notmatch '\s' -and $nx -notmatch '\d|价税|合计|备注|开票') { $r.to += $nx }
                    }
                    if ($tail -match '汽车|大巴|客车' -and $r.cat -eq '市内交通费') {
                        $d0 = 0.0; [void][double]::TryParse($r.amount, [ref]$d0)
                        if ($d0 -ge 100) { $r.cat = '租车/包车费' }
                    }
                }
            }
        }
        # 备注：价税合计下面那几行（有的发票有，有的没有）；销售方地址、银行账号这些不要
        $rm = New-Object System.Collections.ArrayList; $on = $false
        foreach ($l in $lines) {
            if ($l -match '价税合计') { $on = $true; continue }
            if (-not $on) { continue }
            $t0 = $l.Trim()
            if ($t0 -match '^开票人|^收款人|^复核') { break }
            if ($t0 -match '^[备注]$' -or $t0 -match '地址|电话|开户|银行|账号') { continue }
            [void]$rm.Add($t0)
        }
        $r.remark = ($rm -join ' ')
        if ($r.cat -eq '住宿费') {
            if ($all -match '(天|晚|间夜|夜)\s+(\d+)(\s|$)') { $r.nights = $Matches[2] }
            elseif ($all -match '(\d+)\s*(晚|间夜)') { $r.nights = $Matches[1] }
            $r.to = ($r.seller -replace '(有限责任公司|股份有限公司|有限公司|分公司)$', '')
            # 备注里写了入住、离店日期就用它，比如「10 月 15 至 10 月 16」「2026-10-15至2026-10-16」
            $y = if ($r.invDate) { [int]$r.invDate.Substring(0, 4) } else { (Get-Date).Year }
            $im = if ($r.invDate) { [int]$r.invDate.Substring(5, 2) } else { 12 }
            $d1 = $null; $d2 = $null
            if ($r.remark -match '(\d{4})\s*[年\-./]\s*(\d{1,2})\s*[月\-./]\s*(\d{1,2})\s*日?\s*(?:至|到|~|～|—|－|-)\s*(?:(\d{4})\s*[年\-./]\s*)?(\d{1,2})\s*[月\-./]\s*(\d{1,2})') {
                try { $d1 = [datetime]::new([int]$Matches[1], [int]$Matches[2], [int]$Matches[3]); $d2 = [datetime]::new($(if ($Matches[4]) { [int]$Matches[4] } else { [int]$Matches[1] }), [int]$Matches[5], [int]$Matches[6]) } catch {}
            } elseif ($r.remark -match '(\d{1,2})\s*月\s*(\d{1,2})\s*[日号]?\s*(?:至|到|~|～|—|－|-)\s*(?:(\d{1,2})\s*月\s*)?(\d{1,2})\s*[日号]?') {
                $m1 = [int]$Matches[1]; $m2 = if ($Matches[3]) { [int]$Matches[3] } else { $m1 }
                try {
                    $y1 = if ($m1 -gt $im) { $y - 1 } else { $y }; $y2 = if ($m2 -gt $im) { $y - 1 } else { $y }
                    $d1 = [datetime]::new($y1, $m1, [int]$Matches[2]); $d2 = [datetime]::new($y2, $m2, [int]$Matches[4])
                } catch {}
            }
            if ($d1 -and $d2 -and $d2 -ge $d1) {
                $r.date = $d1.ToString('yyyy-MM-dd'); $r.date2 = $d2.ToString('yyyy-MM-dd')
                $n = ($d2 - $d1).Days; if ($n -gt 0) { $r.nights = [string]$n }
            }
        }
    }
    return $r
}

# ==================== 报销 ====================
# 报销单存在 bills.json，每一笔费用存在 expenses.json，发票文件放在 同步文件夹\报销附件\<单号>\ 下面，跟着网盘一起同步
$ReimbCats = @('城市间交通费', '住宿费', '市内交通费', '租车/包车费', '餐费', '出差补助', '会议/会务费', '资料/印刷费', '专家咨询费', '其他')
$BillStates = @('未提交', '已提交', '已退回', '已到账')
$TicketStates = @('已取得', '待补', '无票')
$DefaultDept = ''
$script:Bills = New-Object System.Collections.ArrayList
$script:Exps = New-Object System.Collections.ArrayList
$script:ReimbSel = ''
$script:ReimbEdit = ''
$script:ReimbView = 'list'
$script:ReimbFilter = ''

function ConvertTo-Bill($o) {
    $id = [string]$o.id; if (-not $id) { $id = [guid]::NewGuid().ToString('N') }
    $st = [string]$o.state; if ($BillStates -notcontains $st) { $st = '未提交' }
    [pscustomobject]@{
        id = $id; text = [string]$o.text; no = [string]$o.no; oaNo = [string]$o.oaNo; proj = [string]$o.proj; place = [string]$o.place
        start = [string]$o.start; end = [string]$o.end; person = [string]$o.person; people = [string]$o.people; dept = [string]$o.dept
        state = $st; submit = [string]$o.submit; note = [string]$o.note; created = [string]$o.created; upd = [long]$o.upd
    }
}
function ConvertTo-Expense($o) {
    $id = [string]$o.id; if (-not $id) { $id = [guid]::NewGuid().ToString('N') }
    $cat = [string]$o.cat; if ($ReimbCats -notcontains $cat) { $cat = '其他' }
    $ts = [string]$o.status; if ($TicketStates -notcontains $ts) { $ts = '待补' }
    [pscustomobject]@{
        id = $id; text = [string]$o.text; bill = [string]$o.bill; cat = $cat
        date = [string]$o.date; time = [string]$o.time; date2 = [string]$o.date2; time2 = [string]$o.time2
        from = [string]$o.from; to = [string]$o.to; amount = [string]$o.amount; person = [string]$o.person
        invNo = [string]$o.invNo; invDate = [string]$o.invDate; invType = [string]$o.invType; seller = [string]$o.seller
        file = [string]$o.file; nights = [string]$o.nights; seg = [bool]$o.seg; status = $ts; grp = [string]$o.grp
        note = [string]$o.note; created = [string]$o.created; upd = [long]$o.upd
    }
}
function Save-Bills { Request-Save $script:BillStore }
function Save-Exps { Request-Save $script:ExpStore }
function Find-Bill([string]$id) { foreach ($b in $script:Bills) { if ($b.id -eq $id) { return $b } }; return $null }
function Find-Exp([string]$id) { foreach ($x in $script:Exps) { if ($x.id -eq $id) { return $x } }; return $null }

function Get-Amt($x) {
    $d = 0.0
    $s = if ($x -is [string]) { $x } else { [string]$x.amount }
    [void][double]::TryParse(($s -replace '[,，¥￥元\s]', ''), [Globalization.NumberStyles]::Float, $Inv, [ref]$d)
    return $d
}
function Format-Amt([double]$d) { return $d.ToString('#,##0.00', $Inv) }
function Format-AmtRaw([double]$d) { return $d.ToString('0.00', $Inv) }
function Format-AmtShort([double]$d) { return $d.ToString('0.##', $Inv) }

function Get-ReimbDept { $d = [string]$script:S.reimbDept; if ($d.Trim()) { return $d.Trim() }; return $DefaultDept }
function Get-BillExps([string]$billId) {
    return @($script:Exps | Where-Object { $_.bill -eq $billId } | Sort-Object @{ Expression = { [string]$_.date } }, @{ Expression = { if ($_.time) { [string]$_.time } else { '99' } } }, @{ Expression = { [string]$_.created } })
}
function Get-BillTotal([string]$billId) { $t = 0.0; foreach ($x in $script:Exps) { if ($x.bill -eq $billId) { $t += Get-Amt $x } }; return $t }
function Get-BillFiles([string]$billId) {
    $seen = @{}; $out = New-Object System.Collections.ArrayList
    foreach ($x in (Get-BillExps $billId)) { if ($x.file -and -not $seen.ContainsKey($x.file)) { $seen[$x.file] = $true; [void]$out.Add($x.file) } }
    return @($out)
}
function Get-ReimbDir { return (Join-Path $script:ShareDir '报销附件') }
function Get-ExpFilePath([string]$rel) { if (-not $rel) { return '' }; return (Join-Path $script:ShareDir ($rel -replace '[\\/]', [IO.Path]::DirectorySeparatorChar)) }
function Get-SafeName([string]$s) { return ($s -replace '[\\/:*?"<>|\r\n\t]', '').Trim() }

function New-BillNo([datetime]$d) {
    $pre = 'BX' + $d.ToString('yyyyMM') + '-'
    $max = 0
    foreach ($b in $script:Bills) { if ($b.no -like ($pre + '*')) { $n = 0; if ([int]::TryParse($b.no.Substring($pre.Length), [ref]$n) -and $n -gt $max) { $max = $n } } }
    return ($pre + ($max + 1).ToString('00'))
}
function New-Bill {
    $b = ConvertTo-Bill ([pscustomobject]@{ text = '新报销单'; no = (New-BillNo (Get-Date)); dept = (Get-ReimbDept); person = [string]$script:S.reimbName; state = '未提交'; created = (Get-Date).ToString($Fmt) })
    [void]$script:Bills.Add($b)
    Save-Bills
    return $b
}
function New-Expense([string]$billId, $h) {
    $o = [ordered]@{ bill = $billId; created = (Get-Date).ToString($Fmt) }
    foreach ($k in $h.Keys) { $o[$k] = $h[$k] }
    $x = ConvertTo-Expense ([pscustomobject]$o)
    if (-not $x.text.Trim()) { $x.text = Get-AutoTitle $x }
    [void]$script:Exps.Add($x)
    return $x
}

function Get-AutoTitle($x) {
    $route = if ($x.from -or $x.to) { ([string]$x.from + '→' + [string]$x.to) } else { '' }
    switch ($x.cat) {
        '城市间交通费' {
            $k = if ($x.invType -like '*铁路*') { '高铁/火车' } elseif ($x.invType -like '*机票*') { '飞机' } else { '交通' }
            return (($k + ' ' + $route).Trim())
        }
        '住宿费' { $t = '住宿'; if ($x.to) { $t += ' ' + $x.to }; if ($x.nights) { $t += ' ' + $x.nights + '晚' }; return $t }
        '市内交通费' { if ($route) { return ('打车 ' + $route) }; return '市内交通' }
        '租车/包车费' { if ($route) { return ('租车 ' + $route) }; return '租车/包车' }
        default { if ($x.seller) { return ($x.cat + ' ' + $x.seller) }; return $x.cat }
    }
}

# ---------- PDF/OFD 工具（第一次用到时才编译） ----------
$script:PdfRefs = $null
function Ensure-PdfTypes {
    if ('DtPdfDoc' -as [type]) { return }
    if (-not $script:PdfRefs) {
        Add-Type -AssemblyName System.IO.Compression, System.Xml, System.Core
        $script:PdfRefs = @([System.IO.Compression.ZipArchive].Assembly.Location, [System.Xml.XmlDocument].Assembly.Location, [System.Linq.Enumerable].Assembly.Location)
    }
    Add-Type -TypeDefinition $DtPdfSource -ReferencedAssemblies $script:PdfRefs
}
function Get-InvoiceLines([string]$path) {
    Ensure-PdfTypes
    $ext = [IO.Path]::GetExtension($path).ToLower()
    $bytes = [IO.File]::ReadAllBytes($path)
    if ($ext -eq '.ofd') { $items = [DtOfd]::Text($bytes) }
    elseif ($ext -eq '.pdf') { $items = ([DtPdfDoc]::new($bytes)).Text(2) }
    else { return @() }
    return @(([string][DtPdfDoc]::Lines($items)) -split "\r?\n" | Where-Object { $_.Trim() })
}

# ---------- 导入发票 ----------
function Get-ReimbName {
    $n = [string]$script:S.reimbName
    if ($n.Trim()) { return $n.Trim() }
    $owner = if ($script:CalWin) { $script:CalWin } else { $null }
    $r = Show-TextDialog '你的姓名' '' '发票会改成「姓名＋发票号码」这样的文件名（和单位 OA 里一样），也用作经办人。只问这一次，以后可以在「报销 → 设置」里改。'
    if ($r -and $r.Trim()) { $script:S.reimbName = $r.Trim(); Save-Settings; return $r.Trim() }
    return ''
}

function Update-BillFromExps($b) {
    $xs = @(Get-BillExps $b.id)
    $days = @($xs | ForEach-Object { $_.date; $_.date2 } | Where-Object { $_ -match '^\d{4}-\d{2}-\d{2}$' } | Sort-Object)
    if ($days.Count -gt 0) {
        if (-not $b.start -or $days[0] -lt $b.start) { $b.start = $days[0] }
        if (-not $b.end -or $days[$days.Count - 1] -gt $b.end) { $b.end = $days[$days.Count - 1] }
    }
    if (-not $b.place) {
        $first = @($xs | Where-Object { $_.cat -eq '城市间交通费' -and $_.to } | Select-Object -First 1)
        if ($first.Count -gt 0) { $b.place = $first[0].to }
    }
    if ($b.text -eq '新报销单' -and $b.place) { $b.text = $b.place + '出差' }
}

function Import-ReimbFiles([string[]]$paths) {
    $paths = @($paths | Where-Object { $_ -and (Test-Path -LiteralPath $_ -PathType Leaf) })
    if ($paths.Count -eq 0) { return }
    $b = Find-Bill $script:ReimbSel
    if (-not $b) { $b = New-Bill; $script:ReimbSel = $b.id }
    $name = if ($b.person) { $b.person } else { Get-ReimbName }
    $res = Import-ReimbCore $b $name $paths
    Render
    $msg = '导入 ' + $res.ok + ' 张，已改名放进「报销附件\' + $b.no + '」。'
    if ($res.noAmt -gt 0) { $msg += $res.noAmt.ToString() + ' 张没认出金额，标红了，点开补一下。' }
    if ($res.skip.Count -gt 0) { $msg += "`n重复的没导：" + ($res.skip -join '、') }
    if ($res.bad.Count -gt 0) { $msg += "`n没导进来：" + ($res.bad -join '、') }
    if ($res.skip.Count -gt 0 -or $res.bad.Count -gt 0) { [System.Windows.MessageBox]::Show($script:CalWin, $msg, '导入发票') | Out-Null } else { Notify '发票已导入' $msg }
}

function Import-ReimbCore($b, [string]$name, [string[]]$paths) {
    if (-not $b.person -and $name) { $b.person = $name }
    $dir = Join-Path (Get-ReimbDir) (Get-SafeName $b.no)
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $ok = 0; $skip = New-Object System.Collections.ArrayList; $bad = New-Object System.Collections.ArrayList; $noAmt = 0
    # 同一张发票 PDF 和 OFD 都下了的话，优先用 PDF（能拼进打印文件）
    $paths = @($paths | Sort-Object @{ Expression = { if ([IO.Path]::GetExtension($_).ToLower() -eq '.pdf') { 0 } else { 1 } } }, @{ Expression = { $_ } })
    foreach ($p in $paths) {
        $ext = [IO.Path]::GetExtension($p).ToLower()
        if (@('.pdf', '.ofd', '.jpg', '.jpeg', '.png', '.bmp') -notcontains $ext) { [void]$bad.Add([IO.Path]::GetFileName($p)); continue }
        $lines = @()
        try { $lines = @(Get-InvoiceLines $p) } catch { Write-Log ('读发票失败 ' + $p + '：' + $_.Exception.Message) }
        $r = Parse-InvoiceText ([string[]]$lines)
        if ($r.invNo) {
            $dup = @($script:Exps | Where-Object { $_.invNo -eq $r.invNo } | Select-Object -First 1)
            if ($dup.Count -gt 0) {
                $d0 = $dup[0]
                if ($ext -eq '.pdf' -and $d0.file -and [IO.Path]::GetExtension($d0.file).ToLower() -ne '.pdf') {
                    # 之前导的是 OFD，现在来了同一张的 PDF：换成 PDF，打印时能拼进去
                    $old = Get-ExpFilePath $d0.file
                    $np = [IO.Path]::ChangeExtension($old, '.pdf')
                    try {
                        $nd = Split-Path -Parent $np; if (-not (Test-Path -LiteralPath $nd)) { New-Item -ItemType Directory -Force -Path $nd | Out-Null }
                        Copy-Item -LiteralPath $p -Destination $np -Force
                        $nrel = [IO.Path]::ChangeExtension($d0.file, '.pdf')
                        foreach ($y in @($script:Exps | Where-Object { $_.file -eq $d0.file })) { $y.file = $nrel }
                        $ok++; continue
                    } catch {}
                }
                $db = Find-Bill $d0.bill; [void]$skip.Add([IO.Path]::GetFileName($p) + $(if ($db) { '（已在 ' + $db.no + '）' } else { '' })); continue
            }
        }
        $who = if ($name) { $name } else { $r.person }
        $base = Get-SafeName ($who + $(if ($r.invNo) { $r.invNo } else { [IO.Path]::GetFileNameWithoutExtension($p) }))
        $target = Join-Path $dir ($base + $ext)
        $k = 2
        while (Test-Path -LiteralPath $target) {
            if ((Get-Item -LiteralPath $target).Length -eq (Get-Item -LiteralPath $p).Length) { break }
            $target = Join-Path $dir ($base + '(' + $k + ')' + $ext); $k++
        }
        try { if (-not (Test-Path -LiteralPath $target)) { Copy-Item -LiteralPath $p -Destination $target } }
        catch { [void]$bad.Add([IO.Path]::GetFileName($p) + '（复制失败）'); continue }
        $rel = '报销附件\' + (Get-SafeName $b.no) + '\' + [IO.Path]::GetFileName($target)
        $h = @{
            cat = $r.cat; date = $r.date; time = $r.time; from = $r.from; to = $r.to; amount = $r.amount; person = $r.person
            invNo = $r.invNo; invDate = $r.invDate; invType = $r.invType; seller = $r.seller; file = $rel; nights = $r.nights
            status = '已取得'; seg = (@('城市间交通费', '租车/包车费') -contains $r.cat)
        }
        if (-not $h.cat) { $h.cat = '其他' }
        if ($r.trip) { $h.note = $r.trip }
        if ($r.remark) { $h.note = (([string]$h.note + ' ' + $r.remark).Trim()) }
        if ($r.date2) { $h.date2 = $r.date2 }
        if ($h.cat -eq '住宿费' -and $r.invDate -and -not $r.date) {
            $n0 = 1; [void][int]::TryParse([string]$r.nights, [ref]$n0); if ($n0 -lt 1) { $n0 = 1 }
            $h.date2 = $r.invDate
            $h.date = (Parse-Day $r.invDate).AddDays(-$n0).ToString('yyyy-MM-dd')
            $h.note = (([string]$h.note + ' 入住日期是按开票日期往前推的，核对一下').Trim())
        }
        if (-not $h.date -and $r.invDate) { $h.date = $r.invDate }
        if ($r.legs -gt 1) { $h.note = (([string]$h.note + ' 联程票，共 ' + $r.legs + ' 段，后面几段请自己补').Trim()) }
        $x = New-Expense $b.id $h
        if (-not $x.amount) { $noAmt++ }
        $ok++
    }
    Update-BillFromExps $b
    Save-Bills; Save-Exps
    return @{ ok = $ok; skip = $skip; bad = $bad; noAmt = $noAmt }
}

function Pick-ReimbFiles {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = '发票（PDF、OFD、图片）|*.pdf;*.ofd;*.jpg;*.jpeg;*.png;*.bmp|所有文件|*.*'
    $dlg.Multiselect = $true
    $dl = Join-Path $env:USERPROFILE 'Downloads'
    if (Test-Path -LiteralPath $dl) { $dlg.InitialDirectory = $dl }
    if ($dlg.ShowDialog($script:CalWin)) { Import-ReimbFiles $dlg.FileNames }
}

# ---------- 记一笔：「26日7点 XX酒店至会展中心 27.6」 ----------
function Parse-QuickExpense([string]$s, $b) {
    $h = @{ status = '待补' }
    $t = ' ' + $s.Trim() + ' '
    if ($t -match '([\d]+(?:\.\d{1,2})?)\s*元?\s*$') { $h.amount = ([double]$Matches[1]).ToString('0.00', $Inv); $t = $t.Substring(0, $t.Length - $Matches[0].Length) + ' ' }
    $base = if ($b -and (Parse-Day $b.start)) { Parse-Day $b.start } else { (Get-Date).Date }
    if ($t -match '(\d{4})[年\-/.](\d{1,2})[月\-/.](\d{1,2})日?') { try { $h.date = ([datetime]::new([int]$Matches[1], [int]$Matches[2], [int]$Matches[3])).ToString('yyyy-MM-dd') } catch {}; $t = $t.Replace($Matches[0], ' ') }
    elseif ($t -match '(\d{1,2})月(\d{1,2})[日号]?') { try { $h.date = ([datetime]::new($base.Year, [int]$Matches[1], [int]$Matches[2])).ToString('yyyy-MM-dd') } catch {}; $t = $t.Replace($Matches[0], ' ') }
    elseif ($t -match '(?<![\d.])(\d{1,2})/(\d{1,2})(?![\d.])') { try { $h.date = ([datetime]::new($base.Year, [int]$Matches[1], [int]$Matches[2])).ToString('yyyy-MM-dd') } catch {}; $t = $t.Replace($Matches[0], ' ') }
    elseif ($t -match '(\d{1,2})[日号]') {
        $dd = [int]$Matches[1]
        try {
            $d = [datetime]::new($base.Year, $base.Month, $dd)
            if ($d -lt $base) { $d = $d.AddMonths(1) }
            $h.date = $d.ToString('yyyy-MM-dd')
        } catch {}
        $t = $t.Replace($Matches[0], ' ')
    }
    elseif ($t -match '今天') { $h.date = (Get-Date).ToString('yyyy-MM-dd'); $t = $t.Replace('今天', ' ') }
    elseif ($t -match '昨天') { $h.date = (Get-Date).AddDays(-1).ToString('yyyy-MM-dd'); $t = $t.Replace('昨天', ' ') }
    if ($t -match '(\d{1,2})\s*[:：]\s*(\d{2})') { $h.time = ('{0:D2}:{1}' -f [int]$Matches[1], $Matches[2]); $t = $t.Replace($Matches[0], ' ') }
    elseif ($t -match '(\d{1,2})\s*[点时](半|(\d{1,2})分?)?') {
        $mm = if ($Matches[2] -eq '半') { 30 } elseif ($Matches[3]) { [int]$Matches[3] } else { 0 }
        $hh = [int]$Matches[1]
        if ($hh -le 24) { $h.time = ('{0:D2}:{1:D2}' -f $hh, $mm) }
        $t = $t.Replace($Matches[0], ' ')
    }
    if ($t -match '(\d{1,2})\s*(晚|天)') { $h.nights = $Matches[1]; $t = $t.Replace($Matches[0], ' ') }
    $t = ($t -replace '\s+', ' ').Trim()
    $parts = @([regex]::Split($t, '\s*(?:至|到|→|->|—|－|-)\s*') | Where-Object { $_ })
    if ($parts.Count -ge 2) {
        $f0 = @($parts[0].Trim() -split '\s+'); $t0 = @($parts[1].Trim() -split '\s+')
        $h.from = $f0[$f0.Count - 1]; $h.to = $t0[0]
    }
    if ($t -match '高铁|动车|火车|飞机|航班|机票') { $h.cat = '城市间交通费'; $h.seg = $true }
    elseif ($t -match '租车|包车') { $h.cat = '租车/包车费'; $h.seg = $true }
    elseif ($t -match '补助|补贴') { $h.cat = '出差补助'; $h.status = '无票' }
    elseif ($t -match '早餐|午餐|晚餐|餐|饭') { $h.cat = '餐费' }
    elseif ($parts.Count -lt 2 -and $t -match '住宿|酒店|宾馆') { $h.cat = '住宿费' }
    elseif ($t -match '会议|会务|注册费') { $h.cat = '会议/会务费' }
    elseif ($t -match '打印|复印|印刷|资料') { $h.cat = '资料/印刷费' }
    elseif ($parts.Count -ge 2 -or $t -match '打车|出租|地铁|公交|网约车|滴滴') { $h.cat = '市内交通费' }
    else { $h.cat = '其他' }
    if ($parts.Count -lt 2 -and $t) { $h.text = $t }
    if ($h.cat -eq '住宿费' -and $parts.Count -lt 2) { $hn = ($t -replace '住宿费?|入住', '').Trim(); if ($hn) { $h.to = $hn; $h.Remove('text') } }
    if (-not $h.date -and $b -and $b.start) { $h.date = $b.start }
    if (-not $h.date) { $h.date = (Get-Date).ToString('yyyy-MM-dd') }
    return $h
}

function Add-QuickExpense([string]$s) {
    if (-not $s.Trim()) { return }
    $b = Find-Bill $script:ReimbSel
    if (-not $b) { $b = New-Bill; $script:ReimbSel = $b.id }
    $h = Parse-QuickExpense $s $b
    $x = New-Expense $b.id $h
    Update-BillFromExps $b
    Save-Bills; Save-Exps
    $script:ReimbEdit = $x.id
    Render
}

# 一张发票拆成几笔（比如一张网约车发票里有好几趟）
function Split-Expense([string]$id) {
    $x = Find-Exp $id; if (-not $x) { return }
    $b = Find-Bill $x.bill
    $note = '一行写一笔，格式和「记一笔」一样，例如：' + "`n26日7点 XX酒店至会展中心 27.6`n" + '拆完的几笔共用这张发票。原来这一笔是 ' + (Format-Amt (Get-Amt $x)) + ' 元。'
    $r = Show-TextDialog '拆成几笔' '' $note $true
    if (-not $r -or -not $r.Trim()) { return }
    $rows = @($r -split "\r?\n" | Where-Object { $_.Trim() })
    $sum = 0.0; $hs = @()
    foreach ($l in $rows) { $h = Parse-QuickExpense $l $b; $sum += Get-Amt ([string]$h.amount); $hs += , $h }
    if ([math]::Abs($sum - (Get-Amt $x)) -gt 0.005 -and (Get-Amt $x) -gt 0) {
        $q = [System.Windows.MessageBox]::Show($script:CalWin, ('拆出来的合计 {0} 元，和发票上的 {1} 元对不上。还是要拆吗？' -f (Format-Amt $sum), (Format-Amt (Get-Amt $x))), '拆成几笔', 'YesNo', 'Question')
        if ($q -ne 'Yes') { return }
    }
    foreach ($h in $hs) {
        foreach ($k in 'invNo', 'invDate', 'invType', 'seller', 'file', 'person') { $h[$k] = $x.$k }
        if (-not $h.cat -or $h.cat -eq '其他') { $h.cat = $x.cat }
        $h.status = $x.status; $h.grp = $x.id
        [void](New-Expense $x.bill $h)
    }
    $script:Exps.Remove($x)
    $script:ReimbEdit = ''
    Save-Exps; Render
}

function Remove-Expense([string]$id) {
    $x = Find-Exp $id; if (-not $x) { return }
    $q = [System.Windows.MessageBox]::Show($script:CalWin, ('删掉「' + $x.text + '」这一笔？发票文件留在文件夹里不删。'), '删除', 'YesNo', 'Question')
    if ($q -ne 'Yes') { return }
    $script:Exps.Remove($x)
    if ($script:ReimbEdit -eq $id) { $script:ReimbEdit = '' }
    Save-Exps; Render
}

function Remove-Bill([string]$id) {
    $b = Find-Bill $id; if (-not $b) { return }
    $xs = @(Get-BillExps $id)
    $q = [System.Windows.MessageBox]::Show($script:CalWin, ('删除报销单 ' + $b.no + '「' + $b.text + '」和里面的 ' + $xs.Count + ' 笔费用？发票文件留在文件夹里不删。'), '删除报销单', 'YesNo', 'Warning')
    if ($q -ne 'Yes') { return }
    foreach ($x in $xs) { $script:Exps.Remove($x) }
    $script:Bills.Remove($b)
    $script:ReimbSel = ''; $script:ReimbEdit = ''
    Save-Bills; Save-Exps; Render
}

function Open-ExpFile([string]$rel) {
    $p = Get-ExpFilePath $rel
    if ($p -and (Test-Path -LiteralPath $p)) { Start-Process -FilePath $p }
    else { Notify '找不到发票文件' ('可能另一台电脑还没同步过来：' + $rel) }
}

function Set-ExpFile([string]$id) {
    $x = Find-Exp $id; if (-not $x) { return }
    $b = Find-Bill $x.bill
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = '发票（PDF、OFD、图片）|*.pdf;*.ofd;*.jpg;*.jpeg;*.png;*.bmp|所有文件|*.*'
    if (-not $dlg.ShowDialog($script:CalWin)) { return }
    $p = $dlg.FileName
    $ext = [IO.Path]::GetExtension($p).ToLower()
    $r = $null
    try { $r = Parse-InvoiceText ([string[]]@(Get-InvoiceLines $p)) } catch {}
    $who = if ($b -and $b.person) { $b.person } else { Get-ReimbName }
    $no = if ($r -and $r.invNo) { $r.invNo } else { [IO.Path]::GetFileNameWithoutExtension($p) }
    $dir = Join-Path (Get-ReimbDir) (Get-SafeName $(if ($b) { $b.no } else { '未归单' }))
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $target = Join-Path $dir ((Get-SafeName ($who + $no)) + $ext)
    if (-not (Test-Path -LiteralPath $target)) { Copy-Item -LiteralPath $p -Destination $target }
    $x.file = '报销附件\' + (Split-Path -Leaf $dir) + '\' + [IO.Path]::GetFileName($target)
    $x.status = '已取得'
    if ($r) {
        if ($r.invNo) { $x.invNo = $r.invNo }
        if ($r.invDate) { $x.invDate = $r.invDate }
        if ($r.invType) { $x.invType = $r.invType }
        if ($r.seller -and -not $x.seller) { $x.seller = $r.seller }
        if ($r.amount -and -not $x.amount) { $x.amount = $r.amount }
    }
    Save-Exps; Render
}

function Move-ExpToBill([string]$id, [string]$billId) {
    $x = Find-Exp $id; $b = Find-Bill $billId
    if (-not $x -or -not $b -or $x.bill -eq $billId) { return }
    # 同一张发票拆出来的几笔一起挪
    $group = @($script:Exps | Where-Object { $_.id -eq $id -or ($x.file -and $_.file -eq $x.file -and $_.bill -eq $x.bill) })
    $oldDir = Join-Path (Get-ReimbDir) (Get-SafeName ((Find-Bill $x.bill).no))
    $newDir = Join-Path (Get-ReimbDir) (Get-SafeName $b.no)
    if ($x.file) {
        $src = Get-ExpFilePath $x.file
        if (Test-Path -LiteralPath $src) {
            if (-not (Test-Path -LiteralPath $newDir)) { New-Item -ItemType Directory -Force -Path $newDir | Out-Null }
            $dst = Join-Path $newDir (Split-Path -Leaf $src)
            try { if (-not (Test-Path -LiteralPath $dst)) { Move-Item -LiteralPath $src -Destination $dst } } catch { try { Copy-Item -LiteralPath $src -Destination $dst } catch {} }
            $rel = '报销附件\' + (Split-Path -Leaf $newDir) + '\' + (Split-Path -Leaf $src)
            foreach ($g in $group) { $g.file = $rel }
        }
    }
    foreach ($g in $group) { $g.bill = $billId }
    Update-BillFromExps $b
    Save-Bills; Save-Exps; Render
}

# ---------- OA 填报要的东西 ----------
function Format-DayHour([string]$date, [string]$time) {
    $s = ''
    if ($date -match '^\d{4}-(\d{2})-(\d{2})$') { $s = ([int]$Matches[1]).ToString() + '月' + ([int]$Matches[2]).ToString() + '日' }
    if ($time -match '^(\d{1,2}):(\d{2})') { $s += ([int]$Matches[1]).ToString() + '时'; if ([int]$Matches[2] -ne 0) { $s += ([int]$Matches[2]).ToString() + '分' } }
    return $s
}
function Format-CityRide($x) {
    $s = ''
    if ($x.date -match '^\d{4}-\d{2}-(\d{2})$') { $s += ([int]$Matches[1]).ToString() + '日' }
    if ($x.time -match '^(\d{1,2}):(\d{2})') { $s += ([int]$Matches[1]).ToString() + '点'; if ([int]$Matches[2] -ne 0) { $s += ([int]$Matches[2]).ToString() + '分' } }
    if ($x.from -or $x.to) { $s += [string]$x.from + '至' + [string]$x.to } else { $s += $x.text }
    return ($s + (Format-AmtRaw (Get-Amt $x)) + '元')
}

function Get-OAData($b) {
    $xs = @(Get-BillExps $b.id)
    $segs = New-Object System.Collections.ArrayList
    $map = @{}
    foreach ($x in @($xs | Where-Object { $_.seg })) {
        $k = ([string]$x.date + '|' + [string]$x.time + '|' + [string]$x.from + '|' + [string]$x.to)
        if ($map.ContainsKey($k)) { $sg = $map[$k]; $sg.traffic += Get-Amt $x; $sg.n++; if (-not $sg.date2 -and $x.date2) { $sg.date2 = $x.date2; $sg.time2 = $x.time2 }; [void]$sg.ids.Add($x.id); continue }
        $sg = @{ date = [string]$x.date; time = [string]$x.time; date2 = [string]$x.date2; time2 = [string]$x.time2; from = [string]$x.from; to = [string]$x.to; traffic = (Get-Amt $x); hotel = 0.0; n = 1; ids = (New-Object System.Collections.ArrayList); kind = $x.cat }
        [void]$sg.ids.Add($x.id)
        $map[$k] = $sg; [void]$segs.Add($sg)
    }
    $segs = @($segs | Sort-Object @{ Expression = { $_.date } }, @{ Expression = { if ($_.time) { $_.time } else { '99' } } })
    $hotelLeft = 0.0
    foreach ($h in @($xs | Where-Object { $_.cat -eq '住宿费' })) {
        $target = $null
        foreach ($sg in $segs) { if ($sg.date -and $h.date -and $sg.date -le $h.date) { $target = $sg } }
        if (-not $target -and $segs.Count -gt 0) { $target = $segs[0] }
        if ($target) { $target.hotel += Get-Amt $h } else { $hotelLeft += Get-Amt $h }
    }
    $city = @($xs | Where-Object { $_.cat -eq '市内交通费' -and -not $_.seg })
    $cityTotal = 0.0; foreach ($x in $city) { $cityTotal += Get-Amt $x }
    $cityDesc = if ($city.Count -gt 0) { '市内交通费' + (Format-AmtShort $cityTotal) + '元（' + (($city | ForEach-Object { Format-CityRide $_ }) -join '、') + '）' } else { '' }
    $others = @($xs | Where-Object { -not $_.seg -and $_.cat -ne '住宿费' -and $_.cat -ne '市内交通费' })
    $otherByCat = [ordered]@{}
    foreach ($x in $others) { if (-not $otherByCat.Contains($x.cat)) { $otherByCat[$x.cat] = 0.0 }; $otherByCat[$x.cat] += Get-Amt $x }
    $fixed = ([string]$script:S.reimbNote).Trim()
    $memo = (@($fixed, $cityDesc) | Where-Object { $_ }) -join '；'
    $traffic = 0.0; $hotel = 0.0; foreach ($sg in $segs) { $traffic += $sg.traffic; $hotel += $sg.hotel }
    return @{ segs = $segs; hotelLeft = $hotelLeft; city = $city; cityTotal = $cityTotal; cityDesc = $cityDesc; others = $others; otherByCat = $otherByCat; memo = $memo; traffic = $traffic; hotel = $hotel + $hotelLeft; total = (Get-BillTotal $b.id) }
}

# ---------- 导出 Excel（自己拼一个最简单的 xlsx，不用装 Office） ----------
function Get-ColName([int]$i) { $s = ''; $i++; while ($i -gt 0) { $m = ($i - 1) % 26; $s = ([string][char](65 + $m)) + $s; $i = [int](($i - $m - 1) / 26) }; return $s }
# sheets：@{ name; rows = 二维数组；head = 表头所在行号（从 0 数）；widths；title = 标题行号 }
function Write-SimpleXlsx([string]$path, $sheets) {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $path) { Remove-Item -LiteralPath $path -Force }
    $fs = [IO.File]::Open($path, 'Create')
    $zip = [System.IO.Compression.ZipArchive]::new($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    $put = { param($n, $t) $e = $zip.CreateEntry($n); $w = [IO.StreamWriter]::new($e.Open(), $Utf8); $w.Write($t); $w.Dispose() }
    try {
        $ov = ''; $wbS = ''; $rel = ''
        for ($i = 0; $i -lt $sheets.Count; $i++) {
            $sh = $sheets[$i]; $n = $i + 1
            $ov += '<Override PartName="/xl/worksheets/sheet' + $n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            $wbS += '<sheet name="' + (Get-XmlText $sh.name) + '" sheetId="' + $n + '" r:id="rId' + $n + '"/>'
            $rel += '<Relationship Id="rId' + $n + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + $n + '.xml"/>'
            $sb = [Text.StringBuilder]::new()
            [void]$sb.Append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">')
            if ($sh.head -ne $null) { [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"><pane ySplit="' + ([int]$sh.head + 1) + '" topLeftCell="A' + ([int]$sh.head + 2) + '" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>') }
            if ($sh.widths) { [void]$sb.Append('<cols>'); for ($c = 0; $c -lt $sh.widths.Count; $c++) { [void]$sb.Append('<col min="' + ($c + 1) + '" max="' + ($c + 1) + '" width="' + $sh.widths[$c] + '" customWidth="1"/>') }; [void]$sb.Append('</cols>') }
            [void]$sb.Append('<sheetData>')
            $rows = @($sh.rows)
            for ($r = 0; $r -lt $rows.Count; $r++) {
                [void]$sb.Append('<row r="' + ($r + 1) + '">')
                $cells = @($rows[$r])
                for ($c = 0; $c -lt $cells.Count; $c++) {
                    $v = $cells[$c]; if ($null -eq $v -or ($v -is [string] -and $v -eq '')) { continue }
                    $ref = (Get-ColName $c) + ($r + 1)
                    $st = if ($sh.title -ne $null -and $r -eq [int]$sh.title) { 3 } elseif (@($sh.heads) -contains $r -or ($sh.head -ne $null -and $r -eq [int]$sh.head)) { 1 } elseif ($v -is [double] -or $v -is [int]) { 2 } else { 4 }
                    if ($v -is [double] -or $v -is [int]) { [void]$sb.Append('<c r="' + $ref + '" s="' + $(if ($st -eq 1) { 5 } else { 2 }) + '"><v>' + ([double]$v).ToString('R', $Inv) + '</v></c>') }
                    else { [void]$sb.Append('<c r="' + $ref + '" s="' + $st + '" t="inlineStr"><is><t xml:space="preserve">' + (Get-XmlText ([string]$v)) + '</t></is></c>') }
                }
                [void]$sb.Append('</row>')
            }
            [void]$sb.Append('</sheetData>')
            if ($sh.merges) { [void]$sb.Append('<mergeCells count="' + @($sh.merges).Count + '">'); foreach ($m in $sh.merges) { [void]$sb.Append('<mergeCell ref="' + $m + '"/>') }; [void]$sb.Append('</mergeCells>') }
            [void]$sb.Append('<pageMargins left="0.5" right="0.5" top="0.6" bottom="0.6" header="0.3" footer="0.3"/><pageSetup paperSize="9" orientation="landscape" fitToHeight="0"/></worksheet>')
            & $put ('xl/worksheets/sheet' + $n + '.xml') $sb.ToString()
        }
        & $put '[Content_Types].xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + $ov + '</Types>')
        & $put '_rels/.rels' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
        & $put 'xl/workbook.xml' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + $wbS + '</sheets></workbook>')
        & $put 'xl/_rels/workbook.xml.rels' ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + $rel + '<Relationship Id="rId' + ($sheets.Count + 1) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        & $put 'xl/styles.xml' '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="164" formatCode="#,##0.00"/></numFmts><fonts count="3"><font><sz val="11"/><name val="等线"/><family val="2"/><charset val="134"/></font><font><b/><sz val="11"/><name val="等线"/><family val="2"/><charset val="134"/></font><font><b/><sz val="14"/><name val="等线"/><family val="2"/><charset val="134"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFEDEFF2"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border><border><left style="thin"><color rgb="FFBFC4CC"/></left><right style="thin"><color rgb="FFBFC4CC"/></right><top style="thin"><color rgb="FFBFC4CC"/></top><bottom style="thin"><color rgb="FFBFC4CC"/></bottom><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/><xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/><xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"><alignment vertical="top" wrapText="1"/></xf><xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    } finally { $zip.Dispose(); $fs.Dispose() }
}

function Get-ExpRow($x, [int]$no, [bool]$withBill) {
    $row = @()
    if ($withBill) { $b = Find-Bill $x.bill; $row += $(if ($b) { $b.no } else { '' }) }
    $row += @($no, $x.date, $x.time, $x.cat, $x.text, $x.from, $x.to, (($x.date2 + ' ' + $x.time2).Trim()), $x.person, (Get-Amt $x), $x.invNo, $x.invDate, $x.invType, $x.seller, $x.status, $(if ($x.file) { Split-Path -Leaf $x.file } else { '' }), $x.note)
    return , $row
}
$ExpHead = @('序号', '日期', '时间', '费用类别', '事项/行程', '出发地', '目的地', '到达时间', '人员', '金额', '发票号码', '开票日期', '票据类型', '开票方', '票据状态', '发票文件', '备注')
$ExpWidths = @(6, 11, 7, 12, 30, 16, 16, 15, 8, 11, 23, 11, 14, 22, 9, 30, 24)

function Get-BillSheets($b) {
    $oa = Get-OAData $b
    $xs = @(Get-BillExps $b.id)
    $rows = New-Object System.Collections.ArrayList
    [void]$rows.Add(@('OA 报销填报  ' + $b.no))
    [void]$rows.Add(@('报销事由', $b.text)); [void]$rows.Add(@('出差地点', $b.place))
    [void]$rows.Add(@('出差日期', (($b.start + ' 至 ' + $b.end).Trim(' 至'))))
    [void]$rows.Add(@('成本中心', $b.dept)); [void]$rows.Add(@('项目', $(if ($b.proj) { Get-ProjName $b.proj } else { '' })))
    [void]$rows.Add(@('报销人', $b.person)); [void]$rows.Add(@('出差人员', $b.people)); [void]$rows.Add(@('OA 单号', $b.oaNo))
    [void]$rows.Add(@())
    $h1 = $rows.Count
    [void]$rows.Add(@('段', '出发日期', '出发时', '出发分', '出发地', '到达日期', '到达时', '到达分', '目的地', '交通费', '住宿费'))
    $i = 0
    foreach ($sg in $oa.segs) {
        $i++
        $h0 = ''; $m0 = ''; if ($sg.time -match '^(\d{1,2}):(\d{2})') { $h0 = [double][int]$Matches[1]; if ([int]$Matches[2]) { $m0 = [double][int]$Matches[2] } }
        $h2 = ''; $m2 = ''; if ($sg.time2 -match '^(\d{1,2}):(\d{2})') { $h2 = [double][int]$Matches[1]; if ([int]$Matches[2]) { $m2 = [double][int]$Matches[2] } }
        [void]$rows.Add(@([double]$i, $sg.date, $h0, $m0, $sg.from, $sg.date2, $h2, $m2, $sg.to, [double]$sg.traffic, [double]$sg.hotel))
    }
    if ($oa.hotelLeft -gt 0) { [void]$rows.Add(@('', '', '', '', '', '', '', '', '（没有对应行程的住宿）', '', [double]$oa.hotelLeft)) }
    [void]$rows.Add(@('合计', '', '', '', '', '', '', '', '', [double]$oa.traffic, [double]$oa.hotel))
    [void]$rows.Add(@())
    $h2r = $rows.Count
    [void]$rows.Add(@('费用', '金额', '说明'))
    [void]$rows.Add(@('市内交通费', [double]$oa.cityTotal, $oa.cityDesc))
    foreach ($k in $oa.otherByCat.Keys) { [void]$rows.Add(@($k, [double]$oa.otherByCat[$k], ((@($oa.others | Where-Object { $_.cat -eq $k } | ForEach-Object { $_.text + ' ' + (Format-AmtRaw (Get-Amt $_)) }) -join '；')))) }
    [void]$rows.Add(@('总计', [double]$oa.total, ('单据 ' + @(Get-BillFiles $b.id).Count + ' 张')))
    [void]$rows.Add(@())
    [void]$rows.Add(@('其他事项说明', $oa.memo))
    [void]$rows.Add(@('情况说明', $b.note))
    $s1 = @{ name = 'OA填报'; rows = $rows; title = 0; heads = @($h1, $h2r); widths = @(12, 12, 8, 8, 18, 12, 8, 8, 22, 12, 12) }
    $d = New-Object System.Collections.ArrayList
    [void]$d.Add($ExpHead)
    $n = 0; foreach ($x in $xs) { $n++; [void]$d.Add((Get-ExpRow $x $n $false)) }
    $s2 = @{ name = '费用明细'; rows = $d; head = 0; widths = $ExpWidths }
    $sum = New-Object System.Collections.ArrayList
    [void]$sum.Add(@('费用类别', '笔数', '金额', '待补票据'))
    foreach ($c in $ReimbCats) {
        $g = @($xs | Where-Object { $_.cat -eq $c }); if ($g.Count -eq 0) { continue }
        $t = 0.0; foreach ($x in $g) { $t += Get-Amt $x }
        [void]$sum.Add(@($c, [double]$g.Count, [double]$t, [double]@($g | Where-Object { $_.status -eq '待补' }).Count))
    }
    [void]$sum.Add(@('合计', [double]$xs.Count, [double]$oa.total, ''))
    $s3 = @{ name = '费用汇总'; rows = $sum; head = 0; widths = @(16, 8, 14, 10) }
    return @($s1, $s2, $s3)
}

function Export-BillXlsx([string]$id) {
    $b = Find-Bill $id; if (-not $b) { return }
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Excel 工作簿|*.xlsx'; $dlg.FileName = (Get-SafeName ($b.no + '_' + $b.text)) + '.xlsx'
    if (-not $dlg.ShowDialog($script:CalWin)) { return }
    try { Write-SimpleXlsx $dlg.FileName (Get-BillSheets $b); Start-Process $dlg.FileName }
    catch { [System.Windows.MessageBox]::Show($script:CalWin, '导出失败：' + $_.Exception.Message + "`n如果这个文件正在 Excel 里开着，先关掉再导。", '桌面待办') | Out-Null }
}

function Get-LedgerSheets {
    $bs = @($script:Bills | Sort-Object @{ Expression = { $_.start } ; Descending = $true }, @{ Expression = { $_.no }; Descending = $true })
    $r1 = New-Object System.Collections.ArrayList
    [void]$r1.Add(@('单号', '事由', '项目', '出差地点', '开始', '结束', '报销人', '出差人员', '成本中心', '金额', '单据张数', '待补票据', '状态', '提交日期', 'OA 单号', '情况说明'))
    foreach ($b in $bs) {
        $xs = @(Get-BillExps $b.id)
        [void]$r1.Add(@($b.no, $b.text, $(if ($b.proj) { Get-ProjName $b.proj } else { '' }), $b.place, $b.start, $b.end, $b.person, $b.people, $b.dept, [double](Get-BillTotal $b.id), [double]@(Get-BillFiles $b.id).Count, [double]@($xs | Where-Object { $_.status -eq '待补' }).Count, $b.state, $b.submit, $b.oaNo, $b.note))
    }
    $r2 = New-Object System.Collections.ArrayList
    [void]$r2.Add(@('单号') + $ExpHead)
    foreach ($b in $bs) { $n = 0; foreach ($x in (Get-BillExps $b.id)) { $n++; [void]$r2.Add((Get-ExpRow $x $n $true)) } }
    $r3 = New-Object System.Collections.ArrayList
    [void]$r3.Add(@('项目') + $ReimbCats + @('合计', '报销单数'))
    $byProj = [ordered]@{}
    foreach ($b in $bs) { $k = if ($b.proj) { Get-ProjName $b.proj } else { '（没选项目）' }; if (-not $byProj.Contains($k)) { $byProj[$k] = New-Object System.Collections.ArrayList }; [void]$byProj[$k].Add($b) }
    foreach ($k in $byProj.Keys) {
        $ids = @{}; foreach ($b in $byProj[$k]) { $ids[$b.id] = $true }
        $row = @($k); $tot = 0.0
        foreach ($c in $ReimbCats) { $t = 0.0; foreach ($x in $script:Exps) { if ($ids.ContainsKey($x.bill) -and $x.cat -eq $c) { $t += Get-Amt $x } }; $row += [double]$t; $tot += $t }
        $row += [double]$tot; $row += [double]$byProj[$k].Count
        [void]$r3.Add($row)
    }
    return @(
        @{ name = '报销台账'; rows = $r1; head = 0; widths = @(13, 30, 14, 16, 11, 11, 8, 14, 18, 12, 9, 9, 8, 11, 14, 30) },
        @{ name = '全部明细'; rows = $r2; head = 0; widths = @(13) + $ExpWidths },
        @{ name = '按项目汇总'; rows = $r3; head = 0; widths = @(16) + @($ReimbCats | ForEach-Object { 12 }) + @(12, 9) }
    )
}
function Export-LedgerXlsx {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = 'Excel 工作簿|*.xlsx'; $dlg.FileName = '报销台账_' + (Get-Date).ToString('yyyyMMdd') + '.xlsx'
    if (-not $dlg.ShowDialog($script:CalWin)) { return }
    try { Write-SimpleXlsx $dlg.FileName (Get-LedgerSheets); Start-Process $dlg.FileName }
    catch { [System.Windows.MessageBox]::Show($script:CalWin, '导出失败：' + $_.Exception.Message, '桌面待办') | Out-Null }
}

# ---------- 打印：粘贴单 + 每张发票一页 A4 ----------
function New-PdfText([double]$x, [double]$y, [double]$size, [string]$text) { $t = [DtPdfItem]::new(); $t.X = $x; $t.Y = $y; $t.Size = $size; $t.Text = $text; return $t }

function Get-PasteSheetTexts($b, [int]$count) {
    $l = [System.Collections.Generic.List[DtPdfItem]]::new()
    $amt = Get-BillTotal $b.id
    $l.Add((New-PdfText 500 512 13 (Format-Amt $amt)))
    $cs = [string]$count; $l.Add((New-PdfText (706 - $cs.Length * 3.3) 512 13 $cs))
    $l.Add((New-PdfText 486 481 13 ([string]$b.dept)))
    if ($b.oaNo) { $l.Add((New-PdfText 668 481 12 ([string]$b.oaNo))) }
    $l.Add((New-PdfText 688 450 13 ([string]$b.person)))
    return , $l
}

# 没有模板文件时，自己画一张一样版式的粘贴单
function Add-BlankPasteSheet($out, $texts) {
    $t = [System.Collections.Generic.List[DtPdfItem]]::new()
    $ln = [System.Collections.Generic.List[double[]]]::new()
    $t.Add((New-PdfText 300 546 26 '报 销 单 据 粘 贴 单')); $ln.Add([double[]]@(298, 541, 562, 541))
    $t.Add((New-PdfText 434 512 14 '金  额：¥')); $ln.Add([double[]]@(497, 509, 606, 509))
    $t.Add((New-PdfText 621 512 14 '单据：')); $ln.Add([double[]]@(668, 509, 745, 509)); $t.Add((New-PdfText 746 512 14 '张'))
    $t.Add((New-PdfText 434 481 14 '部  门：')); $ln.Add([double[]]@(484, 478, 606, 478))
    $t.Add((New-PdfText 621 481 14 '编号：')); $ln.Add([double[]]@(666, 478, 760, 478))
    $t.Add((New-PdfText 621 450 14 '经办人：')); $ln.Add([double[]]@(684, 447, 760, 447))
    foreach ($y0 in 120, 260, 400) { $ln.Add([double[]]@(90, $y0, 90, ($y0 + 90))) }
    $t.Add((New-PdfText 85 345 12 '装')); $t.Add((New-PdfText 85 225 12 '订')); $t.Add((New-PdfText 85 105 12 '线'))
    foreach ($x in $texts) { $t.Add($x) }
    $out.AddBlankSheet($t, $ln)
}

# 粘贴单模板：程序文件夹里的「报销单据粘贴单.pdf」（单位换了样式就替换它）；没有的话用程序里带的那份，并放一份到程序文件夹
function Get-PasteTemplate {
    $tpl = Join-Path $AppDir '报销单据粘贴单.pdf'
    if (Test-Path -LiteralPath $tpl) { try { return , ([IO.File]::ReadAllBytes($tpl)) } catch {} }
    if (-not $PasteSheetB64) { return $null }
    $bytes = [Convert]::FromBase64String($PasteSheetB64)
    try { [IO.File]::WriteAllBytes($tpl, $bytes) } catch {}
    return , $bytes
}

function Convert-ImageToJpeg([string]$path) {
    $img = [System.Drawing.Image]::FromFile($path)
    try {
        $bmp = [System.Drawing.Bitmap]::new($img.Width, $img.Height, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $g = [System.Drawing.Graphics]::FromImage($bmp)
        $g.Clear([System.Drawing.Color]::White); $g.DrawImage($img, 0, 0, $img.Width, $img.Height); $g.Dispose()
        $ms = [IO.MemoryStream]::new()
        $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Jpeg)
        $r = @($ms.ToArray(), $bmp.Width, $bmp.Height)
        $bmp.Dispose()
        return , $r
    } finally { $img.Dispose() }
}

function Print-Bill([string]$id, [string]$what = 'all') {
    $b = Find-Bill $id; if (-not $b) { return }
    try { Ensure-PdfTypes } catch { [System.Windows.MessageBox]::Show($script:CalWin, '打印组件加载失败：' + $_.Exception.Message, '桌面待办') | Out-Null; return }
    $files = @(Get-BillFiles $b.id)
    $out = [DtPdfOut]::new()
    $miss = New-Object System.Collections.ArrayList
    if ($what -ne 'inv') {
        $texts = Get-PasteSheetTexts $b $files.Count
        $done = $false
        $tb0 = Get-PasteTemplate
        if ($tb0) { try { $out.AddTemplate($tb0, $texts); $done = $true } catch { Write-Log ('粘贴单模板读取失败：' + $_.Exception.Message) } }
        if (-not $done) { Add-BlankPasteSheet $out $texts }
    }
    if ($what -ne 'sheet') {
        $i = 0
        foreach ($rel in $files) {
            $i++
            $p = Get-ExpFilePath $rel
            $cap = if ($script:S.reimbCaption) { '第 ' + $i + ' 张 / 共 ' + $files.Count + ' 张　' + (Split-Path -Leaf $rel) } else { '' }
            if (-not (Test-Path -LiteralPath $p)) { [void]$miss.Add((Split-Path -Leaf $rel) + '（文件还没同步过来）'); continue }
            $ext = [IO.Path]::GetExtension($p).ToLower()
            try {
                if ($ext -eq '.pdf') { $out.AddInvoice([IO.File]::ReadAllBytes($p), $cap) }
                elseif (@('.jpg', '.jpeg', '.png', '.bmp') -contains $ext) { $j = Convert-ImageToJpeg $p; $out.AddJpeg([byte[]]$j[0], [int]$j[1], [int]$j[2], $cap) }
                else { [void]$miss.Add((Split-Path -Leaf $rel) + '（OFD 要用发票软件单独打印）') }
            } catch { Write-Log ('打印拼页失败 ' + $p + '：' + $_.Exception.Message); [void]$miss.Add((Split-Path -Leaf $rel) + '（这个文件拼不进去，单独打开打印）') }
        }
    }
    if ($out.PageCount -eq 0) { Notify '没有可打印的' '这张报销单还没有发票文件。'; return }
    $dir = Join-Path $DataDir 'print'
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
    $suffix = switch ($what) { 'sheet' { '_粘贴单' } 'inv' { '_发票' } default { '_打印' } }
    $path = Join-Path $dir ((Get-SafeName $b.no) + $suffix + '.pdf')
    try { [IO.File]::WriteAllBytes($path, $out.Save()) }
    catch { $path = Join-Path $dir ((Get-SafeName $b.no) + $suffix + '_' + (Get-Date).ToString('HHmmss') + '.pdf'); [IO.File]::WriteAllBytes($path, $out.Save()) }
    Start-Process -FilePath $path
    if ($miss.Count -gt 0) { [System.Windows.MessageBox]::Show($script:CalWin, ("已经拼好 " + $out.PageCount + " 页，打开后直接打印就行。`n`n下面这些没拼进去：`n" + ($miss -join "`n")), '打印') | Out-Null }
    else { Notify '打印文件已生成' ('共 ' + $out.PageCount + ' 页（A4），在打开的 PDF 里按 Ctrl+P 打印。粘贴单那页是横版。') }
}

# ---------- 报销设置 ----------
function Show-ReimbSettings {
    $script:Dlg = New-DialogWindow '报销设置'
    if ($script:CalWin) { $script:Dlg.Owner = $script:CalWin; $script:Dlg.WindowStartupLocation = 'CenterOwner' }
    $sp = [System.Windows.Controls.StackPanel]::new(); $sp.Margin = Th 16 14 16 14; $sp.Width = 420
    $mk = {
        param($label, $val, $note, $multi)
        $l = [System.Windows.Controls.TextBlock]::new(); $l.Text = $label; $l.Margin = Th 0 8 0 3; $l.FontWeight = [System.Windows.FontWeights]::SemiBold
        [void]$sp.Children.Add($l)
        if ($note) { $n = [System.Windows.Controls.TextBlock]::new(); $n.Text = $note; $n.TextWrapping = 'Wrap'; $n.FontSize = 11; $n.Foreground = Brush '#6E6450'; $n.Margin = Th 0 0 0 3; [void]$sp.Children.Add($n) }
        $tb = [System.Windows.Controls.TextBox]::new(); $tb.Text = $val; $tb.Padding = Th 4 3 4 3; $tb.TextWrapping = 'Wrap'
        if ($multi) { $tb.AcceptsReturn = $true; $tb.MinHeight = 54 }
        [void]$sp.Children.Add($tb)
        return $tb
    }
    $tName = & $mk '我的姓名' ([string]$script:S.reimbName) '发票改名用（姓名＋发票号码），新建报销单时默认的报销人、经办人。' $false
    $tDept = & $mk '默认成本中心 / 部门' (Get-ReimbDept) '新建报销单时自动填这个；给别人代报销时，在那张报销单里单独改。' $false
    $tNote = & $mk '「其他事项说明」固定写的话' ([string]$script:S.reimbNote) '比如收款银行卡信息。OA 填报页会把它和市内交通费明细拼在一起，点一下就能复制。只存在本机设置里，不进同步文件。' $true
    $cCap = [System.Windows.Controls.CheckBox]::new(); $cCap.Content = '打印时在每张发票下面标「第几张 / 共几张」'; $cCap.Margin = Th 0 12 0 0; $cCap.IsChecked = [bool]$script:S.reimbCaption
    [void]$sp.Children.Add($cCap)
    $btns = New-DialogButtons; [void]$sp.Children.Add($btns[0])
    $script:RsBoxes = @($tName, $tDept, $tNote, $cCap)
    $btns[1].add_Click({
        $script:S.reimbName = $script:RsBoxes[0].Text.Trim()
        $script:S.reimbDept = $script:RsBoxes[1].Text.Trim()
        $script:S.reimbNote = $script:RsBoxes[2].Text.Trim()
        $script:S.reimbCaption = [bool]$script:RsBoxes[3].IsChecked
        Save-Settings
        $script:Dlg.DialogResult = $true
    })
    $script:Dlg.Content = $sp
    if ($script:Dlg.ShowDialog()) { Render }
}

# ---------- 界面 ----------
function New-RbText([string]$text, [double]$size = 13, [string]$brush = 'TextBrush') {
    $t = [System.Windows.Controls.TextBlock]::new()
    $t.Text = $text; $t.FontSize = $size; $t.TextWrapping = 'Wrap'
    $t.SetResourceReference([System.Windows.Controls.TextBlock]::ForegroundProperty, $brush)
    return $t
}
$StateColor = @{ '未提交' = '#C9A227'; '已提交' = '#2F6FED'; '已退回' = '#D9534F'; '已到账' = '#2E9E5B' }

$OnBillPick = {
    param($s, $e)
    if ($e.ClickCount -eq 2) { return }
    $id = [string]$s.Tag
    if ($script:ReimbSel -eq $id) { if ($script:ReimbEdit) { $script:ReimbEdit = ''; Render-Reimb }; return }
    Flush-CalFocus
    $script:ReimbSel = $id; $script:ReimbEdit = ''
    Render-Reimb
}
$OnExpPick = {
    param($s, $e)
    Flush-CalFocus
    $id = [string]$s.Tag
    if ($e.ClickCount -eq 2) { $x = Find-Exp $id; if ($x -and $x.file) { Open-ExpFile $x.file }; return }
    $script:ReimbEdit = $id
    Render-Reimb
}
$OnRbFilter = { param($s, $e) $script:ReimbFilter = [string]$s.Tag; Render-Reimb }
$OnCopyCell = {
    param($s, $e)
    $v = [string]$s.Tag
    if (-not $v) { return }
    try { [System.Windows.Clipboard]::SetText($v) } catch { Notify '复制失败' '剪贴板被别的程序占用了，稍后再试一次。'; return }
    $s.Opacity = 0.4
    $script:RbFlash = $s
    if (-not $script:RbFlashTimer) {
        $script:RbFlashTimer = [System.Windows.Threading.DispatcherTimer]::new()
        $script:RbFlashTimer.Interval = [TimeSpan]::FromMilliseconds(350)
        $script:RbFlashTimer.add_Tick({ $script:RbFlashTimer.Stop(); try { $script:RbFlash.Opacity = 1 } catch {} })
    }
    $script:RbFlashTimer.Stop(); $script:RbFlashTimer.Start()
    if ($script:RbCopyTip) { $script:RbCopyTip.Text = '已复制：' + $(if ($v.Length -gt 40) { $v.Substring(0, 40) + '…' } else { $v }) }
}

function Render-BillList {
    $lb = $script:RbList
    $lb.Children.Clear()
    $f = $script:RbFilter
    $f.Children.Clear()
    foreach ($st in @('') + $BillStates) {
        $n = @($script:Bills | Where-Object { -not $st -or $_.state -eq $st }).Count
        $t = New-RbText ($(if ($st) { $st } else { '全部' }) + ' ' + $n) 12 $(if ($script:ReimbFilter -eq $st) { 'AccentBrush' } else { 'SubBrush' })
        $t.Margin = Th 0 0 12 2; $t.Cursor = $Hand; $t.Tag = $st
        if ($script:ReimbFilter -eq $st) { $t.FontWeight = [System.Windows.FontWeights]::Bold }
        $t.add_MouseLeftButtonUp($OnRbFilter)
        [void]$f.Children.Add($t)
    }
    $unsub = 0.0; $pend = 0.0
    foreach ($b in $script:Bills) { $t = Get-BillTotal $b.id; if ($b.state -eq '未提交' -or $b.state -eq '已退回') { $unsub += $t } elseif ($b.state -eq '已提交') { $pend += $t } }
    $script:RbSum.Text = '还没报：¥' + (Format-Amt $unsub) + '　　已提交等到账：¥' + (Format-Amt $pend)
    $bs = @($script:Bills | Where-Object { -not $script:ReimbFilter -or $_.state -eq $script:ReimbFilter } | Sort-Object @{ Expression = { if ($_.start) { $_.start } else { $_.created } }; Descending = $true }, @{ Expression = { $_.no }; Descending = $true })
    if ($bs.Count -eq 0) {
        $em = New-RbText $(if ($script:Bills.Count -eq 0) { '还没有报销单。点「新建报销单」，或者直接把发票文件拖到这个页面上。' } else { '这一类下面没有报销单。' }) 12 'MutedBrush'
        [void]$lb.Children.Add($em)
    }
    foreach ($b in $bs) {
        $xs = @(Get-BillExps $b.id)
        $bd = [System.Windows.Controls.Border]::new()
        $bd.Padding = Th 10 7 10 7; $bd.Margin = Th 0 0 0 4; $bd.CornerRadius = [System.Windows.CornerRadius]::new(5); $bd.Cursor = $Hand; $bd.Tag = $b.id
        $bd.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, $(if ($script:ReimbSel -eq $b.id) { 'HoverBrush' } else { 'CardBrush' }))
        $sp = [System.Windows.Controls.StackPanel]::new()
        $t1 = New-RbText $b.text 13; $t1.FontWeight = [System.Windows.FontWeights]::SemiBold
        $dr = ''
        $d1 = Parse-Day $b.start; $d2 = Parse-Day $b.end
        if ($d1) { $dr = $d1.ToString('M/d'); if ($d2 -and $d2 -ne $d1) { $dr += '—' + $d2.ToString('M/d') } }
        $t2 = New-RbText ((@($b.no, $dr, $(if ($b.proj) { Get-ProjName $b.proj } else { '' })) | Where-Object { $_ }) -join ' · ') 11 'SubBrush'
        $row = [System.Windows.Controls.WrapPanel]::new(); $row.Margin = Th 0 2 0 0
        $am = New-RbText ('¥' + (Format-Amt (Get-BillTotal $b.id)) + '  ' + @(Get-BillFiles $b.id).Count + ' 张') 12
        $am.Margin = Th 0 0 10 0
        $sb = [System.Windows.Controls.Border]::new(); $sb.CornerRadius = [System.Windows.CornerRadius]::new(3); $sb.Padding = Th 5 0 5 1; $sb.Background = Brush $StateColor[$b.state]
        $stt = [System.Windows.Controls.TextBlock]::new(); $stt.Text = $b.state; $stt.FontSize = 11; $stt.Foreground = Brush '#FFFFFF'; $sb.Child = $stt
        [void]$row.Children.Add($am); [void]$row.Children.Add($sb)
        $todo = @($xs | Where-Object { $_.status -eq '待补' -or -not $_.amount }).Count
        if ($todo -gt 0) { $w = New-RbText ('　' + $todo + ' 笔待补') 11 'SubBrush'; $w.Foreground = Brush '#D9534F'; [void]$row.Children.Add($w) }
        [void]$sp.Children.Add($t1); [void]$sp.Children.Add($t2); [void]$sp.Children.Add($row)
        $bd.Child = $sp
        $bd.add_MouseLeftButtonUp($OnBillPick)
        [void]$lb.Children.Add($bd)
    }
}

function New-ExpRow($x) {
    $g = [System.Windows.Controls.Grid]::new()
    foreach ($w in 92, -1, 64, 86, 52) { $c = [System.Windows.Controls.ColumnDefinition]::new(); $c.Width = if ($w -lt 0) { [System.Windows.GridLength]::new(1, 'Star') } else { [System.Windows.GridLength]::new($w) }; [void]$g.ColumnDefinitions.Add($c) }
    $when = ''
    if ($x.date -match '^\d{4}-(\d{2})-(\d{2})$') { $when = ([int]$Matches[1]).ToString() + '/' + ([int]$Matches[2]).ToString() }
    if ($x.time) { $when += ' ' + $x.time }
    $c0 = New-RbText $when 12 'SubBrush'; $c0.TextWrapping = 'NoWrap'
    $txt = $x.text
    if ($x.cat -eq '城市间交通费' -and $x.note -match '^([A-Z0-9]{1,2}\d{1,4})\b' -and $txt -notmatch $Matches[1]) { $txt += ' ' + $Matches[1] }
    $c1 = New-RbText $txt 13
    $c2 = New-RbText ([string]$x.person) 12 'SubBrush'; $c2.TextWrapping = 'NoWrap'
    $c3 = New-RbText $(if ($x.amount) { Format-Amt (Get-Amt $x) } else { '金额？' }) 13
    $c3.HorizontalAlignment = 'Right'; $c3.Margin = Th 0 0 10 0
    if (-not $x.amount) { $c3.Foreground = Brush '#D9534F' }
    $flag = if ($x.file) { '📎' } elseif ($x.status -eq '无票') { '无票' } else { '待补' }
    $c4 = New-RbText $flag 11 'SubBrush'; $c4.TextWrapping = 'NoWrap'
    if (-not $x.file -and $x.status -eq '待补') { $c4.Foreground = Brush '#D9534F' }
    if ($x.file) { $c4.ToolTip = '双击这一行打开发票：' + (Split-Path -Leaf $x.file) }
    $i = 0; foreach ($c in $c0, $c1, $c2, $c3, $c4) { [System.Windows.Controls.Grid]::SetColumn($c, $i); $c.VerticalAlignment = 'Center'; [void]$g.Children.Add($c); $i++ }
    $bd = [System.Windows.Controls.Border]::new()
    $bd.Child = $g; $bd.Padding = Th 6 4 4 4; $bd.CornerRadius = [System.Windows.CornerRadius]::new(4); $bd.Cursor = $Hand; $bd.Tag = $x.id
    $bd.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, $(if ($script:ReimbEdit -eq $x.id) { 'HoverBrush' } else { 'CardBrush' }))
    $bd.add_MouseLeftButtonUp($OnExpPick)
    $cm = [System.Windows.Controls.ContextMenu]::new()
    if ($x.file) { [void]$cm.Items.Add((New-MenuItem '打开发票' $x.file { param($s, $e) Open-ExpFile ([string]$s.Tag) })) }
    [void]$cm.Items.Add((New-MenuItem $(if ($x.file) { '换一个发票文件…' } else { '添加发票文件…' }) $x.id { param($s, $e) Set-ExpFile ([string]$s.Tag) }))
    [void]$cm.Items.Add((New-MenuItem '拆成几笔…' $x.id { param($s, $e) Split-Expense ([string]$s.Tag) }))
    $mv = [System.Windows.Controls.MenuItem]::new(); $mv.Header = '挪到别的报销单'
    foreach ($b in @($script:Bills | Where-Object { $_.id -ne $x.bill } | Sort-Object no -Descending | Select-Object -First 15)) { [void]$mv.Items.Add((New-MenuItem ($b.no + '  ' + $b.text) ($x.id + '|' + $b.id) { param($s, $e) $p = ([string]$s.Tag).Split('|'); Move-ExpToBill $p[0] $p[1] })) }
    if ($mv.Items.Count -gt 0) { [void]$cm.Items.Add($mv) }
    [void]$cm.Items.Add([System.Windows.Controls.Separator]::new())
    [void]$cm.Items.Add((New-MenuItem '删除这一笔' $x.id { param($s, $e) Remove-Expense ([string]$s.Tag) }))
    $bd.ContextMenu = $cm
    return $bd
}

function Render-ExpList($b) {
    $v = $script:RbViewPanel
    $xs = @(Get-BillExps $b.id)
    $hd = New-RbText $b.text 17; $hd.FontWeight = [System.Windows.FontWeights]::SemiBold
    [void]$v.Children.Add($hd)
    $dr = (($b.start + ' 至 ' + $b.end).Trim(' 至'))
    $meta = New-RbText ((@($b.no, $dr, $b.place, $(if ($b.proj) { Get-ProjName $b.proj } else { '' }), $b.state) | Where-Object { $_ }) -join '　·　') 12 'SubBrush'
    $meta.Margin = Th 0 2 0 10
    [void]$v.Children.Add($meta)
    if ($xs.Count -eq 0) {
        $em = New-RbText '这张报销单还是空的。点上面「导入发票」选文件，或者直接把下载好的发票拖进这个窗口；没有发票的（补助、无票的打车）在上面输入框里记一笔。' 13 'MutedBrush'
        [void]$v.Children.Add($em); return
    }
    foreach ($c in $ReimbCats) {
        $g = @($xs | Where-Object { $_.cat -eq $c }); if ($g.Count -eq 0) { continue }
        $t = 0.0; foreach ($x in $g) { $t += Get-Amt $x }
        $gh = [System.Windows.Controls.DockPanel]::new(); $gh.Margin = Th 0 12 0 4
        $ga = New-RbText ('¥' + (Format-Amt $t)) 13; $ga.FontWeight = [System.Windows.FontWeights]::SemiBold; $ga.Margin = Th 0 0 62 0
        [System.Windows.Controls.DockPanel]::SetDock($ga, 'Right'); [void]$gh.Children.Add($ga)
        $gt = New-RbText ($c + '（' + $g.Count + ' 笔）') 13 'SubBrush'; $gt.FontWeight = [System.Windows.FontWeights]::SemiBold
        [void]$gh.Children.Add($gt)
        [void]$v.Children.Add($gh)
        foreach ($x in $g) { [void]$v.Children.Add((New-ExpRow $x)) }
    }
    $line = [System.Windows.Controls.Border]::new(); $line.Height = 1; $line.Margin = Th 0 12 0 8
    $line.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'LineBrush')
    [void]$v.Children.Add($line)
    $tot = Get-BillTotal $b.id
    $sumT = New-RbText ('合计 ¥' + (Format-Amt $tot) + '　　' + $xs.Count + ' 笔，发票 ' + @(Get-BillFiles $b.id).Count + ' 张') 15
    $sumT.FontWeight = [System.Windows.FontWeights]::SemiBold
    [void]$v.Children.Add($sumT)
    $warn = @()
    $n1 = @($xs | Where-Object { -not $_.amount }).Count; if ($n1) { $warn += ($n1.ToString() + ' 笔没填金额') }
    $n2 = @($xs | Where-Object { $_.status -eq '待补' -and -not $_.file }).Count; if ($n2) { $warn += ($n2.ToString() + ' 笔票据待补') }
    $n3 = @($xs | Where-Object { $_.seg -and -not $_.date2 }).Count; if ($n3) { $warn += ($n3.ToString() + ' 段行程没填到达时间（OA 要填，点那一笔补上）') }
    if ($warn.Count -gt 0) { $w = New-RbText ('还差：' + ($warn -join '；')) 12; $w.Foreground = Brush '#D9534F'; $w.Margin = Th 0 4 0 0; [void]$v.Children.Add($w) }
    $tip = New-RbText '单击一笔在右边改；双击打开发票；右键可以拆分、挪到别的报销单、删除。' 11 'MutedBrush'; $tip.Margin = Th 0 10 0 0
    [void]$v.Children.Add($tip)
}

function New-CopyCell([string]$text, [string]$copy, [double]$size = 13) {
    $t = New-RbText $text $size
    if ($null -eq $copy) { $copy = $text }
    $t.Tag = $copy
    if ($copy) { $t.Cursor = $Hand; $t.ToolTip = '点一下复制'; $t.add_MouseLeftButtonUp($OnCopyCell) }
    return $t
}
function Add-OAField($panel, [string]$label, [string]$value) {
    $dp = [System.Windows.Controls.DockPanel]::new(); $dp.Margin = Th 0 2 0 2
    $l = New-RbText $label 12 'SubBrush'; $l.Width = 90
    [void]$dp.Children.Add($l)
    $vv = New-CopyCell $(if ($value) { $value } else { '—' }) $value
    [void]$dp.Children.Add($vv)
    [void]$panel.Children.Add($dp)
}

function Render-OAView($b) {
    $v = $script:RbViewPanel
    $oa = Get-OAData $b
    $hd = New-RbText ('OA 填报　' + $b.no) 17; $hd.FontWeight = [System.Windows.FontWeights]::SemiBold
    [void]$v.Children.Add($hd)
    $script:RbCopyTip = New-RbText '照着 OA 的格子一项一项填：点任意一个值就复制，切到 OA 里粘贴。' 12 'MutedBrush'
    $script:RbCopyTip.Margin = Th 0 2 0 10
    [void]$v.Children.Add($script:RbCopyTip)
    Add-OAField $v '报销事由' $b.text
    Add-OAField $v '出差日期' ((($b.start + ' 至 ' + $b.end).Trim(' 至')))
    Add-OAField $v '成本中心' $b.dept
    Add-OAField $v '项目' $(if ($b.proj) { Get-ProjName $b.proj } else { '' })
    Add-OAField $v '出差地点' $b.place
    if ($b.people) { Add-OAField $v '出差人员' $b.people }

    [void]$v.Children.Add((New-PanelLabel '行程（每一段一行：出发 → 到达，下面是这段的交通费、住宿费）'))
    if ($oa.segs.Count -eq 0) { [void]$v.Children.Add((New-RbText '没有算作行程的费用。高铁、飞机、租车默认算一段；别的要算的，点开那一笔勾上「算作一段行程」。' 12 'MutedBrush')) }
    $g = [System.Windows.Controls.Grid]::new()
    foreach ($w in 28, 86, 52, -1, 86, 52, -1, 76, 76) { $c = [System.Windows.Controls.ColumnDefinition]::new(); $c.Width = if ($w -lt 0) { [System.Windows.GridLength]::new(1, 'Star') } else { [System.Windows.GridLength]::new($w) }; [void]$g.ColumnDefinitions.Add($c) }
    $addRow = {
        param([int]$r, $cells, [bool]$head)
        $rd = [System.Windows.Controls.RowDefinition]::new(); $rd.Height = [System.Windows.GridLength]::Auto; [void]$g.RowDefinitions.Add($rd)
        for ($i = 0; $i -lt $cells.Count; $i++) {
            $cv = $cells[$i]
            $t = if ($head) { New-RbText ([string]$cv) 11 'SubBrush' } elseif ($cv -is [array]) { New-CopyCell ([string]$cv[0]) ([string]$cv[1]) } else { New-CopyCell ([string]$cv) ([string]$cv) }
            $t.Margin = Th 0 3 6 3
            if ($i -ge 7) { $t.HorizontalAlignment = 'Right' }
            [System.Windows.Controls.Grid]::SetRow($t, $r); [System.Windows.Controls.Grid]::SetColumn($t, $i)
            [void]$g.Children.Add($t)
        }
    }
    if ($oa.segs.Count -gt 0) {
        & $addRow 0 @('', '出发日期', '时', '出发地', '到达日期', '时', '到达地', '交通费', '住宿费') $true
        $r = 1
        foreach ($sg in $oa.segs) {
            $hm = { param($t) if ($t -match '^(\d{1,2}):(\d{2})') { $h = [int]$Matches[1]; $m = [int]$Matches[2]; if ($m) { return @(($h.ToString() + '时' + $m + '分'), $h.ToString()) }; return @(($h.ToString() + '时'), $h.ToString()) }; return @('—', '') }
            $d1 = if (Parse-Day $sg.date) { @((Parse-Day $sg.date).ToString('M月d日'), $sg.date) } elseif ($sg.date) { @($sg.date, $sg.date) } else { @('—', '') }
            $d2 = if (Parse-Day $sg.date2) { @((Parse-Day $sg.date2).ToString('M月d日'), $sg.date2) } elseif ($sg.date2) { @($sg.date2, $sg.date2) } else { @('待填', '') }
            $tr = Format-AmtRaw $sg.traffic
            $ho = if ($sg.hotel -gt 0) { Format-AmtRaw $sg.hotel } else { '' }
            & $addRow $r @([string]$r, $d1, (& $hm $sg.time), $sg.from, $d2, (& $hm $sg.time2), $sg.to, @($tr, $tr), @($(if ($ho) { $ho } else { '—' }), $ho)) $false
            $r++
        }
        & $addRow $r @('', '', '', '', '', '', '合计', (Format-AmtRaw $oa.traffic), (Format-AmtRaw ($oa.hotel - $oa.hotelLeft))) $false
    }
    [void]$v.Children.Add($g)
    if ($oa.hotelLeft -gt 0) { [void]$v.Children.Add((New-RbText ('还有 ' + (Format-AmtRaw $oa.hotelLeft) + ' 元住宿费没对上行程（没有行程段），自己填到合适的那一行。') 12 'MutedBrush')) }
    $nb = New-RbText '住宿费放在「到达住宿地」那一段；同一趟车几个人的票会合成一行。' 11 'MutedBrush'; $nb.Margin = Th 0 4 0 0
    [void]$v.Children.Add($nb)

    [void]$v.Children.Add((New-PanelLabel '市内交通费'))
    if ($oa.city.Count -gt 0) {
        Add-OAField $v '金额' (Format-AmtShort $oa.cityTotal)
        Add-OAField $v '说明' $oa.cityDesc
    } else { [void]$v.Children.Add((New-RbText '没有。' 12 'MutedBrush')) }
    if ($oa.otherByCat.Count -gt 0) {
        [void]$v.Children.Add((New-PanelLabel '其他费用'))
        foreach ($k in $oa.otherByCat.Keys) { Add-OAField $v $k (Format-AmtRaw $oa.otherByCat[$k]) }
    }
    [void]$v.Children.Add((New-PanelLabel '其他事项说明'))
    $memo = New-CopyCell $(if ($oa.memo) { $oa.memo } else { '（空）' }) $oa.memo
    [void]$v.Children.Add($memo)
    if (-not ([string]$script:S.reimbNote).Trim()) { $tip = New-RbText '前面固定的那句（比如收款银行卡）可以在「设置」里填一次，以后每张单都会带上。' 11 'MutedBrush'; [void]$v.Children.Add($tip) }
    if ($b.note) { [void]$v.Children.Add((New-PanelLabel '情况说明')); [void]$v.Children.Add((New-CopyCell $b.note $b.note)) }
    $line = [System.Windows.Controls.Border]::new(); $line.Height = 1; $line.Margin = Th 0 12 0 8
    $line.SetResourceReference([System.Windows.Controls.Border]::BackgroundProperty, 'LineBrush')
    [void]$v.Children.Add($line)
    $tt = New-CopyCell ('总计 ¥' + (Format-Amt $oa.total) + '　单据 ' + @(Get-BillFiles $b.id).Count + ' 张') (Format-AmtRaw $oa.total) 15
    $tt.FontWeight = [System.Windows.FontWeights]::SemiBold
    [void]$v.Children.Add($tt)
}

# 右边：报销单信息，或者一笔费用
$OnBillField = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|'); $b = Find-Bill $p[0]; if (-not $b) { return }
    $v = $s.Text.Trim()
    if ($p[1] -eq 'text' -and -not $v) { $s.Text = $b.text; return }
    if ($p[1] -eq 'start' -or $p[1] -eq 'end' -or $p[1] -eq 'submit') { if ($v) { $iso = ConvertTo-IsoDay $v; if (-not $iso -and $v -match '^(\d{1,2})[./月-](\d{1,2})') { try { $iso = ([datetime]::new((Get-Date).Year, [int]$Matches[1], [int]$Matches[2])).ToString('yyyy-MM-dd') } catch {} }; if ($iso) { $v = $iso } } }
    if ([string]$b.($p[1]) -ne $v) { $b.($p[1]) = $v; Save-Bills; Render }
}
$OnExpField = {
    param($s, $e)
    $p = ([string]$s.Tag).Split('|'); $x = Find-Exp $p[0]; if (-not $x) { return }
    $f = $p[1]; $v = $s.Text.Trim()
    if ($f -eq 'date' -or $f -eq 'date2' -or $f -eq 'invDate') { if ($v) { $iso = ConvertTo-IsoDay $v; if (-not $iso -and $v -match '^(\d{1,2})[./月-](\d{1,2})') { $y0 = if ($x.date) { [int]$x.date.Substring(0, 4) } else { (Get-Date).Year }; try { $iso = ([datetime]::new($y0, [int]$Matches[1], [int]$Matches[2])).ToString('yyyy-MM-dd') } catch {} }; if ($iso) { $v = $iso } } }
    if ($f -eq 'time' -or $f -eq 'time2') { if ($v -match '^(\d{1,2})(?:[:：点时.](\d{1,2}))?') { $v = ('{0:D2}:{1:D2}' -f [int]$Matches[1], $(if ($Matches[2]) { [int]$Matches[2] } else { 0 })) } }
    if ($f -eq 'amount' -and $v) { $d = Get-Amt $v; $v = Format-AmtRaw $d }
    if ($f -eq 'text' -and -not $v) { $v = Get-AutoTitle $x }
    if ([string]$x.$f -ne $v) {
        $old = Get-AutoTitle $x
        $x.$f = $v
        if (@('from', 'to', 'nights') -contains $f -and $x.text -eq $old) { $x.text = Get-AutoTitle $x }
        $b = Find-Bill $x.bill; if ($b -and ($f -eq 'date' -or $f -eq 'date2')) { Update-BillFromExps $b; Save-Bills }
        Save-Exps; Render
    }
}
function Add-SideField($p, [string]$label, [string]$value, [string]$tag, [scriptblock]$h, [bool]$multi = $false) {
    [void]$p.Children.Add((New-PanelLabel $label))
    $tb = New-EditBox $value $multi
    $tb.Tag = $tag
    Commit-TextBox $tb $h
    [void]$p.Children.Add($tb)
    return $tb
}
function Add-SidePair($p, [string]$l1, [string]$v1, [string]$t1, [string]$l2, [string]$v2, [string]$t2, [scriptblock]$h) {
    $g = [System.Windows.Controls.Grid]::new()
    foreach ($w in 1, 1) { $c = [System.Windows.Controls.ColumnDefinition]::new(); $c.Width = [System.Windows.GridLength]::new(1, 'Star'); [void]$g.ColumnDefinitions.Add($c) }
    $i = 0
    foreach ($q in @(@($l1, $v1, $t1), @($l2, $v2, $t2))) {
        $sp = [System.Windows.Controls.StackPanel]::new(); $sp.Margin = if ($i -eq 0) { Th 0 0 6 0 } else { Th 6 0 0 0 }
        [void]$sp.Children.Add((New-PanelLabel $q[0]))
        $tb = New-EditBox $q[1] $false; $tb.Tag = $q[2]; Commit-TextBox $tb $h
        [void]$sp.Children.Add($tb)
        [System.Windows.Controls.Grid]::SetColumn($sp, $i); [void]$g.Children.Add($sp); $i++
    }
    [void]$p.Children.Add($g)
}
function New-SimpleCombo($items, [string]$current, [string]$tag, [scriptblock]$onChange) {
    $cb = [System.Windows.Controls.ComboBox]::new()
    foreach ($i in $items) { [void]$cb.Items.Add($i) }
    $cb.SelectedItem = $current; $cb.Tag = $tag
    $cb.add_SelectionChanged($onChange)
    return $cb
}

function Render-BillSide($b) {
    $p = $script:RbSide
    [void]$p.Children.Add((New-PanelTitle '报销单信息' 15))
    $id = $b.id
    [void](Add-SideField $p '事由' $b.text ($id + '|text') $OnBillField $true)
    Add-SidePair $p '单号' $b.no ($id + '|no') 'OA 单号' $b.oaNo ($id + '|oaNo') $OnBillField
    [void]$p.Children.Add((New-PanelLabel '项目'))
    $pc = New-ProjCombo $b.proj { param($s, $e) $bb = Find-Bill ([string]$s.Tag); $np = Get-ComboProj $s; if ($bb -and $np -and $np -ne $bb.proj) { $bb.proj = $np; Save-Bills; Render } } $id
    [void]$p.Children.Add($pc)
    [void](Add-SideField $p '出差地点' $b.place ($id + '|place') $OnBillField)
    Add-SidePair $p '开始日期' $b.start ($id + '|start') '结束日期' $b.end ($id + '|end') $OnBillField
    Add-SidePair $p '报销人 / 经办人' $b.person ($id + '|person') '出差人员' $b.people ($id + '|people') $OnBillField
    [void](Add-SideField $p '成本中心 / 部门' $b.dept ($id + '|dept') $OnBillField)
    [void]$p.Children.Add((New-PanelLabel '报销状态'))
    $sc = New-SimpleCombo $BillStates $b.state $id {
        param($s, $e)
        $bb = Find-Bill ([string]$s.Tag); $v = [string]$s.SelectedItem
        if ($bb -and $v -and $v -ne $bb.state) { $bb.state = $v; if ($v -eq '已提交' -and -not $bb.submit) { $bb.submit = (Get-Date).ToString('yyyy-MM-dd') }; Save-Bills; Render }
    }
    [void]$p.Children.Add($sc)
    [void](Add-SideField $p '提交日期' $b.submit ($id + '|submit') $OnBillField)
    [void](Add-SideField $p '情况说明（为什么出差、超标原因等，OA 和 Excel 里都会带上）' $b.note ($id + '|note') $OnBillField $true)
    $row = [System.Windows.Controls.WrapPanel]::new(); $row.Margin = Th 0 14 0 0
    [void]$row.Children.Add((New-LinkText '删除这张报销单' $id { param($s, $e) Remove-Bill ([string]$s.Tag) }))
    [void]$p.Children.Add($row)
}

function Render-ExpSide($x) {
    $p = $script:RbSide
    [void]$p.Children.Add((New-LinkText '‹ 返回报销单信息' '' { Flush-CalFocus; $script:ReimbEdit = ''; Render-Reimb }))
    $p.Children[$p.Children.Count - 1].Margin = Th 0 0 0 10
    [void]$p.Children.Add((New-PanelTitle '一笔费用' 15))
    $id = $x.id
    [void](Add-SideField $p '事项 / 行程' $x.text ($id + '|text') $OnExpField)
    [void]$p.Children.Add((New-PanelLabel '类别'))
    $cc = New-SimpleCombo $ReimbCats $x.cat $id {
        param($s, $e)
        $xx = Find-Exp ([string]$s.Tag); $v = [string]$s.SelectedItem
        if ($xx -and $v -and $v -ne $xx.cat) { $old = Get-AutoTitle $xx; $xx.cat = $v; if ($xx.text -eq $old) { $xx.text = Get-AutoTitle $xx }; if (@('城市间交通费', '租车/包车费') -contains $v) { $xx.seg = $true } elseif ($v -ne '市内交通费') { $xx.seg = $false }; Save-Exps; Render }
    }
    [void]$p.Children.Add($cc)
    Add-SidePair $p '金额（元）' $x.amount ($id + '|amount') '人员' $x.person ($id + '|person') $OnExpField
    $isHotel = $x.cat -eq '住宿费'
    Add-SidePair $p $(if ($isHotel) { '入住日期' } else { '日期' }) $x.date ($id + '|date') $(if ($isHotel) { '几晚' } else { '时间（几点）' }) $(if ($isHotel) { $x.nights } else { $x.time }) ($id + $(if ($isHotel) { '|nights' } else { '|time' })) $OnExpField
    if ($isHotel) { [void](Add-SideField $p '酒店' $x.to ($id + '|to') $OnExpField); [void](Add-SideField $p '退房日期' $x.date2 ($id + '|date2') $OnExpField) }
    else {
        Add-SidePair $p '出发地' $x.from ($id + '|from') '目的地' $x.to ($id + '|to') $OnExpField
        if ($x.seg -or $x.cat -eq '城市间交通费') { Add-SidePair $p '到达日期' $x.date2 ($id + '|date2') '到达时间' $x.time2 ($id + '|time2') $OnExpField }
    }
    $sg = [System.Windows.Controls.CheckBox]::new()
    $sg.Content = '算作一段行程（OA 里单独一行）'; $sg.IsChecked = [bool]$x.seg; $sg.Tag = $id; $sg.Margin = Th 0 10 0 0
    $sg.SetResourceReference([System.Windows.Controls.Control]::ForegroundProperty, 'TextBrush')
    $sg.add_Click({ param($s, $e) $xx = Find-Exp ([string]$s.Tag); if ($xx) { $xx.seg = [bool]$s.IsChecked; Save-Exps; Render } })
    if (-not $isHotel) { [void]$p.Children.Add($sg) }
    [void]$p.Children.Add((New-PanelLabel '票据'))
    $tc = New-SimpleCombo $TicketStates $x.status $id { param($s, $e) $xx = Find-Exp ([string]$s.Tag); $v = [string]$s.SelectedItem; if ($xx -and $v -and $v -ne $xx.status) { $xx.status = $v; Save-Exps; Render } }
    [void]$p.Children.Add($tc)
    Add-SidePair $p '发票号码' $x.invNo ($id + '|invNo') '开票日期' $x.invDate ($id + '|invDate') $OnExpField
    Add-SidePair $p '票据类型' $x.invType ($id + '|invType') '开票方' $x.seller ($id + '|seller') $OnExpField
    [void](Add-SideField $p '备注' $x.note ($id + '|note') $OnExpField $true)
    $row = [System.Windows.Controls.WrapPanel]::new(); $row.Margin = Th 0 12 0 0
    if ($x.file) {
        [void]$row.Children.Add((New-LinkText '打开发票' $x.file { param($s, $e) Open-ExpFile ([string]$s.Tag) }))
        $fn = New-RbText (Split-Path -Leaf $x.file) 11 'MutedBrush'; $fn.Margin = Th 0 6 0 0
    }
    [void]$row.Children.Add((New-LinkText $(if ($x.file) { '换文件' } else { '添加发票文件' }) $id { param($s, $e) Set-ExpFile ([string]$s.Tag) }))
    [void]$row.Children.Add((New-LinkText '拆成几笔' $id { param($s, $e) Split-Expense ([string]$s.Tag) }))
    [void]$row.Children.Add((New-LinkText '删除' $id { param($s, $e) Remove-Expense ([string]$s.Tag) }))
    [void]$p.Children.Add($row)
    if ($x.file) { [void]$p.Children.Add($fn) }
}

function Render-Reimb {
    if (-not $script:RbList) { return }
    if ($script:ReimbSel -and -not (Find-Bill $script:ReimbSel)) { $script:ReimbSel = '' }
    if (-not $script:ReimbSel) {
        $first = @($script:Bills | Sort-Object @{ Expression = { if ($_.start) { $_.start } else { $_.created } }; Descending = $true } | Select-Object -First 1)
        if ($first.Count -gt 0) { $script:ReimbSel = $first[0].id }
    }
    Render-BillList
    $b = Find-Bill $script:ReimbSel
    $v = $script:RbViewPanel
    $v.Children.Clear()
    $script:RbSide.Children.Clear()
    $script:RbCopyTip = $null
    foreach ($n in 'RbImport', 'RbExcel', 'RbPrint', 'RbFolder', 'RbTabList', 'RbTabOA') { $script:RbBtns[$n].IsEnabled = [bool]$b }
    foreach ($n in 'RbTabList', 'RbTabOA') {
        $on = ($n -eq 'RbTabOA') -eq ($script:ReimbView -eq 'oa')
        $script:RbBtns[$n].FontWeight = if ($on) { [System.Windows.FontWeights]::Bold } else { [System.Windows.FontWeights]::Normal }
    }
    if (-not $b) {
        [void]$v.Children.Add((New-RbText '左边点一张报销单；还没有的话点「新建报销单」，或者直接把发票拖进来，会自动建一张。' 13 'MutedBrush'))
        [void]$script:RbSide.Children.Add((New-RbText '出差回来：把下载的发票拖进来 → 自动改名、认出日期车次金额 → 「OA 填报」照着点一下复制一项 → 「打印」出粘贴单和每张一页的发票。' 12 'MutedBrush'))
        return
    }
    if ($script:ReimbView -eq 'oa') { Render-OAView $b } else { Render-ExpList $b }
    $x = if ($script:ReimbEdit) { Find-Exp $script:ReimbEdit } else { $null }
    if ($x -and $x.bill -eq $b.id) { Render-ExpSide $x } else { $script:ReimbEdit = ''; Render-BillSide $b }
}

$OnReimbDrop = {
    param($s, $e)
    if (-not $e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { return }
    $e.Handled = $true
    $files = @($e.Data.GetData([System.Windows.DataFormats]::FileDrop))
    $script:CalWin.Activate() | Out-Null
    Import-ReimbFiles ([string[]]$files)
}
$OnReimbDragOver = {
    param($s, $e)
    $e.Effects = if ($e.Data.GetDataPresent([System.Windows.DataFormats]::FileDrop)) { [System.Windows.DragDropEffects]::Copy } else { [System.Windows.DragDropEffects]::None }
    $e.Handled = $true
}

function Init-ReimbPanel($win) {
    foreach ($n in 'RbList', 'RbFilter', 'RbSum', 'RbViewPanel', 'RbSide', 'RbQuickHost', 'ReimbPanel') { Set-Variable -Name $n -Value $win.FindName($n) -Scope Script }
    $script:RbBtns = @{}
    foreach ($n in 'RbNew', 'RbLedger', 'RbSet', 'RbImport', 'RbTabList', 'RbTabOA', 'RbExcel', 'RbPrint', 'RbFolder') { $script:RbBtns[$n] = $win.FindName($n) }
    $script:RbBtns['RbNew'].add_Click({ Flush-CalFocus; $b = New-Bill; $script:ReimbSel = $b.id; $script:ReimbEdit = ''; $script:ReimbFilter = ''; Render })
    $script:RbBtns['RbLedger'].add_Click({ Export-LedgerXlsx })
    $script:RbBtns['RbSet'].add_Click({ Show-ReimbSettings })
    $script:RbBtns['RbImport'].add_Click({ Pick-ReimbFiles })
    $script:RbBtns['RbTabList'].add_Click({ $script:ReimbView = 'list'; Render-Reimb })
    $script:RbBtns['RbTabOA'].add_Click({ $script:ReimbView = 'oa'; Render-Reimb })
    $script:RbBtns['RbExcel'].add_Click({ Export-BillXlsx $script:ReimbSel })
    $script:RbBtns['RbPrint'].add_Click({
        param($s, $e)
        $m = [System.Windows.Controls.ContextMenu]::new()
        [void]$m.Items.Add((New-MenuItem '粘贴单 + 全部发票（每张一页 A4）' 'all' { param($s2, $e2) Print-Bill $script:ReimbSel ([string]$s2.Tag) }))
        [void]$m.Items.Add((New-MenuItem '只要发票' 'inv' { param($s2, $e2) Print-Bill $script:ReimbSel ([string]$s2.Tag) }))
        [void]$m.Items.Add((New-MenuItem '只要粘贴单' 'sheet' { param($s2, $e2) Print-Bill $script:ReimbSel ([string]$s2.Tag) }))
        $m.PlacementTarget = $s; $m.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom; $m.IsOpen = $true
    })
    $script:RbBtns['RbFolder'].add_Click({
        $b = Find-Bill $script:ReimbSel; if (-not $b) { return }
        $d = Join-Path (Get-ReimbDir) (Get-SafeName $b.no)
        if (-not (Test-Path -LiteralPath $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
        Start-Process explorer.exe ('"' + $d + '"')
    })
    $box = New-HintBox 'rbquick' '记一笔没发票的或者手写的，回车保存，例如：26日7点 XX酒店至会展中心 27.6' $false
    $box[1].add_TextChanged($OnSubAddText)
    $box[1].add_KeyDown({
        param($s, $e)
        if ($e.Key -ne [System.Windows.Input.Key]::Return -or -not $s.Text.Trim()) { return }
        $e.Handled = $true
        $t = $s.Text; $s.Text = ''
        Add-QuickExpense $t
        [void]$s.Focus()
    })
    [void]$script:RbQuickHost.Children.Add($box[0])
    $script:ReimbPanel.add_PreviewDragOver($OnReimbDragOver)
    $script:ReimbPanel.add_PreviewDrop($OnReimbDrop)
}

function Copy-MissingReimbFiles([string]$from, [string]$to) {
    $src = Join-Path $from '报销附件'
    if (-not (Test-Path -LiteralPath $src)) { return }
    foreach ($f in Get-ChildItem -LiteralPath $src -File -Recurse) {
        $rel = $f.FullName.Substring($src.Length).TrimStart('\', '/')
        $target = Join-Path (Join-Path $to '报销附件') $rel
        if (Test-Path -LiteralPath $target) { continue }
        $td = Split-Path -Parent $target
        if (-not (Test-Path -LiteralPath $td)) { New-Item -ItemType Directory -Force -Path $td | Out-Null }
        Copy-Item -LiteralPath $f.FullName $target
    }
}

# ---------- 快捷方式：桌面、开始菜单、开机启动 ----------
$DesktopLnk = Join-Path ([Environment]::GetFolderPath('Desktop')) '桌面待办.lnk'
$MenuLnk = Join-Path ([Environment]::GetFolderPath('Programs')) '桌面待办.lnk'
function New-AppShortcut([string]$lnk) {
    $ws = New-Object -ComObject WScript.Shell
    $sc = $ws.CreateShortcut($lnk)
    $sc.TargetPath = Join-Path $env:WINDIR 'System32\wscript.exe'
    $sc.Arguments = '"' + (Join-Path $AppDir 'DeskTodo.vbs') + '"'
    $sc.WorkingDirectory = $AppDir
    if ($script:IcoPath) { $sc.IconLocation = $script:IcoPath + ',0' }
    $sc.Description = '桌面待办：随手记待办和工作记录（Ctrl+Alt+N 呼出）'
    $sc.Save()
}
function Ensure-Shortcuts {
    try {
        if (-not $script:S.lnkMade) {
            New-AppShortcut $DesktopLnk
            try { New-AppShortcut $MenuLnk } catch {}
            $script:S.lnkMade = $true; Save-Settings
            Notify '桌面上放好了「桌面待办」图标' '以后点桌面上这个蓝色对勾图标就能打开，开始菜单里也有。已经开着时点它，会直接把窗口叫出来。'
        } else {
            # 已经有的快捷方式顺手刷新一下：程序挪了位置、换了图标都能跟上
            foreach ($l in @($DesktopLnk, $MenuLnk, $StartupLnk)) { if (Test-Path -LiteralPath $l) { New-AppShortcut $l } }
        }
    } catch { Write-Log ('建快捷方式失败：' + $_.Exception.Message) }
}
function Get-AutoStart { return (Test-Path $StartupLnk) }
function Set-AutoStart([bool]$on) {
    if ($on) {
        New-AppShortcut $StartupLnk
    } else {
        Remove-Item -LiteralPath $StartupLnk -Force -ErrorAction SilentlyContinue
    }
}


# ---------- 同步状态：每台电脑在同步文件夹里留一个记号 ----------
function Get-DevDir { return (Join-Path $script:ShareDir 'devices') }
function Test-SyncOn { return ([bool]$script:S.shareDir -and $script:ShareDir -ne $DataDir) }

function Write-DeviceFile([bool]$force = $false) {
    if (-not (Test-SyncOn)) { return }
    if (-not $force -and -not $script:DevDirty) { return }
    if (-not $force -and ((Get-Date) - $script:DevWritten).TotalSeconds -lt 20) { return }
    try {
        $dir = Get-DevDir
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        $obj = [ordered]@{ id = (Get-DeviceId); label = (Get-DeviceLabel); version = $AppVersion; seen = [DateTime]::UtcNow.Ticks; changes = $script:MyChange }
        $tmp = Join-Path $DataDir 'device.writing'
        [IO.File]::WriteAllText($tmp, (ConvertTo-Json -InputObject $obj -Depth 4), $Utf8)
        $path = Join-Path $dir ((Get-DeviceId) + '.json')
        Invoke-Retry { [IO.File]::Copy($tmp, $path, $true) } 4 300
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
        $script:DevDirty = $false
        $script:DevWritten = Get-Date
    } catch { Write-Log ('写同步记号失败：' + $_.Exception.Message) }
}

function Read-OtherDevices {
    $list = @()
    if (-not (Test-SyncOn)) { return $list }
    $dir = Get-DevDir
    if (-not (Test-Path -LiteralPath $dir)) { return $list }
    $me = Get-DeviceId
    foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter '*.json' -File -ErrorAction SilentlyContinue)) {
        try {
            $o = ConvertFrom-Json (Read-SharedText $f.FullName)
            if (-not $o.id -or [string]$o.id -eq $me) { continue }
            $list += $o
        } catch {}
    }
    return $list
}

function Format-Ticks($t) {
    if (-not $t) { return '不知道' }
    try { return (Format-When ([datetime]::new([long]$t, [DateTimeKind]::Utc).ToLocalTime())) } catch { return '不知道' }
}

# level：off 没开同步 / ok 正常 / wait 对方有改动还没到 / bad 有问题
function Get-SyncState {
    $lines = New-Object System.Collections.Generic.List[string]
    if (-not $script:S.shareDir) {
        $lines.Add('没开同步：数据只存在这台电脑上。点下面「选同步文件夹」可以和别的电脑同步。')
        return @{ level = 'off'; lines = $lines }
    }
    $level = 'ok'
    if ($script:ShareDir -eq $DataDir) {
        $level = 'bad'
        $lines.Add('找不到同步文件夹：' + $script:S.shareDir + '。现在先存在本机，文件夹找回来后重启程序，会自动合并。')
    } else { $lines.Add('同步文件夹：' + $script:ShareDir) }
    if ($script:PendingWrite.Count -gt 0) { $level = 'bad'; $lines.Add('有改动暂时写不进同步文件夹（文件被占用），程序会自动重试。') }
    $app = Get-SyncAppInfo
    if ($app.exe) {
        if ($app.running) { $lines.Add('这台电脑的' + $app.name + '：正在运行。') }
        else { $level = 'bad'; $lines.Add('这台电脑的' + $app.name + '没在运行：这台的改动传不出去，别的电脑的改动也收不进来。') }
    } else { $lines.Add('还没告诉我用的是哪个同步软件，没法检查它有没有在运行（在「同步设置」里选一下）。') }
    $devs = @(Read-OtherDevices)
    if ($devs.Count -eq 0) { $lines.Add('还没看到别的电脑。另一台电脑也装上桌面待办、选同一个同步文件夹后，这里会显示它的情况。') }
    foreach ($d in $devs) {
        $name = if ($d.label) { [string]$d.label } else { '另一台电脑' }
        $chg = ConvertTo-Hash $d.changes
        $late = 0
        foreach ($f in @($chg.Keys)) {
            $got = $script:Recv[[string]$d.id + '|' + $f]
            if ([long]$chg[$f] -gt [long]$got) { $late = [math]::Max($late, [long]$chg[$f]) }
        }
        $ver = if ($d.version -and [string]$d.version -ne $AppVersion) { '（它装的是 ' + $d.version + ' 版，这台是 ' + $AppVersion + '）' } else { '' }
        if ($late -gt 0) {
            if ($level -eq 'ok') { $level = 'wait' }
            $lines.Add($name + '：' + (Format-Ticks $late) + ' 的改动还没传到这台。要等它那边的网盘传上去、这台的网盘拉下来，一般几分钟。' + $ver)
        } else {
            $lines.Add($name + '：最近一次打开是 ' + (Format-Ticks $d.seen) + '，它的改动这台都收到了。' + $ver)
        }
    }
    return @{ level = $level; lines = $lines }
}

$SyncColors = @{ ok = '#2E9E5B'; wait = '#D9822B'; bad = '#D0402F' }
function Update-SyncBadge {
    try {
        $st = Get-SyncState
        $script:SyncLevel = $st.level
        if ($st.level -eq 'off') { $SyncDot.Visibility = 'Collapsed' }
        else { $SyncDot.Visibility = 'Visible'; $SyncDot.Foreground = Brush $SyncColors[$st.level] }
        $head = switch ($st.level) { 'ok' { '同步正常' } 'wait' { '同步中：另一台电脑有改动还没到' } 'bad' { '同步有问题' } default { '没开同步' } }
        $SyncDot.ToolTip = $head + "`n" + ($st.lines -join "`n") + "`n点这里打开同步设置"
    } catch { Write-Log ('同步状态出错：' + $_.Exception.Message) }
}

# ---------- 同步软件：可以跟着桌面待办一起启动 ----------
$SyncAppKnown = @(
    @{ name = '百度网盘'; proc = 'BaiduNetdisk'; key = 'baidu'; paths = @('%APPDATA%\baidu\BaiduNetdisk\BaiduNetdisk.exe', '%ProgramFiles(x86)%\baidu\BaiduNetdisk\BaiduNetdisk.exe', '%ProgramFiles%\baidu\BaiduNetdisk\BaiduNetdisk.exe') },
    @{ name = '坚果云'; proc = 'Nutstore'; key = 'nutstore|坚果云'; paths = @('%ProgramFiles%\Nutstore\Nutstore.exe', '%ProgramFiles(x86)%\Nutstore\Nutstore.exe') },
    @{ name = 'OneDrive'; proc = 'OneDrive'; key = 'onedrive'; paths = @('%LOCALAPPDATA%\Microsoft\OneDrive\OneDrive.exe', '%ProgramFiles%\Microsoft OneDrive\OneDrive.exe') }
)
function Get-SyncAppName([string]$exe) {
    if (-not $exe) { return '' }
    $n = [IO.Path]::GetFileNameWithoutExtension($exe)
    foreach ($k in $SyncAppKnown) { if ($n -eq $k.proc) { return $k.name } }
    return $n
}
function Find-SyncApp {
    $order = @($SyncAppKnown | Sort-Object @{ Expression = { if ([string]$script:ShareDir -match $_.key) { 0 } else { 1 } } })
    foreach ($k in $order) {
        foreach ($p in @(Get-Process -Name $k.proc -ErrorAction SilentlyContinue)) { try { if ($p.Path) { return $p.Path } } catch {} }
    }
    foreach ($k in $order) {
        foreach ($p in $k.paths) { $x = [Environment]::ExpandEnvironmentVariables($p); if (Test-Path -LiteralPath $x) { return $x } }
    }
    return ''
}
function Get-SyncAppInfo {
    $exe = [string]$script:S.syncApp
    $r = @{ exe = $exe; name = (Get-SyncAppName $exe); running = $false }
    if ($exe) { $r.running = @(Get-Process -Name ([IO.Path]::GetFileNameWithoutExtension($exe)) -ErrorAction SilentlyContinue).Count -gt 0 }
    return $r
}
# 优先用它自己写在开机启动里的那条命令（通常带"静默启动"参数，不会弹主窗口）
function Get-RunCommand([string]$exe) {
    $name = [IO.Path]::GetFileName($exe)
    foreach ($key in 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Run', 'HKLM:\Software\Microsoft\Windows\CurrentVersion\Run', 'HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Run') {
        try {
            $props = Get-ItemProperty -Path $key -ErrorAction Stop
            foreach ($p in $props.PSObject.Properties) {
                if ($p.Name -like 'PS*') { continue }
                $v = [string]$p.Value
                if ($v -and $v.ToLower().Contains($name.ToLower())) { return $v }
            }
        } catch {}
    }
    return ''
}
function Start-SyncApp {
    $exe = [string]$script:S.syncApp
    if (-not $exe -or -not (Test-Path -LiteralPath $exe)) { return $false }
    try {
        $cmd = Get-RunCommand $exe
        $args0 = ''
        if ($cmd) {
            if ($cmd.StartsWith('"')) { $i = $cmd.IndexOf('"', 1); if ($i -gt 0) { $args0 = $cmd.Substring($i + 1).Trim() } }
            else { $i = $cmd.ToLower().IndexOf('.exe'); if ($i -gt 0) { $args0 = $cmd.Substring($i + 4).Trim() } }
        } elseif ([IO.Path]::GetFileNameWithoutExtension($exe) -eq 'OneDrive') { $args0 = '/background' }
        if ($args0) { Start-Process -FilePath $exe -ArgumentList $args0 -WorkingDirectory (Split-Path $exe) }
        else { Start-Process -FilePath $exe -WorkingDirectory (Split-Path $exe) }
        Write-Log ('已启动同步软件：' + $exe + ' ' + $args0)
        return $true
    } catch { Write-Log ('启动同步软件失败：' + $_.Exception.Message); return $false }
}
$script:SyncAppWarned = $false
function Check-SyncApp {
    if (-not (Test-SyncOn)) { return }
    $app = Get-SyncAppInfo
    if (-not $app.exe -or $app.running) { $script:SyncAppWarned = $false; return }
    if ($script:S.syncAppMode -eq 'together') { [void](Start-SyncApp) }
    elseif (-not $script:SyncAppWarned) {
        $script:SyncAppWarned = $true
        Notify ($app.name + '没在运行') ('它不运行，这台电脑的改动传不到别的电脑。打开' + $app.name + '就好；也可以在同步设置里改成「跟着桌面待办一起启动」。')
    }
}

# ---------- 同步设置窗口 ----------
function Show-SyncDialog([bool]$first = $false) {
    if (-not $window.IsVisible) { Show-Main }
    if (-not $script:S.syncApp -and $script:S.shareDir) { $script:S.syncApp = Find-SyncApp }
    $script:Dlg = New-DialogWindow '多台电脑同步'
    $sp = [System.Windows.Controls.StackPanel]::new()
    $sp.Margin = Th 16 14 16 14
    $sp.Width = 440
    function Add-Head([string]$t) {
        $h = [System.Windows.Controls.TextBlock]::new()
        $h.Text = $t; $h.FontWeight = [System.Windows.FontWeights]::SemiBold; $h.Margin = Th 0 10 0 4
        [void]$sp.Children.Add($h)
    }
    function Add-Note([string]$t) {
        $n = [System.Windows.Controls.TextBlock]::new()
        $n.Text = $t; $n.TextWrapping = 'Wrap'; $n.FontSize = 12; $n.Foreground = Brush '#666666'; $n.Margin = Th 0 0 0 4
        [void]$sp.Children.Add($n)
    }
    if ($first) { Add-Note '同步要靠电脑上的同步软件（百度网盘、坚果云、OneDrive 等）一直开着。先定一下它怎么启动，以后可以在托盘菜单「同步设置」里改。' }

    Add-Head '现在的情况'
    $script:SyncStateBox = [System.Windows.Controls.TextBlock]::new()
    $script:SyncStateBox.TextWrapping = 'Wrap'; $script:SyncStateBox.FontSize = 12; $script:SyncStateBox.LineHeight = 19
    $st = Get-SyncState
    $script:SyncStateBox.Text = ($st.lines -join "`n")
    if ($st.level -ne 'off') { $script:SyncStateBox.Foreground = Brush $(if ($st.level -eq 'ok') { '#333333' } else { $SyncColors[$st.level] }) }
    [void]$sp.Children.Add($script:SyncStateBox)

    $fb = [System.Windows.Controls.StackPanel]::new()
    $fb.Orientation = 'Horizontal'; $fb.Margin = Th 0 6 0 0
    $b1 = [System.Windows.Controls.Button]::new()
    $b1.Content = $(if ($script:S.shareDir) { '换同步文件夹…' } else { '选同步文件夹…' }); $b1.Padding = Th 10 3 10 3; $b1.Margin = Th 0 0 8 0
    $b1.add_Click({ $script:Dlg.Close(); Choose-ShareDir })
    [void]$fb.Children.Add($b1)
    if ($script:S.shareDir) {
        $b2 = [System.Windows.Controls.Button]::new()
        $b2.Content = '停止同步，只存本机'; $b2.Padding = Th 10 3 10 3
        $b2.add_Click({
            $r = [System.Windows.MessageBox]::Show('停止同步后，数据改回只存在这台电脑上。同步文件夹里的文件会保留。确定吗？', '桌面待办', 'YesNo', 'Question')
            if ($r -eq 'Yes') { $script:Dlg.Close(); Set-ShareDir $DataDir }
        })
        [void]$fb.Children.Add($b2)
    }
    [void]$sp.Children.Add($fb)

    if ($script:S.shareDir) {
        Add-Head '这台电脑叫'
        $script:DevLabelBox = [System.Windows.Controls.TextBox]::new()
        $script:DevLabelBox.Text = Get-DeviceLabel; $script:DevLabelBox.Padding = Th 4 3 4 3
        [void]$sp.Children.Add($script:DevLabelBox)
        Add-Note '比如「单位电脑」「家里电脑」，另一台电脑显示同步情况时用这个名字。'

        Add-Head '同步软件'
        $ar = [System.Windows.Controls.DockPanel]::new()
        $pick = [System.Windows.Controls.Button]::new()
        $pick.Content = '选择…'; $pick.Padding = Th 10 3 10 3; $pick.Margin = Th 8 0 0 0
        [System.Windows.Controls.DockPanel]::SetDock($pick, 'Right')
        $script:SyncAppBox = [System.Windows.Controls.TextBlock]::new()
        $script:SyncAppBox.VerticalAlignment = 'Center'; $script:SyncAppBox.TextTrimming = 'CharacterEllipsis'
        $script:SyncAppBox.Tag = [string]$script:S.syncApp
        $script:SyncAppBox.Text = if ($script:S.syncApp) { (Get-SyncAppName $script:S.syncApp) + '：' + $script:S.syncApp } else { '没找到，点右边选它的程序（.exe）' }
        $script:SyncAppBox.ToolTip = $script:SyncAppBox.Text
        $pick.add_Click({
            $ofd = New-Object Microsoft.Win32.OpenFileDialog
            $ofd.Title = '选同步软件的程序'; $ofd.Filter = '程序 (*.exe)|*.exe'
            if ($script:SyncAppBox.Tag) { try { $ofd.InitialDirectory = Split-Path ([string]$script:SyncAppBox.Tag) } catch {} }
            if ($ofd.ShowDialog()) { $script:SyncAppBox.Tag = $ofd.FileName; $script:SyncAppBox.Text = (Get-SyncAppName $ofd.FileName) + '：' + $ofd.FileName }
        })
        [void]$ar.Children.Add($pick); [void]$ar.Children.Add($script:SyncAppBox)
        [void]$sp.Children.Add($ar)

        $script:RbTogether = [System.Windows.Controls.RadioButton]::new()
        $script:RbTogether.Content = '跟着桌面待办一起启动（推荐）'; $script:RbTogether.Margin = Th 0 8 0 0
        [void]$sp.Children.Add($script:RbTogether)
        Add-Note '打开桌面待办时，同步软件没开就帮你打开；桌面待办开着时它被关掉了，过几分钟也会再打开。'
        $script:RbSeparate = [System.Windows.Controls.RadioButton]::new()
        $script:RbSeparate.Content = '各自启动，我自己管'; $script:RbSeparate.Margin = Th 0 4 0 0
        [void]$sp.Children.Add($script:RbSeparate)
        Add-Note '桌面待办不去动它，只在它没运行时提醒一次，标题栏的同步图标变红。'
        if ($script:S.syncAppMode -eq 'separate') { $script:RbSeparate.IsChecked = $true } else { $script:RbTogether.IsChecked = $true }

        Add-Head '程序更新'
        $script:AutoUpCb = [System.Windows.Controls.CheckBox]::new()
        $script:AutoUpCb.Content = '另一台电脑更新了程序，这台打开时自动跟着更新'
        $script:AutoUpCb.IsChecked = ($script:S.autoUpdate -ne $false)
        [void]$sp.Children.Add($script:AutoUpCb)
    }

    $ok = [System.Windows.Controls.Button]::new()
    $ok.Content = '保存'; $ok.Width = 72; $ok.Height = 28; $ok.IsDefault = $true
    $ok.HorizontalAlignment = 'Right'; $ok.Margin = Th 0 14 0 0
    $ok.add_Click({
        if ($script:S.shareDir) {
            $lb = $script:DevLabelBox.Text.Trim(); if ($lb) { $script:S.deviceLabel = $lb }
            $script:S.syncApp = [string]$script:SyncAppBox.Tag
            $script:S.syncAppMode = if ($script:RbSeparate.IsChecked) { 'separate' } else { 'together' }
            $script:S.autoUpdate = [bool]$script:AutoUpCb.IsChecked
            Save-Settings
            Write-DeviceFile $true
            Check-SyncApp
        }
        $script:Dlg.Close()
    })
    [void]$sp.Children.Add($ok)
    $script:Dlg.Content = $sp
    [void]$script:Dlg.ShowDialog()
    Update-SyncBadge
}

# ---------- 一键更新 ----------
function Get-DownloadsDir {
    try {
        $v = (Get-ItemProperty 'HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders' -ErrorAction Stop).'{374DE290-123F-4565-9164-39C4925E467B}'
        if ($v) { $x = [Environment]::ExpandEnvironmentVariables([string]$v); if (Test-Path -LiteralPath $x) { return $x } }
    } catch {}
    return (Join-Path $env:USERPROFILE 'Downloads')
}
function Get-ScriptVersion([string]$text) {
    $m = [regex]::Match($text, "\`$AppVersion\s*=\s*'([0-9\.]+)'")
    if ($m.Success) { return $m.Groups[1].Value }
    return '0'
}
function Test-NewerVersion([string]$a, [string]$b) {
    try { return ([version]$a -gt [version]$b) } catch { return $false }
}
# 读一个更新包：.zip 或单独的 DeskTodo.ps1。返回 @{ ver; files = @{ 文件名 = 字节 } }
function Read-UpdatePackage([string]$path) {
    $files = @{}
    if ($path.ToLower().EndsWith('.zip')) {
        Add-Type -AssemblyName System.IO.Compression, System.IO.Compression.FileSystem
        $z = [IO.Compression.ZipFile]::OpenRead($path)
        try {
            foreach ($en in $z.Entries) {
                if ($en.Name -eq 'DeskTodo.ps1' -or $en.Name -eq 'DeskTodo.vbs' -or $en.Name.EndsWith('.docx')) {
                    $ms = New-Object IO.MemoryStream
                    $es = $en.Open(); try { $es.CopyTo($ms) } finally { $es.Dispose() }
                    $files[$en.Name] = $ms.ToArray()
                }
            }
        } finally { $z.Dispose() }
    } else {
        $files['DeskTodo.ps1'] = [IO.File]::ReadAllBytes($path)
    }
    if (-not $files.ContainsKey('DeskTodo.ps1')) { return $null }
    $ver = Get-ScriptVersion ($Utf8.GetString($files['DeskTodo.ps1']))
    return @{ ver = $ver; files = $files }
}
function Get-Sha([byte[]]$b) {
    $h = [Security.Cryptography.SHA256]::Create()
    try { return ([BitConverter]::ToString($h.ComputeHash($b))).Replace('-', '') } finally { $h.Dispose() }
}
function Get-UpdDir { return (Join-Path $script:ShareDir '程序更新') }
# 把一个版本放进同步文件夹，另一台电脑打开时就会自己更新
function Publish-SharedUpdate($pkg) {
    if (-not (Test-SyncOn)) { return }
    try {
        $dir = Get-UpdDir
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Force -Path $dir | Out-Null }
        foreach ($n in $pkg.files.Keys) { [IO.File]::WriteAllBytes((Join-Path $dir $n), $pkg.files[$n]) }
        $man = [ordered]@{ version = $pkg.ver; sha = (Get-Sha $pkg.files['DeskTodo.ps1']); by = (Get-DeviceLabel); at = (Get-Date).ToString($Fmt) }
        [IO.File]::WriteAllText((Join-Path $dir 'version.json'), (ConvertTo-Json $man), $Utf8)
    } catch { Write-Log ('放更新包到同步文件夹失败：' + $_.Exception.Message) }
}
function Publish-SelfIfNewer {
    if (-not (Test-SyncOn)) { return }
    $sv = Get-SharedUpdateVersion
    if ($sv -and -not (Test-NewerVersion $AppVersion $sv)) { return }
    $files = @{ 'DeskTodo.ps1' = [IO.File]::ReadAllBytes($AppFile) }
    foreach ($n in @(Get-ChildItem -LiteralPath $AppDir -File -ErrorAction SilentlyContinue | Where-Object { $_.Name -eq 'DeskTodo.vbs' -or $_.Name.EndsWith('.docx') })) { $files[$n.Name] = [IO.File]::ReadAllBytes($n.FullName) }
    Publish-SharedUpdate @{ ver = $AppVersion; files = $files }
}
function Get-SharedUpdateVersion {
    try {
        $mf = Join-Path (Get-UpdDir) 'version.json'
        if (-not (Test-Path -LiteralPath $mf)) { return '' }
        return [string](ConvertFrom-Json (Read-SharedText $mf)).version
    } catch { return '' }
}
# 同步文件夹里有更新、而且已经完整传到了（核对过指纹），返回更新包
function Get-SharedUpdate {
    if (-not (Test-SyncOn)) { return $null }
    try {
        $dir = Get-UpdDir
        $mf = Join-Path $dir 'version.json'; $pf = Join-Path $dir 'DeskTodo.ps1'
        if (-not (Test-Path -LiteralPath $mf) -or -not (Test-Path -LiteralPath $pf)) { return $null }
        $man = ConvertFrom-Json (Read-SharedText $mf)
        if (-not (Test-NewerVersion ([string]$man.version) $AppVersion)) { return $null }
        $pkg = Read-UpdatePackage $pf
        if (-not $pkg -or (Get-Sha $pkg.files['DeskTodo.ps1']) -ne [string]$man.sha) { return $null }   # 还没传完整
        foreach ($n in 'DeskTodo.vbs') { $x = Join-Path $dir $n; if (Test-Path -LiteralPath $x) { $pkg.files[$n] = [IO.File]::ReadAllBytes($x) } }
        foreach ($f in @(Get-ChildItem -LiteralPath $dir -Filter '*.docx' -File -ErrorAction SilentlyContinue)) { $pkg.files[$f.Name] = [IO.File]::ReadAllBytes($f.FullName) }
        if ($pkg.ver -ne [string]$man.version) { return $null }   # 版本号对不上，不装，免得反复更新
        $pkg.from = [string]$man.by
        return $pkg
    } catch { return $null }
}

function Install-Update($pkg, [bool]$publish = $true) {
    try {
        $bak = Join-Path $DataDir 'backup'
        if (-not (Test-Path -LiteralPath $bak)) { New-Item -ItemType Directory -Force -Path $bak | Out-Null }
        Copy-Item -LiteralPath $AppFile (Join-Path $bak ('DeskTodo-' + $AppVersion + '.ps1')) -Force
        foreach ($n in $pkg.files.Keys) { [IO.File]::WriteAllBytes((Join-Path $AppDir $n), $pkg.files[$n]) }
    } catch {
        Write-Log ('更新失败：' + $_.Exception.ToString())
        [System.Windows.MessageBox]::Show('更新没装上：' + $_.Exception.Message + "`n`n程序所在的文件夹可能不允许写入。原来的版本没动，可以继续用。", '桌面待办') | Out-Null
        return $false
    }
    if ($publish) { Publish-SharedUpdate $pkg }
    $script:S.justUpdated = $pkg.ver
    Save-Settings
    Restart-App
    return $true
}

function Restart-App {
    # 先退出，过两秒再由一个小脚本重新打开（等这边彻底关掉，热键和"只运行一个"的锁才会放开）
    $vbs = Join-Path $AppDir 'DeskTodo.vbs'
    $tmp = Join-Path $env:TEMP 'DeskTodo_restart.vbs'
    $code = 'WScript.Sleep 2500' + "`r`n" + 'CreateObject("WScript.Shell").Run "wscript.exe ""' + $vbs + '""", 0, False' + "`r`n"
    [IO.File]::WriteAllText($tmp, $code, [Text.Encoding]::Unicode)
    Start-Process -FilePath 'wscript.exe' -ArgumentList ('"' + $tmp + '"')
    Exit-App
}

function Invoke-UpdateFromFile {
    $path = ''
    $dl = Get-DownloadsDir
    $z = @(Get-ChildItem -LiteralPath $dl -Filter 'DeskTodo*.zip' -File -ErrorAction SilentlyContinue | Sort-Object LastWriteTime -Descending | Select-Object -First 1)
    if ($z.Count -gt 0) {
        $r = [System.Windows.MessageBox]::Show(("在「下载」里找到：{0}（{1} 下载的）`n`n用它更新吗？选「否」可以自己选别的文件。" -f $z[0].Name, (Format-When $z[0].LastWriteTime)), '更新桌面待办', 'YesNoCancel', 'Question')
        if ($r -eq 'Cancel') { return }
        if ($r -eq 'Yes') { $path = $z[0].FullName }
    }
    if (-not $path) {
        $ofd = New-Object Microsoft.Win32.OpenFileDialog
        $ofd.Title = '选新版本的压缩包（DeskTodo.zip）或 DeskTodo.ps1'
        $ofd.Filter = '更新包 (*.zip;*.ps1)|*.zip;*.ps1'
        $ofd.InitialDirectory = $dl
        if (-not $ofd.ShowDialog()) { return }
        $path = $ofd.FileName
    }
    $pkg = $null
    try { $pkg = Read-UpdatePackage $path } catch { Write-Log ('读更新包失败：' + $_.Exception.Message) }
    if (-not $pkg) { [System.Windows.MessageBox]::Show('这个文件里没找到 DeskTodo.ps1，不是桌面待办的更新包。', '桌面待办') | Out-Null; return }
    if (-not (Test-NewerVersion $pkg.ver $AppVersion)) {
        $r = [System.Windows.MessageBox]::Show(("这个包的版本是 {0}，不比现在用的 {1} 新。还要装吗？" -f $pkg.ver, $AppVersion), '更新桌面待办', 'YesNo', 'Question')
        if ($r -ne 'Yes') { return }
    }
    [void](Install-Update $pkg)
}

# 打开时先看同步文件夹里有没有另一台电脑放的新版本：有就直接换上并重启（窗口还没出来，几乎感觉不到）
function Check-SharedUpdateAtStart {
    if ($script:S.autoUpdate -eq $false) { return $false }
    $pkg = Get-SharedUpdate
    if (-not $pkg) { return $false }
    Write-Log ('同步文件夹里有新版本 ' + $pkg.ver + '，自动更新')
    try {
        $bak = Join-Path $DataDir 'backup'
        if (-not (Test-Path -LiteralPath $bak)) { New-Item -ItemType Directory -Force -Path $bak | Out-Null }
        Copy-Item -LiteralPath $AppFile (Join-Path $bak ('DeskTodo-' + $AppVersion + '.ps1')) -Force
        foreach ($n in $pkg.files.Keys) { [IO.File]::WriteAllBytes((Join-Path $AppDir $n), $pkg.files[$n]) }
        $script:S.justUpdated = $pkg.ver
        Save-Settings
        return $true
    } catch { Write-Log ('自动更新失败：' + $_.Exception.Message); return $false }
}
$script:SharedUpdNotified = ''
function Check-SharedUpdateRunning {
    if ($script:S.autoUpdate -eq $false) { return }
    $pkg = Get-SharedUpdate
    if (-not $pkg -or $script:SharedUpdNotified -eq $pkg.ver) { return }
    $script:SharedUpdNotified = $pkg.ver
    $script:SharedPkg = $pkg
    Notify ('有新版本 ' + $pkg.ver) ($(if ($pkg.from) { $pkg.from } else { '另一台电脑' }) + '已经更新了。托盘菜单里点「更新到 ' + $pkg.ver + '」，或者下次打开时自动更新。')
}

# ---------- 托盘菜单 ----------
$trayMenu = [System.Windows.Forms.ContextMenuStrip]::new()
$miShow = $trayMenu.Items.Add('显示 / 隐藏（Ctrl+Alt+N）')
$miLog  = $trayMenu.Items.Add('记一条工作记录')
$miWeek = $trayMenu.Items.Add('复制本周周报')
$miCal  = $trayMenu.Items.Add('工作台：周 / 月视图')
$miRep  = $trayMenu.Items.Add('生成报告（日报、周报、月报…）')
$miReimb = $trayMenu.Items.Add('报销（导入发票、OA 填报、打印）')
$miTime = New-Object System.Windows.Forms.ToolStripMenuItem '显示几点几分'
$miAuto = New-Object System.Windows.Forms.ToolStripMenuItem '开机自动启动'
[void]$trayMenu.Items.Add($miTime)
[void]$trayMenu.Items.Add($miAuto)
$miLnk = $trayMenu.Items.Add('在桌面放一个图标')
$miStyle = $trayMenu.Items.Add('外观（透明度、配色）')
$miSync = $trayMenu.Items.Add('多台电脑同步…')
$miDir  = $trayMenu.Items.Add('打开数据文件夹')
[void]$trayMenu.Items.Add(([System.Windows.Forms.ToolStripSeparator]::new()))
$miUpShared = $trayMenu.Items.Add('更新到新版本')
$miUpdate = $trayMenu.Items.Add('更新程序（用下载的新版本）…')
$miVer = $trayMenu.Items.Add('版本 ' + $AppVersion)
$miVer.Enabled = $false
$miExit = $trayMenu.Items.Add('退出')
$script:Tray.ContextMenuStrip = $trayMenu

$trayMenu.add_Opening({
    $miAuto.Checked = Get-AutoStart
    $miTime.Checked = Test-ShowTime
    $miSync.Text = if ($script:S.shareDir) { '同步设置…' } else { '多台电脑同步…' }
    $miUpShared.Visible = [bool]$script:SharedPkg -and (Test-NewerVersion $script:SharedPkg.ver $AppVersion)
    if ($miUpShared.Visible) { $miUpShared.Text = '更新到 ' + $script:SharedPkg.ver + '（另一台电脑已经更新）' }
})
$miShow.add_Click({ Toggle-Main })
$miLog.add_Click({ Show-Main; Set-Mode 'log' })
$miWeek.add_Click({ Copy-WeekSummary })
$miCal.add_Click({ Show-Calendar 'week' })
$miRep.add_Click({ Show-Calendar 'report' })
$miReimb.add_Click({ Show-Calendar 'reimb' })
$miAuto.add_Click({
    $on = -not (Get-AutoStart)
    Set-AutoStart $on
    if ($on) { Notify '已开启开机自动启动' '以后开机会自动打开。注意别再挪动这个程序所在的文件夹。' }
    else { Notify '已关闭开机自动启动' '需要时双击 DeskTodo.vbs 打开。' }
})
$miDir.add_Click({ Start-Process explorer.exe $script:ShareDir })
$miLnk.add_Click({
    try { New-AppShortcut $DesktopLnk; Notify '已放到桌面' '桌面上的「桌面待办」就是它，双击打开。' }
    catch { Notify '没放成' ('建快捷方式失败：' + $_.Exception.Message) }
})
$miSync.add_Click({ Show-Main; if ($script:S.shareDir) { Show-SyncDialog } else { Choose-ShareDir } })
$miUpdate.add_Click({ Invoke-UpdateFromFile })
$miUpShared.add_Click({ if ($script:SharedPkg) { [void](Install-Update $script:SharedPkg $false) } })
$miStyle.add_Click({ Show-StyleDialog })
$miTime.add_Click({ Set-ShowTime (-not (Test-ShowTime)) })
$miExit.add_Click({ Exit-App })
$script:Tray.add_MouseClick({ param($s, $e) if ($e.Button -eq [System.Windows.Forms.MouseButtons]::Left) { Show-Main } })

function Exit-App {
    $script:Exiting = $true
    try { Finish-Undo } catch {}
    try { Save-Items; Flush-Saves } catch {}
    Save-Settings
    try { $script:Timer.Stop(); $script:SyncTimer.Stop() } catch {}
    try { Write-DeviceFile $true } catch {}
    try { if ($script:HotKey -is [DeskTodoHotKey]) { $script:HotKey.Dispose() } } catch {}
    $script:Tray.Visible = $false
    $script:Tray.Dispose()
    $app.Shutdown()
}

# ---------- 事件 ----------
$Header.add_MouseLeftButtonDown({
    param($s, $e)
    if ($e.ClickCount -eq 2) { Set-Collapsed (-not $script:S.collapsed) }
    else { try { $window.DragMove() } catch {} }
})
$BtnPin.add_Click({ Set-Pin (-not $script:S.topmost) })
$UndoBtn.add_MouseLeftButtonUp({ Invoke-Undo })
$BtnMore.add_Click({
    param($s, $e)
    $m = [System.Windows.Controls.ContextMenu]::new()
    $a = [System.Windows.Controls.MenuItem]::new(); $a.Header = '显示几点几分'; $a.IsCheckable = $true; $a.IsChecked = (Test-ShowTime)
    $a.add_Click({ Set-ShowTime (-not (Test-ShowTime)) }); [void]$m.Items.Add($a)
    $b = [System.Windows.Controls.MenuItem]::new(); $b.Header = '窗口置顶'; $b.IsCheckable = $true; $b.IsChecked = [bool]$script:S.topmost
    $b.add_Click({ Set-Pin (-not $script:S.topmost) }); [void]$m.Items.Add($b)
    [void]$m.Items.Add((New-MenuItem '收起窗口（双击标题栏也行）' '' { Set-Collapsed (-not $script:S.collapsed) }))
    [void]$m.Items.Add((New-MenuItem '外观：透明度和配色…' '' { Show-StyleDialog }))
    [void]$m.Items.Add([System.Windows.Controls.Separator]::new())
    [void]$m.Items.Add((New-MenuItem $(if ($script:S.shareDir) { '同步设置…' } else { '多台电脑同步…' }) '' { if ($script:S.shareDir) { Show-SyncDialog } else { Choose-ShareDir } }))
    [void]$m.Items.Add((New-MenuItem '更新程序（用下载的新版本）…' '' { Invoke-UpdateFromFile }))
    $m.PlacementTarget = $s; $m.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Bottom
    $m.IsOpen = $true
})
$SyncDot.add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true })
$SyncDot.add_MouseLeftButtonUp({ param($s, $e) $e.Handled = $true; Show-SyncDialog })
$BtnSync.add_Click({ Show-SyncDialog })
$BtnCollapse.add_Click({ Set-Collapsed (-not $script:S.collapsed) })
$BtnHide.add_Click({ Hide-Main })
$BtnShowDone.add_MouseLeftButtonUp({ $script:S.showDone = -not $script:S.showDone; Save-Settings; Render })
$BtnExport.add_MouseLeftButtonUp({ Copy-WeekSummary })

$InputBox.add_TextChanged({
    $Hint.Visibility = if ($InputBox.Text.Length -gt 0) { 'Collapsed' } else { 'Visible' }
    Update-ParseHint
    Update-ChipsVis
})
$InputBox.add_GotKeyboardFocus({ Update-ChipsVis })
$InputBox.add_LostKeyboardFocus({ Update-ChipsVis })
$InputBox.add_PreviewKeyDown({
    param($s, $e)
    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    if ($script:S.mode -eq 'todo' -and $ctrl -and $e.Key -eq [System.Windows.Input.Key]::V -and (Test-ClipHasOnlyImage)) {
        $imgs = Get-ClipImages
        if ($imgs.Count -eq 0) { return }
        $e.Handled = $true
        if (-not $InputBox.Text.Trim()) { $InputBox.Text = '图片待办（' + (Get-Date).ToString('M月d日 HH:mm') + '）' }
        Add-Todo
        $new = $script:Items[0]
        $script:Open[$new.id] = $true
        [void](Add-SubImages $new.id $imgs)
    }
})
$InputBox.add_KeyDown({
    param($s, $e)
    if ($e.Key -eq [System.Windows.Input.Key]::Return) { $e.Handled = $true; Add-Entry }
    elseif ($e.Key -eq [System.Windows.Input.Key]::Escape) { $e.Handled = $true; Hide-Main }
})

$TabTodo.add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true; if ($script:S.collapsed) { Set-Collapsed $false }; Set-Mode 'todo' })
$TabLog.add_MouseLeftButtonDown({ param($s, $e) $e.Handled = $true; if ($script:S.collapsed) { Set-Collapsed $false }; Set-Mode 'log' })
$BtnPrevDay.add_Click({ $script:ViewDate = $script:ViewDate.AddDays(-1); Render })
$BtnNextDay.add_Click({ $script:ViewDate = $script:ViewDate.AddDays(1); Render })
$DayLabel.add_MouseLeftButtonUp({ $script:ViewDate = (Get-Date).Date; Render })
$BtnImport.add_MouseLeftButtonUp({ Import-DoneTodos })
$BtnCopyLog.add_MouseLeftButtonUp({ Copy-LogToSheet })
$BtnCols.add_MouseLeftButtonUp({
    param($s, $e)
    $m = [System.Windows.Controls.ContextMenu]::new()
    [void]$m.Items.Add((New-MenuItem '表格列设置…' '' { Edit-Columns }))
    [void]$m.Items.Add((New-MenuItem '这一天全部标成「没复制过」' '' {
        $ds = $script:ViewDate.ToString('yyyy-MM-dd')
        foreach ($r0 in (Get-DayRows $ds)) { if ($r0.w) { $r0.w.copied = $false } else { $r0.t.copied = $false } }
        Save-Work; Save-Items; Render }))
    $m.PlacementTarget = $s; $m.Placement = [System.Windows.Controls.Primitives.PlacementMode]::Top
    $m.IsOpen = $true
})
$BtnLogExpand.add_MouseLeftButtonUp({
    $closed = @($script:LogVisible | Where-Object { -not $script:LogOpen.ContainsKey($_) })
    if ($closed.Count -gt 0) { foreach ($id in $script:LogVisible) { $script:LogOpen[$id] = $true } } else { $script:LogOpen = @{} }
    Render
})

$BtnStyle.add_Click({ Show-StyleDialog })
$window.add_MouseEnter({ Update-Opacity })
$window.add_MouseLeave({ Update-Opacity })
$window.add_Activated({ Update-Opacity })
$window.add_Deactivated({ Update-Opacity })

$BtnSearch.add_Click({ if ($SearchBar.Visibility -eq 'Visible') { Close-Search } else { Show-Search } })
$BtnCloseSearch.add_Click({ Close-Search })
$SearchBox.add_TextChanged({
    $SearchHint.Visibility = if ($SearchBox.Text) { 'Collapsed' } else { 'Visible' }
    $script:Query = $SearchBox.Text
    Update-SearchOpen
    Render
})
$SearchBox.add_KeyDown({ param($s, $e) if ($e.Key -eq [System.Windows.Input.Key]::Escape) { $e.Handled = $true; Close-Search } })
$window.add_PreviewKeyDown({
    param($s, $e)
    $ctrl = ([System.Windows.Input.Keyboard]::Modifiers -band [System.Windows.Input.ModifierKeys]::Control) -ne 0
    if ($ctrl -and $e.Key -eq [System.Windows.Input.Key]::F) { $e.Handled = $true; Show-Search }
})
$BtnExpandAll.add_MouseLeftButtonUp({ Toggle-ExpandAll })
$BtnCal.add_Click({ Show-Calendar 'week' })
$BtnTime.add_Click({ Set-ShowTime (-not (Test-ShowTime)) })
$BtnGroup.add_MouseLeftButtonUp({ $script:S.groupProj = -not $script:S.groupProj; Save-Settings; Render })

$window.add_PreviewMouseLeftButtonDown({
    param($s, $e)
    $fe = [System.Windows.Input.Keyboard]::FocusedElement
    if ($fe -is [System.Windows.Controls.TextBox] -and -not $fe.IsReadOnly -and $null -ne $fe.Resources['inlineCancel'] -and -not $fe.IsMouseOver) { [System.Windows.Input.Keyboard]::ClearFocus() }
})

$window.add_Closing({ param($s, $e) if (-not $script:Exiting) { $e.Cancel = $true; Hide-Main } })
$window.add_LocationChanged({ $script:S.left = $window.Left; $script:S.top = $window.Top })
$window.add_SizeChanged({
    if (-not $script:S.collapsed -and $window.SizeToContent -eq [System.Windows.SizeToContent]::Manual -and $window.ActualHeight -ge 200) {
        $script:S.height = $window.ActualHeight
        $script:S.width = $window.ActualWidth
    }
})
$window.add_ContentRendered({
    if ($script:FirstRenderDone) { return }
    $script:FirstRenderDone = $true
    if (-not $script:HotKey.Ok) {
        Notify '快捷键没注册上' 'Ctrl+Alt+N 被别的软件占用了，可以点右下角托盘图标打开。'
    }
    if (-not $script:S.asked) {
        $script:S.asked = $true
        Save-Settings
        $msg = "要开机自动启动吗？`n`n建议先把这个程序文件夹放到固定位置（比如 D:\工具\桌面待办），再选「是」。以后也可以在右下角托盘图标的右键菜单里开关。"
        $r = [System.Windows.MessageBox]::Show($window, $msg, '桌面待办', 'YesNo', 'Question')
        if ($r -eq 'Yes') { Set-AutoStart $true }
    }
    Check-Reminders
})

$app.add_DispatcherUnhandledException({
    param($s, $e)
    Write-Log $e.Exception.ToString()
    $e.Handled = $true
})
$app.add_SessionEnding({ try { Flush-Saves } catch {}; Save-Settings })

$script:Timer = [System.Windows.Threading.DispatcherTimer]::new()
$script:Timer.Interval = [TimeSpan]::FromSeconds(30)
$script:Timer.add_Tick({
    Check-Reminders
    if ((Get-Sig) -ne $script:RenderSig -and -not (Test-Typing)) { Render }
    Save-Settings
})

# ---------- 启动 ----------
try {
    Load-Settings
    Ensure-AppIco
    if ($script:IcoPath) { $script:S.icoVer = $IcoVersion }
    Set-WindowIcon $window
    $script:Recv = ConvertTo-Hash $script:S.recv; $script:S.recv = $script:Recv
    $script:MyChange = ConvertTo-Hash $script:S.myChange; $script:S.myChange = $script:MyChange
    [void](Get-DeviceId)
    $script:TodoStore = New-Store 'todos.json' $script:Items { param($o) ConvertTo-Todo $o }
    $script:WorkStore = New-Store 'worklog.json' $script:Work { param($o) ConvertTo-WorkEntry $o }
    $script:ProjStore = New-Store 'projects.json' $script:Projects { param($o) ConvertTo-Project $o }
    $script:BillStore = New-Store 'bills.json' $script:Bills { param($o) ConvertTo-Bill $o }
    $script:ExpStore = New-Store 'expenses.json' $script:Exps { param($o) ConvertTo-Expense $o }
    $script:StoreList = @($script:TodoStore, $script:WorkStore, $script:ProjStore, $script:BillStore, $script:ExpStore)
    if ($script:S.shareDir) {
        if (Test-Path -LiteralPath ([string]$script:S.shareDir)) { $script:ShareDir = [string]$script:S.shareDir }
        else { $script:ShareMissing = $true }
    }
    if ($null -eq $script:S.quick) { $script:S.quick = @($DefaultQuick) }
    # 另一台电脑更新过程序：直接换成新版本重新打开
    if (Check-SharedUpdateAtStart) {
        $script:Mutex.ReleaseMutex(); $script:Mutex.Dispose()
        $tmp = Join-Path $env:TEMP 'DeskTodo_restart.vbs'
        [IO.File]::WriteAllText($tmp, ('WScript.Sleep 1500' + "`r`n" + 'CreateObject("WScript.Shell").Run "wscript.exe ""' + (Join-Path $AppDir 'DeskTodo.vbs') + '""", 0, False' + "`r`n"), [Text.Encoding]::Unicode)
        Start-Process -FilePath 'wscript.exe' -ArgumentList ('"' + $tmp + '"')
        try { $script:Tray.Visible = $false; $script:Tray.Dispose() } catch {}
        exit
    }
    Load-Items
    Load-Work
    [void](Load-Store $script:ProjStore)
    [void](Load-Store $script:BillStore)
    [void](Load-Store $script:ExpStore)
    Invoke-Migration
    Invoke-SubBackfill

    $wa = [System.Windows.SystemParameters]::WorkArea
    $window.Width = [math]::Max(280, [double]$script:S.width)
    $window.Height = Get-ExpandedHeight
    $l = $script:S.left; $t = $script:S.top
    $vl = [System.Windows.SystemParameters]::VirtualScreenLeft
    $vt = [System.Windows.SystemParameters]::VirtualScreenTop
    $vw = [System.Windows.SystemParameters]::VirtualScreenWidth
    $vh = [System.Windows.SystemParameters]::VirtualScreenHeight
    if ($null -eq $l -or $null -eq $t -or $l -lt $vl -or $t -lt $vt -or $l -gt ($vl + $vw - 100) -or $t -gt ($vt + $vh - 60)) {
        $l = $wa.Right - $window.Width - 16
        $t = $wa.Top + 80
    }
    $window.Left = [double]$l
    $window.Top = [double]$t

    Set-Pin ([bool]$script:S.topmost)
    Set-Collapsed ([bool]$script:S.collapsed)
    Apply-Theme ([string]$script:S.theme)
    Set-Mode ([string]$script:S.mode)
    Update-Opacity

    # Ctrl(2) + Alt(1) + 不重复(0x4000)，N 键 = 0x4E
    try {
        $script:HotKey = [DeskTodoHotKey]::new([uint32]0x4003, [uint32]0x4E)
        $script:HotKey.add_Pressed({ Toggle-Main })
    } catch {
        Write-Log ('热键初始化失败：' + $_.Exception.Message)
        $script:HotKey = [pscustomobject]@{ Ok = $false }
    }

    $script:Timer.Start()
    $script:SyncTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:SyncTimer.Interval = [TimeSpan]::FromSeconds(5)
    $script:SyncTick = 0
    $script:SyncTimer.add_Tick({
        if ($script:ShareDir -ne $DataDir) { Poll-Stores }
        else { foreach ($st in $script:StoreList) { if ($script:PendingWrite.ContainsKey($st.FileName)) { Save-Store $st } } }
        $script:SyncTick++
        Write-DeviceFile ($script:SyncTick % 120 -eq 0)          # 有改动 20 秒内写一次记号，没改动每 10 分钟写一次
        if ($script:SyncTick % 3 -eq 0) { Update-SyncBadge }      # 每 15 秒刷新标题栏的同步图标
        if ($script:SyncTick -eq 12 -or $script:SyncTick % 36 -eq 0) { Check-SyncApp }   # 打开 1 分钟后，之后每 3 分钟
        if ($script:SyncTick % 120 -eq 6) { Check-SharedUpdateRunning }
        if ($script:AskSyncApp) { $script:AskSyncApp = $false; Show-SyncDialog $true }
    })
    $script:SyncTimer.Start()
    $script:ShowTimer = [System.Windows.Threading.DispatcherTimer]::new()
    $script:ShowTimer.Interval = [TimeSpan]::FromMilliseconds(400)
    $script:ShowTimer.add_Tick({ try { if ($script:ShowEvent -and $script:ShowEvent.WaitOne(0)) { Show-Main } } catch {} })
    $script:ShowTimer.Start()
    Ensure-Shortcuts
    Update-SyncBadge
    if (Test-SyncOn) {
        Write-DeviceFile $true
        Publish-SelfIfNewer
        if (-not $script:S.syncApp) { $script:S.syncApp = Find-SyncApp }
        if ($script:S.syncAppMode -eq 'together') { Check-SyncApp }
        elseif (-not $script:S.syncAppMode) { $script:AskSyncApp = $true }
    }
    if ($script:S.justUpdated) { Notify ('已更新到 ' + $script:S.justUpdated) '数据没有变化，接着用就行。'; $script:S.justUpdated = ''; Save-Settings }
    if ($script:ShareMissing) { Notify '找不到同步文件夹' ('先用本机数据。同步文件夹：' + $script:S.shareDir + '。网盘恢复后重启程序即可。') }
    $window.Show()
    [void]$InputBox.Focus()
    [void]$app.Run()
} catch {
    Write-Log $_.Exception.ToString()
    [System.Windows.MessageBox]::Show("桌面待办启动出错：`n" + $_.Exception.Message + "`n`n详细记录在：$LogFile", '桌面待办') | Out-Null
    try { $script:Tray.Visible = $false; $script:Tray.Dispose() } catch {}
}
