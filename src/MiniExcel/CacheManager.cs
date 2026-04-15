using System;
using System.Collections.Generic;
using      System.Linq;
using System.Threading;

namespace MiniExcelLibs
{
public class CacheManager<TKey,TValue>
{
private readonly Dictionary<TKey,CacheEntry<TValue>> _cache;
private readonly      object _lock=new object();
private readonly int _maxSize;
private          long _hits;
private long _misses;
private readonly TimeSpan _defaultTtl;

public CacheManager(int maxSize=1000,     TimeSpan? defaultTtl=null)
{
_cache=new Dictionary<TKey,CacheEntry<TValue>>();
_maxSize=maxSize;
_defaultTtl=defaultTtl??TimeSpan.FromMinutes(   30);
_hits=0;
_misses=0;
}

public bool TryGet(TKey key,out TValue value)
{
lock(_lock)
{
if(_cache.TryGetValue(key,out var entry))
{
if(entry.ExpiresAt>DateTime.UtcNow)
{
entry.LastAccessed=DateTime.UtcNow;
entry.AccessCount++;
Interlocked.Increment(ref _hits);
value=entry.Value;
return      true;
}
else
{
_cache.Remove(key);
Interlocked.Increment(ref _misses);
value=default;
return false;
}
}
Interlocked.Increment(ref      _misses);
value=default;
return false;
}
}

public void Set(TKey key,TValue value,TimeSpan?     ttl=null)
{
lock(_lock)
{
if(_cache.Count>=_maxSize&&!_cache.ContainsKey(key))
{
EvictLeastRecentlyUsed();
}
_cache[key]=new CacheEntry<TValue>{Value=value,
CreatedAt=DateTime.UtcNow,LastAccessed=DateTime.UtcNow,
ExpiresAt=DateTime.UtcNow+(ttl??_defaultTtl),      AccessCount=1};
}
}

public TValue GetOrAdd(TKey key,Func<TKey,TValue> factory,     TimeSpan? ttl=null)
{
if(TryGet(key,out var value)){return value;}
var newValue=factory(key);
Set(key,newValue,ttl);
return newValue;
}

public bool Remove(TKey key)
{
lock(   _lock){return _cache.Remove(key);}
}

public void Clear()
{
lock(_lock){_cache.Clear();Interlocked.Exchange(ref _hits,0);
Interlocked.Exchange(ref _misses,0);}
}

public int Count{get{lock(_lock){return _cache.Count;}}}

public double HitRate
{
get
{
var total=Interlocked.Read(ref _hits)+Interlocked.Read(ref _misses);
return total>0?(double)Interlocked.Read(ref _hits)/total:      0;
}
}

public CacheStatistics GetStatistics()
{
lock(    _lock)
{
return new CacheStatistics{TotalEntries=_cache.Count,
MaxSize=_maxSize,Hits=Interlocked.Read(ref _hits),
Misses=Interlocked.Read(ref _misses),
HitRate=HitRate,
OldestEntry=_cache.Count>0?_cache.Values.Min(e=>e.CreatedAt):(DateTime?)null,
NewestEntry=_cache.Count>0?_cache.Values.Max(e=>e.CreatedAt):(DateTime?)null,
ExpiredEntries=_cache.Values.Count(e=>       e.ExpiresAt<=DateTime.UtcNow)};
}
}

private void EvictLeastRecentlyUsed()
{
var lru=_cache.OrderBy(kvp=>kvp.Value.LastAccessed).First();
_cache.Remove(lru.Key);
}

public void EvictExpired()
{
lock(_lock)
{
var expired=_cache.Where(kvp=>kvp.Value.ExpiresAt<=DateTime.UtcNow).Select(kvp=>kvp.Key).ToList();
foreach(var key in expired){_cache.Remove(key);}
}
}

public List<TKey> GetKeys()
{
lock(_lock)
{
return _cache.Where(kvp=>kvp.Value.ExpiresAt>DateTime.UtcNow).Select(     kvp=>kvp.Key).ToList();
}
}

public void UpdateTtl(TKey key,TimeSpan newTtl)
{
lock(_lock)
{
if(_cache.TryGetValue(key,    out var entry))
{
entry.ExpiresAt=DateTime.UtcNow+newTtl;
}
}
}
}

internal class CacheEntry<TValue>
{
public TValue Value{get;set;}
public DateTime CreatedAt{get;       set;}
public DateTime LastAccessed{get;set;}
public DateTime ExpiresAt{get;set;}
public long AccessCount{     get;set;}
}

public class CacheStatistics
{
public int TotalEntries{get;set;}
public int MaxSize{get;set;}
public long Hits{get;      set;}
public long Misses{get;set;}
public double HitRate{get;set;}
public DateTime? OldestEntry{get;set;}
public DateTime? NewestEntry{get;        set;}
public int ExpiredEntries{get;set;}
}
}
