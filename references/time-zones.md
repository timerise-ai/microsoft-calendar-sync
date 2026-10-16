# Mailbox time zones

Timed events carry their instant. All-day events do not: Outlook stores them as floating dates, and
Graph returns them as midnight to midnight in whatever zone the request asked for. In UTC, a Warsaw
user's vacation on 3 to 7 August arrives as `2026-08-03T00:00:00` to `2026-08-08T00:00:00`, which is
two hours late at both ends. The dates are right; the instants need the zone the user lives in.

## Which zone

`allDayTimeZoneFor` resolves it per user, first match wins:

| Source | When | Notes |
|---|---|---|
| Mailbox zone | `MailboxSettings.Read` granted | Read on connect and on every calendar-list refresh, so a move is picked up |
| `deps.hostTimeZone(userId)` | the host supplies it | Usually the business's zone; ignored unless it is a valid IANA name |
| `config.defaultTimeZone` | always | Required config, so nothing silently falls back to UTC |

The mailbox zone is the user's own setting in Outlook, which is what they mean by "all day". The host's
business zone is right for most staff and wrong by the offset for someone working remotely.

## Windows names

`GET /me/mailboxSettings/timeZone` returns the zone the way Outlook stores it: a Windows name such as
`W. Europe Standard Time` for most mailboxes, an IANA name for some. `Intl` understands only IANA.
`toIanaTimeZone` maps Windows names through CLDR's `windowsZones` table (the default-territory
`001` entries) and passes IANA names through after checking that `Intl` accepts them. Anything else
becomes null, and the next source in the table is used.

The table changes rarely (a new Windows zone every year or two). Regenerate it from
`https://github.com/unicode-org/cldr/blob/main/common/supplemental/windowsZones.xml`, taking each
`mapZone` with `territory="001"`, when CLDR adds one.

## Edge cases the mapper handles

- **A day without a midnight** (DST starting at 00:00, e.g. `America/Santiago`): the block starts at the
  first instant that exists.
- **Days of 23 or 25 hours**: midnight to midnight, computed per date, never `start + 24 h`.
- **An all-day event that arrives at a time other than midnight**: Graph sent real instants after
  all, and they are used as they are.
- **The process `TZ`**: nothing reads it. The suite runs under four zones to keep it that way.

## Code

```ts
// microsoft-calendar/timeZones.ts
/**
 * Windows time-zone names to IANA, the CLDR `windowsZones` "001" (default territory) mapping.
 * Outlook stores a mailbox's zone as a Windows name ("W. Europe Standard Time") unless the user
 * picked an IANA one; `Intl` understands only IANA.
 */
export const WINDOWS_TO_IANA: Readonly<Record<string, string>> = {
  "Dateline Standard Time": "Etc/GMT+12",
  "UTC-11": "Etc/GMT+11",
  "Aleutian Standard Time": "America/Adak",
  "Hawaiian Standard Time": "Pacific/Honolulu",
  "Marquesas Standard Time": "Pacific/Marquesas",
  "Alaskan Standard Time": "America/Anchorage",
  "UTC-09": "Etc/GMT+9",
  "Pacific Standard Time (Mexico)": "America/Tijuana",
  "UTC-08": "Etc/GMT+8",
  "Pacific Standard Time": "America/Los_Angeles",
  "US Mountain Standard Time": "America/Phoenix",
  "Mountain Standard Time (Mexico)": "America/Mazatlan",
  "Mountain Standard Time": "America/Denver",
  "British Columbia Standard Time": "America/Vancouver",
  "Yukon Standard Time": "America/Whitehorse",
  "Alberta Standard Time": "America/Edmonton",
  "Central America Standard Time": "America/Guatemala",
  "Central Standard Time": "America/Chicago",
  "Easter Island Standard Time": "Pacific/Easter",
  "Central Standard Time (Mexico)": "America/Mexico_City",
  "Canada Central Standard Time": "America/Regina",
  "Manitoba Standard Time": "America/Winnipeg",
  "SA Pacific Standard Time": "America/Bogota",
  "Eastern Standard Time (Mexico)": "America/Cancun",
  "Eastern Standard Time": "America/New_York",
  "Haiti Standard Time": "America/Port-au-Prince",
  "Cuba Standard Time": "America/Havana",
  "US Eastern Standard Time": "America/Indianapolis",
  "Turks And Caicos Standard Time": "America/Grand_Turk",
  "Paraguay Standard Time": "America/Asuncion",
  "Atlantic Standard Time": "America/Halifax",
  "Venezuela Standard Time": "America/Caracas",
  "Central Brazilian Standard Time": "America/Cuiaba",
  "SA Western Standard Time": "America/La_Paz",
  "Pacific SA Standard Time": "America/Santiago",
  "Newfoundland Standard Time": "America/St_Johns",
  "Tocantins Standard Time": "America/Araguaina",
  "E. South America Standard Time": "America/Sao_Paulo",
  "SA Eastern Standard Time": "America/Cayenne",
  "Argentina Standard Time": "America/Buenos_Aires",
  "Greenland Standard Time": "America/Godthab",
  "Montevideo Standard Time": "America/Montevideo",
  "Magallanes Standard Time": "America/Punta_Arenas",
  "Saint Pierre Standard Time": "America/Miquelon",
  "Bahia Standard Time": "America/Bahia",
  "UTC-02": "Etc/GMT+2",
  "Azores Standard Time": "Atlantic/Azores",
  "Cape Verde Standard Time": "Atlantic/Cape_Verde",
  "UTC": "Etc/UTC",
  "GMT Standard Time": "Europe/London",
  "Greenwich Standard Time": "Atlantic/Reykjavik",
  "Sao Tome Standard Time": "Africa/Sao_Tome",
  "Morocco Standard Time": "Africa/Casablanca",
  "W. Europe Standard Time": "Europe/Berlin",
  "Central Europe Standard Time": "Europe/Budapest",
  "Romance Standard Time": "Europe/Paris",
  "Central European Standard Time": "Europe/Warsaw",
  "W. Central Africa Standard Time": "Africa/Lagos",
  "Jordan Standard Time": "Asia/Amman",
  "GTB Standard Time": "Europe/Bucharest",
  "Middle East Standard Time": "Asia/Beirut",
  "Egypt Standard Time": "Africa/Cairo",
  "E. Europe Standard Time": "Europe/Chisinau",
  "Syria Standard Time": "Asia/Damascus",
  "West Bank Standard Time": "Asia/Hebron",
  "South Africa Standard Time": "Africa/Johannesburg",
  "FLE Standard Time": "Europe/Kiev",
  "Israel Standard Time": "Asia/Jerusalem",
  "South Sudan Standard Time": "Africa/Juba",
  "Kaliningrad Standard Time": "Europe/Kaliningrad",
  "Sudan Standard Time": "Africa/Khartoum",
  "Libya Standard Time": "Africa/Tripoli",
  "Namibia Standard Time": "Africa/Windhoek",
  "Arabic Standard Time": "Asia/Baghdad",
  "Turkey Standard Time": "Europe/Istanbul",
  "Arab Standard Time": "Asia/Riyadh",
  "Belarus Standard Time": "Europe/Minsk",
  "Russian Standard Time": "Europe/Moscow",
  "E. Africa Standard Time": "Africa/Nairobi",
  "Iran Standard Time": "Asia/Tehran",
  "Arabian Standard Time": "Asia/Dubai",
  "Astrakhan Standard Time": "Europe/Astrakhan",
  "Azerbaijan Standard Time": "Asia/Baku",
  "Russia Time Zone 3": "Europe/Samara",
  "Mauritius Standard Time": "Indian/Mauritius",
  "Saratov Standard Time": "Europe/Saratov",
  "Georgian Standard Time": "Asia/Tbilisi",
  "Volgograd Standard Time": "Europe/Volgograd",
  "Caucasus Standard Time": "Asia/Yerevan",
  "Afghanistan Standard Time": "Asia/Kabul",
  "West Asia Standard Time": "Asia/Tashkent",
  "Ekaterinburg Standard Time": "Asia/Yekaterinburg",
  "Pakistan Standard Time": "Asia/Karachi",
  "Qyzylorda Standard Time": "Asia/Qyzylorda",
  "India Standard Time": "Asia/Calcutta",
  "Sri Lanka Standard Time": "Asia/Colombo",
  "Nepal Standard Time": "Asia/Katmandu",
  "Central Asia Standard Time": "Asia/Bishkek",
  "Bangladesh Standard Time": "Asia/Dhaka",
  "Omsk Standard Time": "Asia/Omsk",
  "Myanmar Standard Time": "Asia/Rangoon",
  "SE Asia Standard Time": "Asia/Bangkok",
  "Altai Standard Time": "Asia/Barnaul",
  "W. Mongolia Standard Time": "Asia/Hovd",
  "North Asia Standard Time": "Asia/Krasnoyarsk",
  "N. Central Asia Standard Time": "Asia/Novosibirsk",
  "Tomsk Standard Time": "Asia/Tomsk",
  "China Standard Time": "Asia/Shanghai",
  "North Asia East Standard Time": "Asia/Irkutsk",
  "Singapore Standard Time": "Asia/Singapore",
  "W. Australia Standard Time": "Australia/Perth",
  "Taipei Standard Time": "Asia/Taipei",
  "Ulaanbaatar Standard Time": "Asia/Ulaanbaatar",
  "Aus Central W. Standard Time": "Australia/Eucla",
  "Transbaikal Standard Time": "Asia/Chita",
  "Tokyo Standard Time": "Asia/Tokyo",
  "North Korea Standard Time": "Asia/Pyongyang",
  "Korea Standard Time": "Asia/Seoul",
  "Yakutsk Standard Time": "Asia/Yakutsk",
  "Cen. Australia Standard Time": "Australia/Adelaide",
  "AUS Central Standard Time": "Australia/Darwin",
  "E. Australia Standard Time": "Australia/Brisbane",
  "AUS Eastern Standard Time": "Australia/Sydney",
  "West Pacific Standard Time": "Pacific/Port_Moresby",
  "Tasmania Standard Time": "Australia/Hobart",
  "Vladivostok Standard Time": "Asia/Vladivostok",
  "Lord Howe Standard Time": "Australia/Lord_Howe",
  "Bougainville Standard Time": "Pacific/Bougainville",
  "Russia Time Zone 10": "Asia/Srednekolymsk",
  "Magadan Standard Time": "Asia/Magadan",
  "Norfolk Standard Time": "Pacific/Norfolk",
  "Sakhalin Standard Time": "Asia/Sakhalin",
  "Central Pacific Standard Time": "Pacific/Guadalcanal",
  "Russia Time Zone 11": "Asia/Kamchatka",
  "New Zealand Standard Time": "Pacific/Auckland",
  "UTC+12": "Etc/GMT-12",
  "Fiji Standard Time": "Pacific/Fiji",
  "Chatham Islands Standard Time": "Pacific/Chatham",
  "UTC+13": "Etc/GMT-13",
  "Tonga Standard Time": "Pacific/Tongatapu",
  "Samoa Standard Time": "Pacific/Apia",
  "Line Islands Standard Time": "Pacific/Kiritimati",
};

export function isIanaTimeZone(zone: string): boolean {
  try {
    new Intl.DateTimeFormat('en-US', { timeZone: zone });
    return true;
  } catch {
    return false;
  }
}

/** An IANA zone for what Outlook reports, or null when it is neither a known Windows name nor IANA. */
export function toIanaTimeZone(zone: string | null | undefined): string | null {
  if (!zone) return null;
  const mapped = WINDOWS_TO_IANA[zone];
  if (mapped) return mapped;
  return isIanaTimeZone(zone) ? zone : null;
}
```

## Checklist

- [ ] `MailboxSettings.Read` requested, or `hostTimeZone` supplied
- [ ] `MICROSOFT_CALENDAR_DEFAULT_TIME_ZONE` set to the business's IANA zone
- [ ] Tests pass under at least two process time zones
