use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use serde::Serializer;
use serde::de::{Deserializer, Error as _};

pub fn deserialize_date<'de, D>(deserializer: D) -> std::result::Result<NaiveDate, D::Error>
where
    D: Deserializer<'de>,
{
    calamine::deserialize_as_date_or_string(deserializer)?.map_err(D::Error::custom)
}

pub fn deserialize_datetime<'de, D>(deserializer: D) -> std::result::Result<NaiveDateTime, D::Error>
where
    D: Deserializer<'de>,
{
    calamine::deserialize_as_datetime_or_string(deserializer)?.map_err(D::Error::custom)
}

pub fn deserialize_time<'de, D>(deserializer: D) -> std::result::Result<NaiveTime, D::Error>
where
    D: Deserializer<'de>,
{
    calamine::deserialize_as_time_or_string(deserializer)?.map_err(D::Error::custom)
}

pub fn deserialize_duration<'de, D>(deserializer: D) -> std::result::Result<Duration, D::Error>
where
    D: Deserializer<'de>,
{
    calamine::deserialize_as_duration_or_string(deserializer)?.map_err(D::Error::custom)
}

pub fn serialize_date_to_excel<S>(
    value: &NaiveDate,
    serializer: S,
) -> std::result::Result<S::Ok, S::Error>
where
    S: Serializer,
{
    rust_xlsxwriter::utility::serialize_datetime_to_excel(value, serializer)
}

pub fn serialize_datetime_to_excel<S>(
    value: &NaiveDateTime,
    serializer: S,
) -> std::result::Result<S::Ok, S::Error>
where
    S: Serializer,
{
    rust_xlsxwriter::utility::serialize_datetime_to_excel(value, serializer)
}

pub fn serialize_time_to_excel<S>(
    value: &NaiveTime,
    serializer: S,
) -> std::result::Result<S::Ok, S::Error>
where
    S: Serializer,
{
    rust_xlsxwriter::utility::serialize_datetime_to_excel(value, serializer)
}
