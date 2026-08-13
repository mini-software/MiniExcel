use chrono::{Duration, NaiveDate, NaiveDateTime, NaiveTime};
use serde::de::{Deserializer, Error as _};

pub use calamine::{
    deserialize_as_date_or_none, deserialize_as_datetime_or_none, deserialize_as_duration_or_none,
    deserialize_as_f64_or_none, deserialize_as_i64_or_none, deserialize_as_time_or_none,
};
pub use rust_xlsxwriter::utility::{
    serialize_datetime_to_excel, serialize_option_datetime_to_excel,
};

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
