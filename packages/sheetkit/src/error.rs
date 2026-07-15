use napi::bindgen_prelude::{AsyncTask, BigInt, JsObjectValue, ToNapiValue};
use napi::{Env, Error, Result, Status, Task};
use napi_derive::napi;
use sheetkit_core::resource::ResourceLimitError;

use crate::Workbook;

fn create_resource_limit_error(
    env: &Env,
    metadata: ResourceLimitError,
    message: String,
) -> Result<Error> {
    let mut object = env.create_error(Error::new(Status::GenericFailure, message))?;
    object.set_named_property("code", ResourceLimitError::CODE)?;
    object.set_named_property("operation", metadata.operation_code())?;
    object.set_named_property("category", metadata.category_code())?;
    object.set_named_property("dimension", metadata.dimension_code())?;
    object.set_named_property(
        "limit",
        BigInt {
            sign_bit: false,
            words: vec![metadata.limit()],
        },
    )?;
    object.set_named_property(
        "observed",
        BigInt {
            sign_bit: false,
            words: vec![metadata.observed()],
        },
    )?;
    object.set_named_property("stage", metadata.stage_code())?;

    let value = (&object).into_unknown(env)?;
    Ok(Error::from(value))
}

pub(crate) fn map_core_error(env: &Env, error: sheetkit_core::error::Error) -> Error {
    let message = error.to_string();
    match error.resource_metadata().copied() {
        Some(metadata) => create_resource_limit_error(env, metadata, message)
            .unwrap_or_else(|bridge_error| bridge_error),
        None => Error::from_reason(message),
    }
}

fn reject_task_error(
    env: Env,
    metadata: Option<ResourceLimitError>,
    error: Error,
) -> Result<Workbook> {
    match metadata {
        Some(metadata) => Err(create_resource_limit_error(
            &env,
            metadata,
            error.reason.clone(),
        )?),
        None => Err(error),
    }
}

pub(crate) struct OpenPathTask {
    path: String,
    options: sheetkit_core::workbook::OpenOptions,
    resource_error: Option<ResourceLimitError>,
}

impl OpenPathTask {
    pub(crate) fn new(path: String, options: sheetkit_core::workbook::OpenOptions) -> Self {
        Self {
            path,
            options,
            resource_error: None,
        }
    }

    pub(crate) fn into_async_task(self) -> AsyncTask<Self> {
        AsyncTask::new(self)
    }
}

#[napi]
impl Task for OpenPathTask {
    type Output = sheetkit_core::workbook::Workbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        sheetkit_core::workbook::Workbook::open_with_options(&self.path, &self.options).map_err(
            |error| {
                self.resource_error = error.resource_metadata().copied();
                Error::from_reason(error.to_string())
            },
        )
    }

    fn resolve(&mut self, _env: Env, inner: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook { inner })
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}

pub(crate) struct OpenBufferTask {
    data: Vec<u8>,
    options: sheetkit_core::workbook::OpenOptions,
    resource_error: Option<ResourceLimitError>,
}

impl OpenBufferTask {
    pub(crate) fn new(data: Vec<u8>, options: sheetkit_core::workbook::OpenOptions) -> Self {
        Self {
            data,
            options,
            resource_error: None,
        }
    }

    pub(crate) fn into_async_task(self) -> AsyncTask<Self> {
        AsyncTask::new(self)
    }
}

#[napi]
impl Task for OpenBufferTask {
    type Output = sheetkit_core::workbook::Workbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        sheetkit_core::workbook::Workbook::open_from_buffer_with_options(&self.data, &self.options)
            .map_err(|error| {
                self.resource_error = error.resource_metadata().copied();
                Error::from_reason(error.to_string())
            })
    }

    fn resolve(&mut self, _env: Env, inner: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook { inner })
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}

pub(crate) struct OpenPasswordTask {
    path: String,
    password: String,
    resource_error: Option<ResourceLimitError>,
}

impl OpenPasswordTask {
    pub(crate) fn new(path: String, password: String) -> Self {
        Self {
            path,
            password,
            resource_error: None,
        }
    }

    pub(crate) fn into_async_task(self) -> AsyncTask<Self> {
        AsyncTask::new(self)
    }
}

#[napi]
impl Task for OpenPasswordTask {
    type Output = sheetkit_core::workbook::Workbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        sheetkit_core::workbook::Workbook::open_with_password(&self.path, &self.password).map_err(
            |error| {
                self.resource_error = error.resource_metadata().copied();
                Error::from_reason(error.to_string())
            },
        )
    }

    fn resolve(&mut self, _env: Env, inner: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook { inner })
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}
