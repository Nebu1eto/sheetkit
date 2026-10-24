use std::sync::Arc;

use napi::bindgen_prelude::{AsyncTask, BigInt, Buffer, JsObjectValue, ToNapiValue};
use napi::{Env, Error, Result, Status, Task};
use napi_derive::napi;
use sheetkit_core::resource::ResourceLimitError;

use crate::Workbook;

/// Stack reserved for open threads: the 8 MiB default of the Node main
/// thread and of libuv pool threads on macOS and Linux.
const OPEN_THREAD_STACK_BYTES: usize = 8 * 1024 * 1024;

/// Run an async-task open on a scoped Rust thread and wait for it.
///
/// The same allocation-heavy parse measured about 10% slower when run
/// directly on a libuv pool thread than on the Node main thread or a Rust
/// thread. A panic resumes on the calling thread as before, and if no
/// thread can be spawned the open runs inline.
fn open_on_thread(
    open: impl FnOnce() -> sheetkit_core::error::Result<sheetkit_core::workbook::Workbook> + Send,
) -> sheetkit_core::error::Result<sheetkit_core::workbook::Workbook> {
    let pending = std::sync::Mutex::new(Some(open));
    let take = || {
        pending
            .lock()
            .unwrap_or_else(std::sync::PoisonError::into_inner)
            .take()
    };
    let spawned = std::thread::scope(|scope| {
        std::thread::Builder::new()
            .name("sheetkit-open".to_string())
            .stack_size(OPEN_THREAD_STACK_BYTES)
            .spawn_scoped(scope, || take().map(|open| open()))
            .map(|handle| {
                handle
                    .join()
                    .unwrap_or_else(|payload| std::panic::resume_unwind(payload))
            })
    });
    match spawned {
        Ok(Some(opened)) => opened,
        _ => match take() {
            Some(open) => open(),
            None => Err(sheetkit_core::error::Error::Internal(
                "open thread finished without running the open".to_string(),
            )),
        },
    }
}

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

fn reject_task_error<T>(env: Env, metadata: Option<ResourceLimitError>, error: Error) -> Result<T> {
    match metadata {
        Some(metadata) => Err(create_resource_limit_error(
            &env,
            metadata,
            error.reason.clone(),
        )?),
        None => Err(error),
    }
}

pub(crate) struct SavePathTask {
    workbook: Arc<sheetkit_core::workbook::Workbook>,
    path: String,
    limits: sheetkit_core::save::SaveLimits,
    options: sheetkit_core::save::SaveOptions,
    resource_error: Option<ResourceLimitError>,
}

impl SavePathTask {
    pub(crate) fn new(
        workbook: Arc<sheetkit_core::workbook::Workbook>,
        path: String,
        limits: sheetkit_core::save::SaveLimits,
        options: sheetkit_core::save::SaveOptions,
    ) -> Self {
        Self {
            workbook,
            path,
            limits,
            options,
            resource_error: None,
        }
    }

    pub(crate) fn into_async_task(self) -> AsyncTask<Self> {
        AsyncTask::new(self)
    }
}

#[napi]
impl Task for SavePathTask {
    type Output = ();
    type JsValue = ();

    fn compute(&mut self) -> Result<Self::Output> {
        self.workbook
            .save_with_options(&self.path, self.limits, self.options)
            .map_err(|error| {
                self.resource_error = error.resource_metadata().copied();
                Error::from_reason(error.to_string())
            })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}

pub(crate) struct WriteBufferTask {
    workbook: Arc<sheetkit_core::workbook::Workbook>,
    limits: sheetkit_core::save::SaveLimits,
    resource_error: Option<ResourceLimitError>,
}

impl WriteBufferTask {
    pub(crate) fn new(
        workbook: Arc<sheetkit_core::workbook::Workbook>,
        limits: sheetkit_core::save::SaveLimits,
    ) -> Self {
        Self {
            workbook,
            limits,
            resource_error: None,
        }
    }

    pub(crate) fn into_async_task(self) -> AsyncTask<Self> {
        AsyncTask::new(self)
    }
}

pub(crate) struct SavePasswordTask {
    workbook: Arc<sheetkit_core::workbook::Workbook>,
    path: String,
    password: String,
    limits: sheetkit_core::crypt::CryptLimits,
    options: sheetkit_core::save::SaveOptions,
    resource_error: Option<ResourceLimitError>,
}

impl SavePasswordTask {
    pub(crate) fn new(
        workbook: Arc<sheetkit_core::workbook::Workbook>,
        path: String,
        password: String,
        limits: sheetkit_core::crypt::CryptLimits,
        options: sheetkit_core::save::SaveOptions,
    ) -> Self {
        Self {
            workbook,
            path,
            password,
            limits,
            options,
            resource_error: None,
        }
    }

    pub(crate) fn into_async_task(self) -> AsyncTask<Self> {
        AsyncTask::new(self)
    }
}

#[napi]
impl Task for SavePasswordTask {
    type Output = ();
    type JsValue = ();

    fn compute(&mut self) -> Result<Self::Output> {
        self.workbook
            .save_with_password_with_options(
                &self.path,
                &self.password,
                self.limits,
                sheetkit_core::save::SaveLimits::default(),
                self.options,
            )
            .map_err(|error| {
                self.resource_error = error.resource_metadata().copied();
                Error::from_reason(error.to_string())
            })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(output)
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}

#[napi]
impl Task for WriteBufferTask {
    type Output = Vec<u8>;
    type JsValue = Buffer;

    fn compute(&mut self) -> Result<Self::Output> {
        self.workbook
            .save_to_buffer_with_limits(self.limits)
            .map_err(|error| {
                self.resource_error = error.resource_metadata().copied();
                Error::from_reason(error.to_string())
            })
    }

    fn resolve(&mut self, _env: Env, output: Self::Output) -> Result<Self::JsValue> {
        Ok(Buffer::from(output))
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
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
        let (path, options) = (&self.path, &self.options);
        open_on_thread(|| sheetkit_core::workbook::Workbook::open_with_options(path, options))
            .map_err(|error| {
                self.resource_error = error.resource_metadata().copied();
                Error::from_reason(error.to_string())
            })
    }

    fn resolve(&mut self, _env: Env, inner: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook::from_core(inner))
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
        let (data, options) = (&self.data, &self.options);
        open_on_thread(|| {
            sheetkit_core::workbook::Workbook::open_from_buffer_with_options(data, options)
        })
        .map_err(|error| {
            self.resource_error = error.resource_metadata().copied();
            Error::from_reason(error.to_string())
        })
    }

    fn resolve(&mut self, _env: Env, inner: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook::from_core(inner))
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}

pub(crate) struct OpenPasswordTask {
    path: String,
    password: String,
    options: sheetkit_core::workbook::OpenOptions,
    resource_error: Option<ResourceLimitError>,
}

impl OpenPasswordTask {
    pub(crate) fn new(
        path: String,
        password: String,
        options: sheetkit_core::workbook::OpenOptions,
    ) -> Self {
        Self {
            path,
            password,
            options,
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
        let (path, password, options) = (&self.path, &self.password, &self.options);
        open_on_thread(|| {
            sheetkit_core::workbook::Workbook::open_with_password_with_options(
                path, password, options,
            )
        })
        .map_err(|error| {
            self.resource_error = error.resource_metadata().copied();
            Error::from_reason(error.to_string())
        })
    }

    fn resolve(&mut self, _env: Env, inner: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook::from_core(inner))
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}

pub(crate) struct OpenBufferPasswordTask {
    data: Vec<u8>,
    password: String,
    options: sheetkit_core::workbook::OpenOptions,
    resource_error: Option<ResourceLimitError>,
}

impl OpenBufferPasswordTask {
    pub(crate) fn new(
        data: Vec<u8>,
        password: String,
        options: sheetkit_core::workbook::OpenOptions,
    ) -> Self {
        Self {
            data,
            password,
            options,
            resource_error: None,
        }
    }

    pub(crate) fn into_async_task(self) -> AsyncTask<Self> {
        AsyncTask::new(self)
    }
}

#[napi]
impl Task for OpenBufferPasswordTask {
    type Output = sheetkit_core::workbook::Workbook;
    type JsValue = Workbook;

    fn compute(&mut self) -> Result<Self::Output> {
        let (data, password, options) = (&self.data, &self.password, &self.options);
        open_on_thread(|| {
            sheetkit_core::workbook::Workbook::open_from_buffer_with_password_with_options(
                data, password, options,
            )
        })
        .map_err(|error| {
            self.resource_error = error.resource_metadata().copied();
            Error::from_reason(error.to_string())
        })
    }

    fn resolve(&mut self, _env: Env, inner: Self::Output) -> Result<Self::JsValue> {
        Ok(Workbook::from_core(inner))
    }

    fn reject(&mut self, env: Env, error: Error) -> Result<Self::JsValue> {
        reject_task_error(env, self.resource_error, error)
    }
}

#[cfg(test)]
mod tests {
    use super::open_on_thread;

    #[test]
    fn open_runs_on_a_named_thread_and_returns_its_result() {
        let caller = std::thread::current().id();
        let seen = std::sync::Mutex::new(None);
        let opened = open_on_thread(|| {
            let current = std::thread::current();
            *seen.lock().unwrap() = Some((current.id(), current.name().map(str::to_string)));
            Err(sheetkit_core::error::Error::SheetNotFound {
                name: "Missing".to_string(),
            })
        });
        assert!(matches!(
            opened,
            Err(sheetkit_core::error::Error::SheetNotFound { ref name }) if name == "Missing"
        ));
        let (thread, name) = seen.into_inner().unwrap().unwrap();
        assert_ne!(thread, caller);
        assert_eq!(name.as_deref(), Some("sheetkit-open"));
        assert!(open_on_thread(|| Ok(sheetkit_core::workbook::Workbook::new())).is_ok());
    }

    #[test]
    fn open_panic_resumes_on_the_calling_thread() {
        let payload = std::panic::catch_unwind(|| {
            let _ = open_on_thread(|| panic!("open panicked"));
        })
        .unwrap_err();
        assert_eq!(payload.downcast_ref::<&str>(), Some(&"open panicked"));
    }
}
