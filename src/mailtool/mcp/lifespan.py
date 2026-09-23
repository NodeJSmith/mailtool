"""Lifespan Management for MCP Server

This module provides Outlook bridge lifecycle management for the MCP server.
It handles:
- Creating and warming up OutlookBridge instance on startup
- Setting module-level bridge state for tool access
- Releasing COM objects and forcing garbage collection on shutdown
"""

import asyncio
import gc
import logging
from contextlib import asynccontextmanager

import pythoncom

from mailtool.bridge import OutlookBridge

# Configure logging for the lifespan manager
# Logs are written to stderr for debugging and monitoring
logger = logging.getLogger(__name__)


@asynccontextmanager
async def outlook_lifespan(app, default_account: str | None = None):
    """Async context manager for Outlook bridge lifecycle

    This function manages the complete lifecycle of the Outlook COM bridge:
    1. Creates OutlookBridge instance on startup (main thread)
    2. Warms up the connection with retry attempts
    3. Sets module-level bridge state for tool access
    4. Cleans up COM objects on shutdown

    CRITICAL: COM objects must be created and accessed from the same thread.
    The bridge is created directly in the main async context, not in an executor,
    to ensure thread affinity for all COM calls.

    Args:
        app: The FastMCP server instance (used to access module state)
        default_account: Optional default account name/email to use for folder operations
                        If None, will read from global _default_account in server module

    Yields:
        None: The bridge is set in server._bridge module state

    Raises:
        Exception: If Outlook cannot be connected to after retry attempts
    """
    import traceback as tb

    bridge = None
    com_initialized = False
    try:
        # If default_account not provided, read from global variable in server module
        if default_account is None:
            from mailtool.mcp import server as server_module

            default_account = getattr(server_module, "_default_account", None)
            if default_account:
                logger.info(
                    f"Read default_account from server module: {default_account}"
                )
        elif default_account:
            logger.info(f"Using provided default_account parameter: {default_account}")

        # Log lifespan start (using stderr to ensure visibility)
        logger.error("=" * 60)
        logger.error("LIFESPAN: Starting Outlook bridge initialization")
        logger.error("=" * 60)

        # Initialize COM in the main thread BEFORE creating the bridge
        # This ensures all subsequent COM calls happen from the same thread
        logger.info("Initializing COM in main thread...")
        pythoncom.CoInitialize()
        com_initialized = True
        logger.debug("COM initialized successfully")

        # Create Outlook bridge instance directly in main thread
        # NOT in executor - COM objects require thread affinity
        logger.info("Creating Outlook bridge...")
        if default_account:
            logger.info(f"Using default account: {default_account}")
            bridge = OutlookBridge(default_account=default_account)
        else:
            bridge = OutlookBridge()
        logger.info("Outlook bridge created successfully")

        # Warmup: Test that COM is responsive with retries
        max_retries = 5
        retry_delay = 0.5  # seconds

        for attempt in range(1, max_retries + 1):
            try:
                logger.debug(f"Warmup attempt {attempt}/{max_retries}")
                # Run a real COM call to ensure Outlook is responsive
                _warmup_bridge(bridge)
                logger.info("Outlook bridge warmed up successfully")
                break  # Success - exit retry loop
            except Exception as e:
                logger.warning(f"Warmup attempt {attempt}/{max_retries} failed: {e}")
                logger.error(
                    f"Exception traceback:\n{''.join(tb.format_exception(type(e), e, e.__traceback__))}"
                )
                if attempt == max_retries:
                    logger.error(
                        f"Outlook warmup failed after {max_retries} attempts: {e}"
                    )
                    logger.error(f"Full traceback:\n{''.join(tb.format_exc())}")
                    raise Exception(
                        f"Outlook warmup failed after {max_retries} attempts: {e}"
                    ) from e
                # Wait before retry
                await asyncio.sleep(retry_delay)

        # Set module-level bridge state for tools to access
        # Import here to avoid circular imports
        import mailtool.mcp.server as server_module

        # Use setattr to ensure we're setting the module-level variable correctly
        # This modifies the module's __dict__ directly to ensure _get_bridge() sees it
        server_module._bridge = bridge
        logger.info(f"Bridge set via setattr: {server_module._bridge is not None}")

        # Set bridge in resources module for resource access
        from mailtool.mcp import resources

        resources._set_bridge(bridge)
        logger.info("Outlook bridge initialized and ready")
        logger.error("=" * 60)
        logger.error("LIFESPAN: Bridge initialization complete, yielding to server")
        logger.error("=" * 60)

        # Yield for server to start
        yield

    except Exception as e:
        # Log any unexpected errors during lifespan startup
        logger.error("=" * 60)
        logger.error("LIFESPAN: Unexpected error during startup")
        logger.error(f"Error: {type(e).__name__}: {e}")
        logger.error(f"Full traceback:\n{''.join(tb.format_exc())}")
        logger.error("=" * 60)
        raise
    finally:
        # Cleanup: Release COM objects, uninitialize COM, and force garbage collection
        logger.info("Shutting down Outlook bridge...")
        if bridge is not None:
            try:
                # Release COM references
                bridge.outlook = None
                bridge.namespace = None
                logger.debug("Released COM references")
            except Exception as e:
                logger.error(f"Error releasing COM references: {e}")

        # Uninitialize COM for this thread (only if we initialized it)
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
                logger.debug("Uninitialized COM")
            except Exception as e:
                logger.error(f"Error uninitializing COM: {e}")

        # Force Python garbage collection to release COM objects
        gc.collect()
        logger.info("Outlook bridge shutdown complete")


def _warmup_bridge(bridge: OutlookBridge) -> None:
    """Synchronous warmup function to test COM connectivity

    Args:
        bridge: The OutlookBridge instance to test

    Raises:
        Exception: If COM call fails (Outlook not responsive)
    """
    logger.debug("Testing COM connectivity via warmup")
    inbox = bridge.get_inbox()
    # Make a real COM call to test connectivity
    count = inbox.Items.Count
    logger.debug(f"Warmup successful: Inbox has {count} items")
