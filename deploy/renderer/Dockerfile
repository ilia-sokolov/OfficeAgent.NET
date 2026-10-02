# OfficeAgent renderer sandbox: the reference image.
#
# Build context is deploy/renderer after publishing the worker:
#   dotnet publish deploy/renderer/worker -c Release -o deploy/renderer/worker/bin/publish
#   docker build --target hardened -t officeagent-renderer:reference deploy/renderer
#
# Every input is pinned. The base image is pinned by digest, and every package by exact
# version. When Debian supersedes a pinned security update the build fails rather than
# drifting; update the pins and rerun the reference tests, as a recorded change.
#
# The worker is published self-contained for linux-x64 and carries its own .NET 10 runtime.
# Microsoft publishes no .NET 10 image on Debian 12, and the LibreOffice and Poppler pins below
# are Debian 12 builds, so the base is plain Debian rather than a .NET runtime image.

FROM debian:bookworm-slim@sha256:3783cc01769c7b2b1b83a5c5ad96c815348e28ed7da68e2e3687004faa906251 AS base

ARG LIBREOFFICE_VERSION=4:7.4.7-1+deb12u14
ARG POPPLER_VERSION=22.12.0-2+deb12u3
RUN apt-get update \
 && apt-get install -y --no-install-recommends \
      libreoffice-core-nogui=${LIBREOFFICE_VERSION} \
      libreoffice-writer-nogui=${LIBREOFFICE_VERSION} \
      libreoffice-calc-nogui=${LIBREOFFICE_VERSION} \
      libreoffice-impress-nogui=${LIBREOFFICE_VERSION} \
      poppler-utils=${POPPLER_VERSION} \
      fonts-dejavu-core=2.37-6 \
      fonts-liberation2=2.1.5-1 \
      fontconfig=2.14.1-4 \
 && rm -rf /var/lib/apt/lists/* \
 && useradd --uid 10001 --no-create-home --home-dir /tmp --shell /usr/sbin/nologin renderer

COPY worker/bin/publish/ /opt/officeagent/worker/
RUN chmod 0755 /opt/officeagent/worker/officeagent-render-worker
ENV HOME=/tmp \
    DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1
WORKDIR /tmp
ENTRYPOINT ["/opt/officeagent/worker/officeagent-render-worker"]

# TEST ONLY: the control image for the hardening tests. Never deploy this stage.
FROM base AS permissive
COPY test/officeagent-permissive.xcd /usr/lib/libreoffice/share/registry/officeagent-hardening.xcd
USER 10001:10001

# The deployment image. It is the last stage, so a build without --target gets this one.
FROM base AS hardened
COPY officeagent-hardening.xcd /usr/lib/libreoffice/share/registry/officeagent-hardening.xcd
USER 10001:10001
