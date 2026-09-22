# OfficeAgent renderer sandbox: the reference image.
#
# Build context is deploy/renderer after publishing the worker:
#   dotnet publish deploy/renderer/worker -c Release -o deploy/renderer/worker/bin/publish
#   docker build --target hardened -t officeagent-renderer:reference deploy/renderer
#
# Every input is pinned. The base image is pinned by digest, and every package by exact
# version. When Debian supersedes a pinned security update the build fails rather than
# drifting; update the pins and rerun the reference tests, as a recorded change.

FROM mcr.microsoft.com/dotnet/runtime:8.0-bookworm-slim@sha256:37466ea190f696105c1c3ae67c15e32d4e199face9a0b2ad5b9a37c464db8f30 AS base

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
ENV HOME=/tmp \
    DOTNET_CLI_TELEMETRY_OPTOUT=1 \
    DOTNET_NOLOGO=1
WORKDIR /tmp
ENTRYPOINT ["dotnet", "/opt/officeagent/worker/officeagent-render-worker.dll"]

# TEST ONLY: the control image for the hardening tests. Never deploy this stage.
FROM base AS permissive
COPY test/officeagent-permissive.xcd /usr/lib/libreoffice/share/registry/officeagent-hardening.xcd
USER 10001:10001

# The deployment image. It is the last stage, so a build without --target gets this one.
FROM base AS hardened
COPY officeagent-hardening.xcd /usr/lib/libreoffice/share/registry/officeagent-hardening.xcd
USER 10001:10001
