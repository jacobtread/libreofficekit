# ===
# Docker file for sample reproducable builds with libreofficekit
# you can use this as a template for your own docker files that
# compile rust binaries which use libreofficekit
# ===

FROM debian:bookworm-slim

# Set environment variables to avoid interaction during installation
ENV DEBIAN_FRONTEND=noninteractive

# Add the Bookworm Backports repository
RUN echo "deb http://deb.debian.org/debian bookworm-backports main" > /etc/apt/sources.list.d/bookworm-backports.list \
    && apt-get update

# Update and install required packages
RUN apt-get -t bookworm-backports install -y libreofficekit-dev

# Update and install required packages
RUN apt-get install -y clang curl build-essential libssl-dev pkg-config ca-certificates && \
    apt-get clean && \
    rm -rf /var/lib/apt/lists/*

# Update CA certificates
RUN update-ca-certificates


# Install Rust
RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- -y

# Setup environment
ENV LO_INCLUDE_PATH=/usr/include/LibreOfficeKit
ENV PATH="/root/.cargo/bin:${PATH}"

WORKDIR /app

# Copy cargo manifest and lock
COPY Cargo.toml .
COPY Cargo.lock .

# Copy build script
COPY build.rs .

# Create library entry point
RUN mkdir src && echo "" >src/lib.rs

# Copy C wrapper
COPY src/wrapper.cpp ./src/wrapper.cpp
COPY src/wrapper.hpp ./src/wrapper.hpp

# Build with no actual code (Dependency precaching)
RUN cargo build --target x86_64-unknown-linux-gnu --release

# Copy and build actual code
COPY src src
RUN touch src/lib.rs
RUN cargo build --target x86_64-unknown-linux-gnu --release
