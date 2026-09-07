#!/bin/bash
# Seeds this image's blank document templates into the invoking user's
# Templates folder, so GNOME Files' "New Document" submenu offers text,
# Markdown, Word, Excel and PowerPoint documents out of the box.
#
# Runs as a per-user systemd unit at every graphical login (see
# user-templates-install.service) rather than being shipped in /etc/skel,
# because /etc/skel only reaches accounts whose home directory this image's
# own tooling creates. FreeIPA domain users - the accounts this image
# exists for - get their homes from oddjob-mkhomedir at first login on
# whichever machine they log into, potentially long after the image was
# deployed, and homes on a networked filesystem may never be created
# locally at all. A login-time unit covers local and domain accounts
# identically, and needs no first-boot special casing.
#
# Templates the user has since deleted are not restored: each template is
# seeded once per user, recorded in a stamp file, and never looked at
# again. That makes repeated logins a no-op while still letting a later
# image add a new template and have it show up.
set -euo pipefail

source_dir="/usr/share/bluefin-freeipa/templates"
stamp_file="${XDG_DATA_HOME:-$HOME/.local/share}/bluefin-freeipa/templates-seeded"

[[ -d $source_dir ]] || exit 0

# xdg-user-dir resolves the localized Templates directory (Vorlagen,
# Modèles, ...) from the user's own user-dirs.dirs, which is what GNOME
# Files actually reads. It falls back to printing $HOME when the templates
# directory is disabled (XDG_TEMPLATES_DIR set to "$HOME" - a deliberate
# opt-out) or when the binary isn't installed; treat both as "nothing to
# seed" rather than dumping template files into the home directory itself.
templates_dir=""
if command -v xdg-user-dir > /dev/null; then
    templates_dir="$(xdg-user-dir TEMPLATES)"
fi
if [[ -z $templates_dir || $templates_dir == "$HOME" ]]; then
    exit 0
fi

mkdir -p "$templates_dir" "$(dirname "$stamp_file")"
touch "$stamp_file"

seeded_any=0
while IFS= read -r -d '' template; do
    name="$(basename "$template")"

    # Already handled for this user on an earlier login - including the
    # case where they deleted it afterwards, which is left alone.
    if grep -qxF -- "$name" "$stamp_file"; then
        continue
    fi

    # Never clobber a same-named file the user put there themselves; still
    # record it, so this only ever gets checked once per user per template.
    if [[ ! -e $templates_dir/$name ]]; then
        install -Dm644 "$template" "$templates_dir/$name"
        seeded_any=1
    fi

    printf '%s\n' "$name" >> "$stamp_file"
done < <(find "$source_dir" -mindepth 1 -maxdepth 1 -type f -print0 | sort -z)

if ((seeded_any)); then
    echo "Seeded document templates into $templates_dir"
fi
