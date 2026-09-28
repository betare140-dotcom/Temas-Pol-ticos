
                    resumen_modelo = resumen_entrenamiento(
                        df_entrenamiento
                    )
                    st.caption(
                        "Entrenamiento usado: "
                        + " · ".join(
                            f"{k}: {v}"
                            for k, v in resumen_modelo.items()
                        )
                    )

                    conteo_sent = Counter(
                        r.get("sentimiento", "")
                        for r in registros
                    )
                    st.write(
                        "Clasificación: "
                        + " · ".join(
                            f"{_titulo_sentimiento(k)}: {v}"
                            for k, v in conteo_sent.items()
                        )
                    )

                    st.info(
                        "La base inicial contiene 99 publicaciones "
                        "del caso de Gaby 'La Bonita' Sánchez. "
                        "Es una semilla: conviene revisar los casos "
                        "marcados como REVISAR y agregar ejemplos "
                        "de otros actores y temas."
                    )

                total_por_red = Counter(r["red"] for r in registros)
                st.success(f"Se extrajeron {len(registros)} menciones relacionadas con el tema.")
                st.write(
                    " · ".join(
                        f"{red.title() if red != 'X' else 'X'}: {total_por_red.get(red, 0)}"
                        for red in ["X", "FACEBOOK", "INSTAGRAM", "TIKTOK"]
                        if red in redes_seleccionadas
                    )
                )

                st.caption(
                    "Omitidas/depuradas: "
                    f"fuera de tema {stats['fuera_tema']}, "
                    f"cuentas excluidas {stats['cuenta_excluida']}, "
                    f"otras redes {stats['otras_redes']}, RT {stats['rt']}, "
                    f"duplicados por link {stats['dup_link']}, "
                    f"duplicados por texto {stats['dup_texto']}, "
                    f"sin texto {stats['sin_texto']}."
                )

                # Vista de control: permite comprobar qué entró y qué quedó fuera antes de descargar.
                tab_in, tab_out = st.tabs([
                    f"Incluidas ({len(registros)})",
                    f"Excluidas por tema/reglas ({len(excluidos)})",
                ])
                with tab_in:
                    vista_in = pd.DataFrame([
                        {
                            "Red": r["red"],
                            "Usuario": (f"@{r['handle']}" if r["handle"] else r["autor"]),
                            "Texto": r["texto"],
                            "Sentimiento": r.get("sentimiento", ""),
                            "Confianza": (
                                f"{r.get('confianza_sentimiento', 0):.1%}"
                                if analizar_sentimiento
                                else ""
                            ),
                            "Motivo": r.get("motivo_tema", ""),
                            "Link": r["link"],
                        }
                        for r in registros
                    ])
                    st.dataframe(vista_in, use_container_width=True, hide_index=True)
                with tab_out:
                    if excluidos:
                        vista_out = pd.DataFrame([
                            {
                                "Red": r["red"],
                                "Usuario": (f"@{r['handle']}" if r.get("handle") else r.get("autor")),
                                "Motivo de exclusión": r.get("motivo", ""),
                                "Texto": r.get("texto", ""),
                                "Link": r.get("link", ""),
                            }
                            for r in excluidos
                        ])
                        st.dataframe(vista_out, use_container_width=True, hide_index=True)
                    else:
                        st.info("No hubo publicaciones excluidas por el filtro temático.")

                faltan_handle = sum(
                    1 for r in registros if r["red"] in {"X", "INSTAGRAM", "TIKTOK"} and not r["handle"]
                )
                if faltan_handle:
                    st.info(
                        f"{faltan_handle} publicación(es) no permitieron identificar un @usuario con los datos disponibles. "
                        "Se conservó el nombre del autor/página cuando estaba disponible."
                    )

                if analizar_sentimiento:
                    st.download_button(
                        "📥 Descargar CSV para corregir y seguir entrenando",
                        data=csv_revision(registros, tema),
                        file_name="revision_sentimiento.csv",
                        mime="text/csv",
                        key="onclusive_revision_sentimiento",
                        help=(
                            "Corrige la columna etiqueta y vuelve a cargar "
                            "ese archivo como dataset adicional en una "
                            "ejecución futura."
                        ),
                    )

                base = re.sub(r"[^A-Za-z0-9ÁÉÍÓÚáéíóúÑñ_-]+", "_", tema).strip("_") or "tema"
                word = crear_word(registros, tema)
                html_bytes = crear_html(registros)
                txt_bytes = crear_txt(registros)

                st.download_button(
                    "📥 Descargar Word",
                    data=word,
                    file_name=f"Extraccion_{base}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                st.download_button(
                    "📥 Descargar HTML",
                    data=html_bytes,
                    file_name=f"Extraccion_{base}.html",
                    mime="text/html",
                )
                st.download_button(
                    "📥 Descargar TXT",
                    data=txt_bytes,
                    file_name=f"Extraccion_{base}.txt",
                    mime="text/plain",
                )
            except Exception as exc:
                st.error(f"Error al procesar el archivo: {exc}")

