# Set Plugged In profile to never turn off display or sleep
powercfg /change monitor-timeout-ac 0
powercfg /change standby-timeout-ac 0

# Set On Battery profile to never turn off display or sleep
powercfg /change monitor-timeout-dc 0
powercfg /change standby-timeout-dc 0