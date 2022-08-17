class Logging
  require "logger"

  def create_entry(ent, i)
    if i == 0
    logger = Logger.new('KTP_log.log', 10, 29024000)
    logger.info ent
    else
      logger = Logger.new('LAP_log.log', 10, 29024000)
      logger.info ent
    end
  end
end