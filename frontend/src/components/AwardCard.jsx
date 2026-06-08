import { motion } from "framer-motion";
import { Heart, Share2 } from "lucide-react";
import { useState } from "react";

export default function AwardCard({ index }) {
  const [liked, setLiked] = useState(false);
  const [votes, setVotes] = useState(1000 + index * 50);

  const awards = [
    { title: "Fluid Glass Morphism", author: "Aurora Design", year: 2024, color: "from-blue-600 to-cyan-500" },
    { title: "Neural Network Viz", author: "Tech Artisans", year: 2024, color: "from-purple-600 to-pink-500" },
    { title: "Quantum Scroll Effects", author: "Future Labs", year: 2024, color: "from-orange-600 to-red-500" },
    { title: "3D Text Transformations", author: "Creative Studio", year: 2023, color: "from-green-600 to-emerald-500" },
    { title: "Holographic Buttons", author: "UI Masters", year: 2023, color: "from-indigo-600 to-purple-500" },
    { title: "Ambient Particles", author: "Motion Design", year: 2023, color: "from-pink-600 to-rose-500" },
  ];

  const award = awards[index % awards.length];

  const handleLike = () => {
    setLiked(!liked);
    setVotes(liked ? votes - 1 : votes + 1);
  };

  return (
    <motion.div
      initial={{ opacity: 0, y: 20 }}
      whileInView={{ opacity: 1, y: 0 }}
      transition={{ duration: 0.5, delay: index * 0.1 }}
      viewport={{ once: true }}
      whileHover={{ y: -8 }}
      className="group relative"
    >
      <div className="absolute inset-0 bg-gradient-to-r from-blue-600 to-purple-600 rounded-2xl blur-xl opacity-0 group-hover:opacity-40 transition-opacity duration-500" />

      <div className="relative bg-gradient-to-br from-slate-800 to-slate-900 rounded-2xl border border-slate-700 overflow-hidden">
        {/* Card header with gradient background */}
        <div className={`h-32 bg-gradient-to-br ${award.color} relative overflow-hidden`}>
          <motion.div
            animate={{
              backgroundPosition: ["0% 0%", "100% 100%", "0% 0%"],
            }}
            transition={{ duration: 15, repeat: Infinity }}
            className="absolute inset-0 opacity-20"
            style={{
              backgroundImage: "url('data:image/svg+xml,%3Csvg width='100' height='100' xmlns='http://www.w3.org/2000/svg'%3E%3Cpath d='M0,50 Q25,0 50,50 T100,50' fill='none' stroke='white' stroke-width='2'/%3E%3C/svg%3E')",
              backgroundSize: "100px 100px",
            }}
          />

          {/* Floating badge */}
          <motion.div
            animate={{ y: [0, -10, 0], rotate: [0, 5, 0] }}
            transition={{ duration: 4, repeat: Infinity }}
            className="absolute -top-8 -right-8 w-24 h-24 bg-white/10 rounded-full backdrop-blur-sm border border-white/20"
          />
        </div>

        {/* Card content */}
        <div className="p-6 relative z-10">
          <motion.span
            initial={{ opacity: 0 }}
            whileInView={{ opacity: 1 }}
            transition={{ delay: 0.3 }}
            className="inline-block px-3 py-1 bg-gradient-to-r from-blue-500/20 to-purple-500/20 border border-blue-500/30 rounded-full text-xs font-semibold text-blue-300 mb-3"
          >
            {award.year} Winner
          </motion.span>

          <h3 className="text-xl font-black mb-2 text-white group-hover:text-transparent group-hover:bg-gradient-to-r group-hover:from-blue-400 group-hover:to-purple-400 group-hover:bg-clip-text transition-all duration-300">
            {award.title}
          </h3>

          <p className="text-gray-400 text-sm mb-4 line-clamp-2">
            An exceptional example of design innovation pushing the boundaries of what CSS can achieve.
          </p>

          <div className="flex items-center justify-between pt-4 border-t border-slate-700">
            <div>
              <p className="text-gray-500 text-xs">By</p>
              <p className="font-semibold text-gray-200">{award.author}</p>
            </div>

            <div className="flex gap-3">
              <motion.button
                whileHover={{ scale: 1.2 }}
                whileTap={{ scale: 0.9 }}
                onClick={handleLike}
                className={`p-2 rounded-lg transition-all ${
                  liked
                    ? "bg-red-500/20 text-red-400"
                    : "bg-slate-700 text-gray-400 hover:bg-slate-600"
                }`}
              >
                <Heart size={18} fill={liked ? "currentColor" : "none"} />
              </motion.button>

              <motion.button
                whileHover={{ scale: 1.2 }}
                whileTap={{ scale: 0.9 }}
                className="p-2 rounded-lg bg-slate-700 text-gray-400 hover:bg-slate-600 transition-all"
              >
                <Share2 size={18} />
              </motion.button>
            </div>
          </div>

          <motion.div
            initial={{ scaleX: 0 }}
            whileInView={{ scaleX: 1 }}
            transition={{ delay: 0.5, duration: 0.6 }}
            className="mt-3 h-1 bg-gradient-to-r from-blue-500 to-purple-500 rounded-full origin-left"
          />

          <p className="text-center text-gray-400 text-xs mt-3">
            <span className="font-bold text-white">{votes}</span> votes
          </p>
        </div>
      </div>
    </motion.div>
  );
}
